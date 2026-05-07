"""
飞书全自动 Excel 处理机器人（长连接模式）
群里发链接 → 自动下载 → 自动解析 → 自动回复
"""

import json
import os
import re
import sys
import time
import urllib.parse
from io import BytesIO
from pathlib import Path

import openpyxl
import requests
from dotenv import load_dotenv

from lark_oapi import Client as ApiClient, LogLevel
from lark_oapi.api.im.v1 import CreateMessageRequest, CreateMessageRequestBody, P2ImMessageReceiveV1
from lark_oapi.event.dispatcher_handler import EventDispatcherHandler
from lark_oapi.ws import Client as WsClient

load_dotenv()

APP_ID = os.getenv("FEISHU_APP_ID", "")
APP_SECRET = os.getenv("FEISHU_APP_SECRET", "")
DATA_SHEET = "综拓开场数据基础模板2"

# 备用 Sheet 名和自动探测
FALLBACK_SHEET_KEYWORDS = ["综拓", "开场数据", "基础模板"]

REQUIRED_COLUMNS = [
    "场次数",
    "劣势影城数",
    "排片占比",
]

OPTIONAL_COLUMNS = [
    "影片距离满足红包场次数差1场影城数",
    "影片距离满足红包场次数差2场影城数",
]

ALL_COLUMNS = REQUIRED_COLUMNS + OPTIONAL_COLUMNS

# ---- 文件名解析 ----

def parse_filename(filename: str) -> dict | None:
    """从文件名解析影片名、监控日期、提取时间。"""
    name = Path(filename).stem
    t = r"(\d{2})[:,_](\d{2})(?:[:,_](\d{2}))?"
    pattern = (
        r"^(.+?)"
        r"\((\d{4}-\d{2}-\d{2})[+ ]" + t
        + r"-(\d{4}-\d{2}-\d{2})[+ ]" + t + r"\)"
        r"(\d{4}-\d{2}-\d{2})[+ ](\d{2})[:,_](\d{2})"
        r"$"
    )
    m = re.match(pattern, name)
    if not m:
        return None

    mon_start_str = f"{m.group(2)} {m.group(3)}:{m.group(4)}:{m.group(5) or '00'}"
    mon_end_str = f"{m.group(6)} {m.group(7)}:{m.group(8)}:{m.group(9) or '00'}"

    return {
        "movie": m.group(1),
        "mon_start": mon_start_str,
        "mon_end": mon_end_str,
        "extract_date": m.group(10),
        "extract_hour": int(m.group(11)),
    }


# ---- Excel 数据提取 ----

def find_data_sheet(wb) -> str | None:
    """找到包含目标列头的数据 Sheet。"""
    # 优先精确匹配
    if DATA_SHEET in wb.sheetnames:
        return DATA_SHEET

    # 按关键词模糊匹配
    for name in wb.sheetnames:
        for kw in FALLBACK_SHEET_KEYWORDS:
            if kw in name:
                return name

    # 自动探测：查找包含「场次数」列头的 sheet
    for name in wb.sheetnames:
        ws = wb[name]
        for row_idx in range(1, min(ws.max_row + 1, 10)):
            for c in range(1, min(ws.max_column + 1, 30)):
                val = str(ws.cell(row=row_idx, column=c).value or "")
                if "场次数" in val:
                    return name

    return None


def find_header_row(ws) -> int | None:
    for row_idx in range(1, min(ws.max_row + 1, 15)):
        for c in range(1, ws.max_column + 1):
            if "场次数" in str(ws.cell(row=row_idx, column=c).value or ""):
                return row_idx
    return None


def find_heji_row(ws, start_row: int) -> int | None:
    for row_idx in range(start_row + 1, ws.max_row + 1):
        if str(ws.cell(row=row_idx, column=1).value or "").strip() == "合计":
            return row_idx
    for row_idx in range(start_row + 1, ws.max_row + 1):
        if "合计" in str(ws.cell(row=row_idx, column=1).value or "").strip():
            return row_idx
    return None


def extract_all_data(ws, header_row: int, main_movie: str = "") -> dict | None:
    """提取主电影和对比电影（如有）的数据。"""
    heji_row = find_heji_row(ws, header_row)
    if heji_row is None:
        print("  [警告] 未找到合计行")
        return None

    # 第一遍：扫描所有列，记录每列出现次数和位置
    col_occurrences: dict[str, list[int]] = {}  # col_name -> [col_idx1, col_idx2, ...]
    for col_idx in range(1, ws.max_column + 1):
        val = str(ws.cell(row=header_row, column=col_idx).value or "").strip()
        if val:
            if val not in col_occurrences:
                col_occurrences[val] = []
            col_occurrences[val].append(col_idx)

    # 检查是否是对比模式（场次数出现 >= 2 次）
    is_cmp = "场次数" in col_occurrences and len(col_occurrences["场次数"]) >= 2

    # 映射主电影列
    main_col_map = {}
    for col_name in ALL_COLUMNS:
        if col_name in col_occurrences:
            main_col_map[col_name] = col_occurrences[col_name][0]

    missing = [c for c in REQUIRED_COLUMNS if c not in main_col_map]
    if missing:
        print(f"  [警告] 缺少必要列: {missing}")
        return None

    main_data = {name: ws.cell(row=heji_row, column=col).value for name, col in main_col_map.items()}

    result = {"data": main_data, "cmp_movie": None, "cmp_data": None}

    if not is_cmp:
        return result

    # 提取对比电影名：从 header 上方行中查找
    first_cmp_col = col_occurrences["场次数"][1]
    cmp_movie = None
    for r in range(header_row - 1, max(0, header_row - 4), -1):
        val = str(ws.cell(row=r, column=first_cmp_col).value or "").strip()
        if val and val not in ("", "None") and val != main_movie and "场次" not in val:
            cmp_movie = val
            break

    # 如果没在第一个对比列上方找到，扫描整行
    if not cmp_movie:
        for r in range(header_row - 1, max(0, header_row - 4), -1):
            for c in range(first_cmp_col, ws.max_column + 1):
                val = str(ws.cell(row=r, column=c).value or "").strip()
                if val and len(val) > 1 and val != main_movie and "场次" not in val and "占比" not in val and "差值" not in val:
                    cmp_movie = val
                    break
            if cmp_movie:
                break

    if not cmp_movie:
        cmp_movie = "对比影片"

    print(f"  对比模式: {main_movie} vs {cmp_movie}")

    # 映射对比电影列（取第 2 次出现的通用列名 + 对比专用列名）
    cmp_col_map = {}

    # 对比列中也会出现的通用列名
    cmp_shared = ["场次数", "排片占比", "劣势影城数"]
    for name in cmp_shared:
        if name in col_occurrences and len(col_occurrences[name]) >= 2:
            cmp_col_map[name] = col_occurrences[name][1]

    # 对比专用列名
    cmp_specific = [
        "场次数新增",
        "主影片与当前影片场次差值",
        "主影片与当前影片排片占比差值",
        "拍片占比",
    ]
    for name in cmp_specific:
        if name in col_occurrences:
            cmp_col_map[name] = col_occurrences[name][0]

    cmp_data = {name: ws.cell(row=heji_row, column=col).value for name, col in cmp_col_map.items()}

    # 统一 拍片占比 → 排片占比
    if "拍片占比" in cmp_data and "排片占比" not in cmp_data:
        cmp_data["排片占比"] = cmp_data.pop("拍片占比")

    result["cmp_movie"] = cmp_movie
    result["cmp_data"] = cmp_data
    return result


# ---- 文案生成 ----

def calc_deadline(extract_date: str, extract_hour: int) -> str:
    deadline_hour = 10 if extract_hour < 12 else 16
    parts = extract_date.split("-")
    return f"{int(parts[1])}月{int(parts[2])}日{deadline_hour}点"


def format_mon_date(date_str: str) -> str:
    parts = date_str.split(" ")[0].split("-")
    return f"{int(parts[1])}月{int(parts[2])}日"


def format_mon_day(date_str: str) -> str:
    return f"{int(date_str.split(' ')[0].split('-')[2])}日"


def build_message(entries: list[dict]) -> tuple[str, str]:
    movie = entries[0]["movie"]
    last = entries[-1]
    deadline = calc_deadline(last["extract_date"], last["extract_hour"])
    date_range = f"{format_mon_date(entries[0]['mon_start'])}-{format_mon_date(entries[-1]['mon_start'])}"

    # 判断是否对比模式
    has_cmp = any(e.get("cmp_movie") for e in entries)
    cmp_movie = entries[0].get("cmp_movie", "") if has_cmp else ""

    # 根据是否有对比影城拼接不同的第1条文案
    if has_cmp:
        first_line = f"1）优先推进已开预售未开《{movie}》的影城，攻克劣势影城，加速影城开场进度"
    else:
        first_line = f"1）优先推进已开预售未开《{movie}》的影城，攻克劣势影城"

    lines = [
        "辛苦同步",
        first_line,
        "2）目标进度低于均值的小伙伴们继续加油",
        "",
        f"截止至{deadline}，【{movie}】{date_range}开场数据如上",
    ]

    def _format_pp(val):
        if isinstance(val, (int, float)) and val < 1:
            return round(val * 100, 2)
        return val

    def _hb_text(d):
        parts = []
        if "影片距离满足红包场次数差1场影城数" in d:
            parts.append(f"排片红包差值1场影院:{int(d['影片距离满足红包场次数差1场影城数'])}家")
        if "影片距离满足红包场次数差2场影城数" in d:
            parts.append(f"排片红包差值2场影院:{int(d['影片距离满足红包场次数差2场影城数'])}家")
        return "，".join(parts)

    n = 1
    for entry in entries:
        day_label = format_mon_day(entry["mon_start"])
        d = entry["data"]
        cc = int(d["场次数"])
        ls = int(d["劣势影城数"])
        pp = _format_pp(d["排片占比"])

        # 主电影行
        line = f"{n}）{day_label}《{movie}》已开{cc}场，未排《{movie}》的有{ls}家，排片占比{pp}%"
        hb = _hb_text(d)
        if hb:
            line += "，" + hb
        lines.append(line)
        n += 1

        # 对比电影行
        if has_cmp and entry.get("cmp_data"):
            cd = entry["cmp_data"]
            ccc = int(cd["场次数"])
            cpp = _format_pp(cd["排片占比"])
            diff = cd.get("主影片与当前影片排片占比差值")
            if diff is None and isinstance(d["排片占比"], (int, float)) and isinstance(cd["排片占比"], (int, float)):
                diff = round((_format_pp(d["排片占比"]) - _format_pp(cd["排片占比"])), 2)
            elif isinstance(diff, (int, float)) and diff < 1 and diff != 0:
                diff = round(diff * 100, 2)

            cmp_line = f"{n}）{day_label}《{cmp_movie}》已开{ccc}场，排片占比{cpp}%"
            if diff is not None:
                cmp_line += f"，{movie}与{cmp_movie}大盘差值为{diff}%"
            lines.append(cmp_line)
            n += 1

        lines.append("")

    main_message = "\n".join(lines)

    # 第二条消息：总结跟进语
    summary_message = f"以上为截止{deadline}，《{movie}》{date_range}开场情况，辛苦大家参考跟进。"

    return main_message, summary_message


# ---- 文件下载 ----

def download_file(url: str) -> tuple[str | None, bytes | None]:
    """下载 Excel，返回 (文件名, 内容)。"""
    try:
        resp = requests.get(url, timeout=60, headers={"Accept-Encoding": "identity"})
        resp.raise_for_status()

        fname = None
        cd = resp.headers.get("Content-Disposition", "")
        if "filename*=" in cd:
            parts = cd.split("filename*=")
            if len(parts) > 1:
                encoded = parts[1].split(";")[0].strip()
                if "''" in encoded:
                    _, fname_encoded = encoded.split("''", 1)
                    fname = urllib.parse.unquote(fname_encoded)
        if not fname and "filename=" in cd:
            fname = cd.split("filename=")[1].strip().strip('"')
        if not fname:
            fname = url.split("/")[-1].split("?")[0]

        return fname, resp.content
    except Exception as e:
        print(f"  [下载失败] {e}")
        return None, None


# ---- 消息发送 ----

def send_message(chat_id: str, text: str):
    api_client = ApiClient.builder() \
        .app_id(APP_ID) \
        .app_secret(APP_SECRET) \
        .log_level(LogLevel.ERROR) \
        .build()

    content = json.dumps({"text": text}, ensure_ascii=False)
    req = CreateMessageRequest.builder() \
        .receive_id_type("chat_id") \
        .request_body(CreateMessageRequestBody.builder()
            .receive_id(chat_id)
            .msg_type("text")
            .content(content)
            .build()) \
        .build()

    resp = api_client.im.v1.message.create(req)
    if resp.success():
        print("  已发送到群聊")
    else:
        print(f"  [发送失败] code={resp.code}, msg={resp.msg}")


# 消息去重（持久化到文件，重启不丢失）
_SEEN_FILE = Path(__file__).parent / ".seen_msg_ids"
_SEEN_MAX = 500  # 触发裁剪的阈值
_SEEN_KEEP = 300  # 裁剪时保留最近条数


def _load_seen() -> set[str]:
    """加载所有已处理的 msg_id。"""
    if _SEEN_FILE.exists():
        with open(_SEEN_FILE, "r") as f:
            return set(line.strip() for line in f if line.strip())
    return set()


def _trim_seen():
    """按文件追加顺序裁剪，保留最近 _SEEN_KEEP 条（避免 set 无序导致误删新条目）。"""
    if not _SEEN_FILE.exists():
        return
    with open(_SEEN_FILE, "r") as f:
        all_lines = [line.strip() for line in f if line.strip()]
    if len(all_lines) <= _SEEN_MAX:
        return
    with open(_SEEN_FILE, "w") as f:
        for line in all_lines[-_SEEN_KEEP:]:
            f.write(line + "\n")


def _is_duplicate(msg_id: str) -> bool:
    seen = _load_seen()
    if msg_id in seen:
        return True
    # 先持久化再裁剪（保证不丢当前 ID）
    with open(_SEEN_FILE, "a") as f:
        f.write(msg_id + "\n")
    seen.add(msg_id)
    if len(seen) > _SEEN_MAX:
        _trim_seen()
    return False


# ---- 事件处理 ----

def process_urls(urls: list[str]) -> tuple[str, str] | None:
    """下载、解析、生成文案。"""
    log_path = Path(__file__).parent / "bot.log"

    def log(msg: str):
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(msg + "\n")

    print("开始处理...", flush=True)
    entries = []

    for url in urls:
        log(f"  下载: {url}")
        fname, content = download_file(url)
        if not content:
            log(f"  [失败] 下载失败")
            continue

        safe_name = re.sub(r'[<>:"/\\|?*]', '_', fname)
        log(f"  文件: {safe_name}")

        info = parse_filename(safe_name)
        if info is None:
            log(f"  [警告] 文件名格式不匹配: {safe_name}")
            continue

        log(f"  影片: {info['movie']}  监控: {info['mon_start']} ~ {info['mon_end']}")

        wb = openpyxl.load_workbook(BytesIO(content), data_only=True)

        sheet_name = find_data_sheet(wb)
        if sheet_name is None:
            log(f"  [警告] 未找到数据 Sheet，可用: {wb.sheetnames}")
            wb.close()
            continue

        ws = wb[sheet_name]
        log(f"  Sheet: {sheet_name}  (max_row={ws.max_row}, max_col={ws.max_column})")
        header_row = find_header_row(ws)
        if header_row is None:
            log("  [警告] 未找到表头")
            wb.close()
            continue

        log(f"  表头行: {header_row}")

        all_data = extract_all_data(ws, header_row, info["movie"])
        wb.close()

        if all_data is None:
            log("  [失败] 数据提取失败")
            continue

        info["data"] = all_data["data"]
        if all_data["cmp_movie"]:
            info["cmp_movie"] = all_data["cmp_movie"]
            info["cmp_data"] = all_data["cmp_data"]

        log(f"  数据: {info['data']}")
        if info.get("cmp_movie"):
            log(f"  对比: {info['cmp_movie']} -> {info['cmp_data']}")
        entries.append(info)

    if len(entries) < 2:
        log(f"  至少需要 2 个有效文件，当前只有 {len(entries)} 个")
        return None

    entries.sort(key=lambda e: e["mon_start"])
    log("处理完成，生成文案")
    return build_message(entries)


def on_message_receive(event: P2ImMessageReceiveV1) -> None:
    """接收消息事件回调。"""
    msg = event.event.message
    msg_id = msg.message_id

    # 去重（持久化）
    if _is_duplicate(msg_id):
        return

    with open(Path(__file__).parent / "bot.log", "a", encoding="utf-8") as log:
        chat_id = msg.chat_id
        msg_type = msg.message_type

        log.write(f"[事件] chat_id={chat_id}  msg_type={msg_type}\n")

        if msg_type != "text":
            log.write(f"[跳过] 非文本消息: {msg_type}\n")
            return

        content_str = msg.content
        log.write(f"[内容] {content_str}\n")

        try:
            content_json = json.loads(content_str)
            text = content_json.get("text", "")
        except (json.JSONDecodeError, AttributeError):
            text = content_str

        if not text:
            log.write("[跳过] 空文本\n")
            return

        text = re.sub(r'@\S+\s*', '', text).strip()
        log.write(f"[文本] {text}\n")

        # 统一分隔符：中文标点 → 空格
        text = text.replace("；", " ").replace("，", " ").replace("\n", " ")
        urls = re.findall(r"https?://[^\s]+", text)
        if not urls:
            log.write("[跳过] 未发现链接\n")
            return

        log.write(f"收到 {len(urls)} 个链接: {urls}\n")

    print(f"收到 {len(urls)} 个链接，处理中...")
    result = process_urls(urls)
    if result:
        print("文案已生成")
        main_msg, summary_msg = result
        send_message(chat_id, main_msg)
        time.sleep(0.5)  # 避免飞书 API 连续调用限流
        send_message(chat_id, summary_msg)
    else:
        print("处理失败，发送错误提示")
        send_message(chat_id, "处理失败：请确认两个链接的 Excel 文件都包含正确的数据 Sheet")


# ---- 主入口 ----

def main():
    if not APP_ID or not APP_SECRET:
        print("请先在 .env 中配置 FEISHU_APP_ID 和 FEISHU_APP_SECRET")
        sys.exit(1)

    print(f"启动机器人 (App ID: {APP_ID})...")
    print("监听群聊消息中，可直接发送链接到群里测试\n")

    handler = EventDispatcherHandler.builder("", "") \
        .register_p2_im_message_receive_v1(on_message_receive) \
        .build()

    client = WsClient(
        app_id=APP_ID,
        app_secret=APP_SECRET,
        event_handler=handler,
        log_level=LogLevel.DEBUG,
    )
    client.start()


if __name__ == "__main__":
    main()
