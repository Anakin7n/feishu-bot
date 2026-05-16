# feishu-bot — 飞书全自动 Excel 处理机器人

## 概述

飞书群聊机器人，群里发 Excel 链接 → 自动下载 → 解析影片排片数据 → 生成报告文案 → 回复到群里。WebSocket 长连接模式，无需公网 IP。

不依赖飞书官方 SDK（`lark-oapi`），使用原生 `websockets` + `requests` + 自写 protobuf 解析器，启动从 20+ 秒缩短到 1 秒以内。

## 文件结构

```
D:\feishu-bot\
├── main.py              # 机器人主程序（单文件，约 580 行）
├── requirements.txt     # Python 依赖
├── start.vbs            # 手动启动（双击弹出 PowerShell 蓝窗口，零闪屏）
├── start.ps1            # PowerShell 启动脚本
├── start.bat            # 手动启动（Windows Terminal，有闪屏，备用）
├── start_auto.bat       # 自启脚本（任务计划调用，无窗口）
├── start_silent.vbs     # VBS 静默包装器
├── .venv/               # Python 虚拟环境（不提交 git）
├── .env                 # 飞书应用凭证（不提交 git）
├── .gitignore           # Git 忽略规则
├── .seen_msg_ids        # 消息去重缓存（自动生成）
├── bot.log              # 运行日志（自动生成）
└── CLAUDE.md            # 本文件
```

所有代码自包含在一个目录中，无需外部文件。

## 环境要求

- Windows 10+
- Python 3.12（装在 `%LocalAppData%\Programs\Python\Python312\`，已在系统 PATH）
- 使用项目自带的虚拟环境（`.venv/`），不依赖系统 Python：
  ```
  python -m venv .venv
  .\.venv\Scripts\pip install -r requirements.txt
  ```
- 飞书应用（需开启机器人能力，订阅 `im.message.receive_v1` 事件）

## 飞书应用配置

1. 飞书开放平台 → 创建应用 → 开启**机器人**能力
2. **权限管理**添加：
   - `im:message` — 读取消息
   - `im:message:send_as_bot` — 发送消息
3. **事件订阅**添加 `im.message.receive_v1`（WebSocket 模式无需回调地址）
4. 发布应用并通过审核
5. 获取 App ID / App Secret 填入 `.env`

## .env 格式

```
FEISHU_APP_ID=cli_xxxxxxxx
FEISHU_APP_SECRET=xxxxxxxxxxxxxxxxxxxx
```

## 入口

| 方式 | 文件 | 说明 |
|------|------|------|
| 手动启动 | `start.vbs` | 双击弹出 PowerShell 蓝窗口，零闪屏 |
| 自启 | `start_auto.bat` | 任务计划调用，无窗口静默运行 |
| 备用启动 | `start.bat` | Windows Terminal + PowerShell（有闪屏，备用） |
| 直接运行 | `.\.venv\Scripts\python.exe main.py` | 使用虚拟环境 Python |

## 部署

- **开机自启**：Windows 任务计划 `FeishuBot`，触发条件为当前用户登录时，运行 `start_silent.vbs` → `start_auto.bat`
- **日志**：`bot.log`（追加写入，不会自动清理）
- **停止**：任务管理器结束 python.exe，或 `taskkill /F /IM python.exe`

## 移植到新设备

1. 复制整个 `feishu-bot\` 文件夹到新设备
2. 安装 Python 3.12+
3. 创建虚拟环境并安装依赖：
   ```
   python -m venv .venv
   .\.venv\Scripts\pip install -r requirements.txt
   ```
4. 填入飞书凭证到 `.env`
5. 双击 `start.vbs` 启动

## 代码说明

### 消息流程

```
群聊文本消息 → on_message_receive()
                ├─ 消息去重（.seen_msg_ids）
                ├─ 过滤非文本消息
                ├─ 提取 URL（支持中英文逗号、分号、换行分隔）
                │
                ▼
              process_urls()
                ├─ download_file() → 下载 Excel
                ├─ parse_filename() → 正则解析文件名
                ├─ find_data_sheet() → 智能定位数据 Sheet
                ├─ find_header_row() → 定位表头行
                ├─ extract_all_data() → 提取主电影 + 对比电影数据
                │
                ▼
              build_message() → 生成两条文案
                ├─ 第1条：详细数据报告
                └─ 第2条：总结跟进语
                │
                ▼
              send_message() × 2 → 飞书 API 回复到群聊
```

### 关键函数

| 函数 | 作用 |
|------|------|
| `_get_token()` | 获取并缓存 tenant_access_token（带过期控制） |
| `_feishu_post()` | 封装的飞书 HTTP POST 请求（自动带 token） |
| `send_message()` | 使用 requests 直接调用飞书发送消息 API |
| `FeishuWsClient` | 精简 WebSocket 客户端，替代 lark_oapi.ws.Client |
| `encode_ping_frame()` | 自写 protobuf 编码 ping 帧 |
| `decode_frame()` | 自写 protobuf 解码 WebSocket 帧 |
| `on_message_receive()` | 消息事件分发（接收 raw dict 替代 P2ImMessageReceiveV1） |
| `parse_filename()` | 从文件名正则提取影片名、监控时段、提取时间 |
| `find_data_sheet()` | 三级匹配定位数据 Sheet（精确→模糊→自动探测） |
| `find_header_row()` | 扫描前 15 行定位含「场次数」的表头行 |
| `find_heji_row()` | 定位「合计」行 |
| `extract_all_data()` | 提取主电影数据，自动检测对比模式 |
| `build_message()` | 生成详细数据报告 + 总结跟进语 |
| `download_file()` | 从 URL 下载 Excel，解析 Content-Disposition 获取文件名 |
| `_is_duplicate()` | 消息去重，内存 set + 文件持久化，重启不丢失 |
| `calc_deadline()` | 根据提取时间计算截止时间（上午→10点，下午→16点） |

### 文件名格式

```
哪吒之魔童闹海(2025-05-10+08:00-2025-05-10+12:00)2025-05-10+08:00.xlsx
└─ 影片名 ─┘└── 监控开始 ──┘└── 监控结束 ──┘└─ 提取日期 ─┘└时间┘
```

### 对比模式

表头出现两次「场次数」列时自动启用。从表头上方行提取对比影片名，提取对比影片的场次数、排片占比、排片占比差值等数据。

### 去重机制

启动时从 `.seen_msg_ids` 加载已处理 ID 到内存 set，每次新消息追加写入文件。超过 500 条自动裁剪到 300 条。

### 注意事项

- 两个 Excel 一组（同电影两个监控时段，按 `mon_start` 排序），缺一则返回错误提示
- 文件名格式必须严格匹配正则，否则跳过
- 密钥走 `.env`，不硬编码，不提交 git
- URL 支持中文逗号、英文逗号（带或不带空格）、中文分号、换行分隔
- 数据 Sheet 优先匹配「综拓开场数据基础模板2」，找不到则自动探测

### WebSocket 长连接架构

不使用 lark-oapi SDK，改用原生 `websockets` 库直连飞书 WebSocket 网关：

```
FeishuWsClient.start()
  └─ connect()          ← 无限重连循环
       └─ _try_connect()
            ├─ _get_ws_url()   → POST /callback/ws/endpoint 获取 WS 地址
            ├─ websockets.connect(url, proxy=None)   ← 禁用代理直连
            ├─ _ping_loop()    → 定时发送 protobuf 编码的 ping 帧
            └─ _read_loop()
                 ├─ recv() → decode_frame() → 帧类型判断
                 └─ type=1: asyncio.create_task(_process_event())
                              └─ loop.run_in_executor(线程池) → on_message_receive()
```

关键细节：
- WS 端点 URL 为 `https://open.feishu.cn/callback/ws/endpoint`
- 从 URL 参数 `service_id` 提取后用于 ping 帧编码
- `websockets.connect(url, proxy=None)` 强制禁用系统代理，避免代理引发的连接超时
- 事件处理用 `loop.run_in_executor` 丢到 `ThreadPoolExecutor(max_workers=2)` 避免阻塞事件循环

### Protobuf 编解码

自写简化版 protobuf 编解码器，替代 lark-oapi 内置的 google protobuf：
- `encode_ping_frame(service_id)` — 编码 ping 帧（Header: type=ping）
- `decode_frame(data)` — 解码 WebSocket 接收帧，支持 varint、length-delimited、嵌套 header 解析
- `frame_type(frame)` / `frame_data_payload(frame)` — 提取帧类型和载荷
