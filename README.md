# feishu-bot — 飞书全自动 Excel 处理机器人

## 概述

飞书群聊机器人，三种方式喂 Excel：① 群聊发链接（自动下载）② 群聊直接发文件 ③ 命令行传本地文件。解析影片排片数据 → 生成报告文案 → 回复到群里。WebSocket 长连接模式，无需公网 IP。

不依赖飞书官方 SDK（`lark-oapi`），使用原生 `websockets` + `requests` + 自写 protobuf 解析器，启动从 20+ 秒缩短到 1 秒以内。

## 文件结构

```
D:\feishu-bot\
├── main.py              # 机器人主程序（单文件）
├── requirements.txt     # Python 依赖
├── start.vbs            # 手动启动（双击弹出 PowerShell 蓝窗口，零闪屏）
├── start.ps1            # PowerShell 启动脚本
├── start.bat            # 手动启动（Windows Terminal，有闪屏，备用）
├── start_auto.bat       # 自启脚本（注册表 Run 键调用，无窗口）
├── start_silent.vbs     # VBS 静默包装器
├── install.bat           # 一键安装脚本（移植新设备用）
├── .venv/               # Python 虚拟环境（不提交 git）
├── .env                 # 飞书应用凭证（不提交 git）
├── .gitignore           # Git 忽略规则
├── .seen_msg_ids        # 消息去重缓存（自动生成）
├── bot.log              # 运行日志（自动轮转，5MB × 4 个文件）
└── README.md            # 本文件
```

所有代码自包含在一个目录中，无需外部文件。

## 环境要求

- Windows 10+
- Python 3.12+（安装时勾选 "Add to PATH"）
- 启动脚本自动检测：有 `.venv\` 用虚拟环境，没有则用系统 Python
- 飞书应用（需开启机器人能力，订阅 `im.message.receive_v1` 事件）

## 飞书应用配置

1. 飞书开放平台 → 创建应用 → 开启**机器人**能力
2. **权限管理**添加：
   - `im:message.group_msg` — 获取群组中所有消息（敏感权限）
   - `im:message:send_as_bot` — 以应用的身份发消息
   - `im:resource` — 获取与上传图片或文件资源
3. **事件订阅**添加 `im.message.receive_v1`（WebSocket 模式无需回调地址）
4. 发布应用并通过审核
5. 获取 App ID / App Secret 填入 `.env`

## .env 格式

```
FEISHU_APP_ID=cli_xxxxxxxx
FEISHU_APP_SECRET=xxxxxxxxxxxxxxxxxxxx
```

## 入口

启动脚本自动检测环境：有 `.venv\` 用虚拟环境，没有则用系统 Python。

| 方式 | 文件 | 说明 |
|------|------|------|
| 手动启动 | `start.vbs` | 双击弹出 PowerShell 蓝窗口 |
| 自启 | `start_auto.bat` | 注册表 Run 键调用，无窗口静默运行 |
| 直接运行 | `python main.py` | 当前目录下命令行启动 |
| CLI 测试 | `python main.py a.xlsx b.xlsx` | 直接传入本地文件，打印报告后退出 |

## 部署

- **开机自启**：注册表 `HKCU\Software\Microsoft\Windows\CurrentVersion\Run` 下 `FeishuBot` 键，值为 `start_silent.vbs` 完整路径。运行 `install.bat` 自动完成，或手动执行：
  ```powershell
  Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -Name "FeishuBot" -Value "D:\feishu-bot\start_silent.vbs"
  ```
- **取消自启**：
  ```powershell
  Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -Name "FeishuBot"
  ```
- **日志**：`bot.log`（自动轮转，超过 5MB 切分，保留最近 3 个历史文件）
- **停止**：任务管理器结束 python.exe，或 `taskkill /F /IM python.exe`

## 移植到新设备

1. 复制整个 `feishu-bot\` 文件夹到目标设备（不需要 `.venv`）
2. 安装 Python 3.12+（勾选 "Add Python to PATH"）
3. 双击 `install.bat`，自动完成：虚拟环境创建 → 依赖安装 → `.env` 模板生成 → 开机自启注册
4. 编辑 `.env` 填入飞书凭证
5. 双击 `start.vbs` 启动

> 如果不想用虚拟环境，可跳过第 3 步，直接 `pip install -r requirements.txt`。启动脚本会自动识别。

## 代码说明

### 消息流程

```
群聊消息 → on_message_receive()
             ├─ message_id 去重（.seen_msg_ids，线程锁 + fsync）
             ├─ 时间过滤（> 3 分钟跳过）
             ├─ 内容哈希去重（SHA-256，10 分钟 TTL）
             ├─ 过滤机器人自产消息
             │
             ├─ msg_type = "text"           msg_type = "file"
             │   提取 URL                    download_feishu_file()
             │   process_urls()              累积配对（等第二个文件，5s 超时提醒）
             │      download_file()           │
             │      _parse_single_file()  ←──┘
             │         parse_filename()
             │         find_data_sheet()
             │         find_header_row()
             │         extract_all_data()
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
| `download_feishu_file()` | 从飞书消息下载文件附件（调用资源 API），解析错误码返回具体提示 |
| `_parse_single_file()` | 解析单个 Excel 文件（文件名 + 内容），供 URL / 文件消息 / CLI 三路复用 |
| `_process_two_files()` | 两个文件到齐后合并排序，调用 build_message 生成一份文案 |
| `_on_file_timeout()` | 文件消息 5 秒未凑齐两个时提醒用户，文件继续暂存不丢弃 |
| `_is_duplicate()` | 消息去重（message_id），内存 set + 文件持久化，fsync 防关机丢失 |
| `_is_content_duplicate()` | 内容哈希去重（SHA-256），10 分钟 TTL，防 WS 重推不同 message_id 的相同消息 |
| `calc_deadline()` | 根据提取时间计算截止时间（上午→10点，下午→16点） |

### 文件名格式

```
哪吒之魔童闹海(2025-05-10+08:00-2025-05-10+12:00)2025-05-10+08:00.xlsx
└─ 影片名 ─┘└── 监控开始 ──┘└── 监控结束 ──┘└─ 提取日期 ─┘└时间┘
```

### 对比模式

表头出现两次「场次数」列时自动启用。从表头上方行提取对比影片名，提取对比影片的场次数、排片占比、排片占比差值等数据。

### 去重机制

四层防线，防止多台设备错时开机、WS 重连重推等场景下的重复回复：

| 层 | 机制 | 持久化 | 适用场景 |
|---|------|--------|---------|
| 1 | `message_id` 去重 | 文件（`.seen_msg_ids`，启动加载 + 即时 fsync） | 同一设备并发/重启回放 |
| 2 | 时间过滤（3 分钟） | 无 | 跨设备错时开机回放旧消息 |
| 3 | 内容哈希去重（SHA-256，10 分钟 TTL） | 内存 | WS 重连重推不同 message_id 的相同内容 |
| 4 | 机器人自产消息过滤 | 无 | 防止回复自己 |

关键实现细节：
- 第 1 层加 `threading.RLock` 防止并发竞态，`os.fsync` 保证关机不丢数据
- 第 2 层取 `msg.create_time`（毫秒时间戳）与当前时间比对
- 第 3 层 TTL 设为 10 分钟，避免误杀用户隔段时间重发的验证消息
- `.seen_msg_ids` 超过 500 条自动裁剪到 300 条

### 注意事项

- 两个 Excel 一组（同电影两个监控时段，按 `mon_start` 排序），缺一则返回错误提示
- 文件消息模式：发送第一个文件后暂存，等第二个文件到齐后合并处理；每 5 秒提醒一次，文件不丢失
- CLI 模式：`python main.py a.xlsx b.xlsx`，直接打报告到终端
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
