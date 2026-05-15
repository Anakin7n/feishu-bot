# feishu-bot

飞书群聊机器人：群里发 Excel 链接 → 自动下载 → 解析影片排片数据 → 生成报告文案 → 回复到群里。

## 环境

- Python 3.12 装在 `%LocalAppData%\Programs\Python\Python312\`，已在系统 PATH
- 使用虚拟环境 `.venv`，依赖与系统 Python 隔离

### 首次配置

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

## 入口

| 方式 | 文件 | 说明 |
|------|------|------|
| 手动启动 | `start.bat` | Windows Terminal + PowerShell，有窗口可看日志 |
| 自启 | `start_auto.bat` | 任务计划调用，无窗口静默运行 |
| 直接运行 | `.\.venv\Scripts\python.exe main.py` | 使用虚拟环境 Python |

## 部署

- **开机自启**：Windows 任务计划 `FeishuBot`，触发条件为当前用户登录时，运行 `start_silent.vbs` → `start_auto.bat`
- **日志**：`bot.log`（追加写入，不会自动清理）
- **停止**：任务管理器结束 python.exe，或 `taskkill /F /IM python.exe`

## 关键约定

- 两个 Excel 一组（同电影两个监控时段，按 `mon_start` 排序），缺一则返回错误提示
- 文件名格式必须匹配 `parse_filename()` 的正则
- 消息去重持久化在 `.seen_msg_ids`，重启不丢失
- 对比模式：表头出现两次「场次数」列时自动启用
- 密钥走 `.env`，不硬编码
- URL 支持中文逗号、英文逗号（带或不带空格）、中文分号、换行分隔
