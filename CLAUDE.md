# feishu-bot

飞书群聊机器人：群里发 Excel 链接 → 自动下载 → 解析影片排片数据 → 生成报告文案 → 回复到群里。

## 技术栈
- Python 3.12, lark-oapi（飞书 SDK WebSocket 长连接）
- openpyxl 读 Excel, python-dotenv 管理密钥

## 入口
`python main.py`（或双击 `start.bat`）

## 关键约定
- 两个 Excel 一组（同电影两个监控时段），缺一则静默跳过
- 文件名格式必须匹配 `parse_filename()` 的正则
- 去重持久化在 `.seen_msg_ids`，重启不丢失
- 对比模式：当表头出现两次"场次数"列时自动启用
- 不要在代码中硬编码密钥，统一走 `.env`
