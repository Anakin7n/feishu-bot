请帮我在当前电脑上完成飞书 Bot 的部署。项目文件就在当前目录下。

### 你需要做的事情（按顺序）：

**1. 检查 Python 环境**
运行 `python --version`，确保 Python 3.9+ 已安装。如果没有，请提醒我安装，然后继续。

**2. 安装依赖**
```bash
pip install -r requirements.txt
```
如果 pip 命令不可用，尝试 `python -m pip install -r requirements.txt`。

**3. 验证 Bot 能否启动**
运行 `python main.py` 测试，看到 "监听群聊消息中" 就说明正常，然后 Ctrl+C 停止即可。

**4. 设置开机自启（可选）**
帮我创建一个 Windows 任务计划，让 Bot 开机自动在后台运行。可以直接用 start.bat（已改为相对路径，放到任何目录都能跑）。

### 注意事项
- `.env` 文件中的飞书应用凭证（APP_ID、APP_SECRET）已配置好，无需修改
- 如果公司网络有代理，可能需要在 `.env` 中额外配置 HTTP_PROXY
- 初次在公司电脑运行时，飞书开放平台可能会发送验证通知，留意飞书管理后台
