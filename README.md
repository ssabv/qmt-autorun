# QMT 自动登录工具

定时启动国金 QMT 并自动填写账号密码登录。

**注意：本项目不含任何账号密码信息，config.json 不会被提交。**

## 功能

- 图形界面配置账号、密码、程序路径
- 周一至周五 8:00 自动启动并登录
- 一键手动启动测试
- 打包后无需 Python 环境

## 运行

```bash
pip install -r requirements.txt
python main.py
```

## 打包为 exe

```bash
pip install pyinstaller
pyinstaller --onefile --noconsole main.py
```

## 配置文件

首次运行后会生成 `config.json`，包含你的账号密码信息。此文件不会上传到 Git。

## License

MIT
