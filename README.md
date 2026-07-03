# 项目费用结算 v2

影视/平面广告制片人专用项目记账工具。

[⬇ 一键下载](https://github.com/luijett/project-cost-sheet/archive/refs/heads/master.zip)

## 安装运行

```bash
# 1. 确保已安装 Python 3.12+
python --version

# 2. 安装依赖
pip install PySide6 pydantic openpyxl pandas reportlab

# 3. 启动
python main.py
```

## 功能

- 📷🎬 平面/视频/套拍项目预算管理
- 预设行业分类模板（场地费、模特费、美术、器材、后期...）
- 💱 多币种 + 税率自动计算
- 📥 粘贴导入（从 PDF/Excel/微信 复制费用明细）
- 📥 Excel 导入（自动匹配列）
- 📤 导出 Excel / PDF
- ↩  Ctrl+Z 撤销
- 🌙 深色模式 UI

## 技术栈

Python 3.12 / PySide6 / SQLite
