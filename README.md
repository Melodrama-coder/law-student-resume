# law-student-resume

法学生求职简历生成器 —— 一个用于 WorkBuddy 的 Skill，基于 Python + python-docx，一键生成专业、可投递的中文 / 中英双语法律简历。

## ✨ 功能特性
- **5 套模板**：minimal（极简）、classic（经典）、compact（紧凑）、centered（居中）、modern（现代），适配不同律所 / 法务岗审美
- **中英双语**：同一份数据可输出中文版与英文版
- **多方向版本**：内置「法学 / 法律」等方向措辞，自动匹配对应术语
- **自定义配色 + 智能对比色**：`--accent` 指定主题色（9 种预设或 #RRGGBB），浅色背景自动用黑字、深色背景自动用白字，保证可读性
- **友好报错**：颜色非法、数据缺失等情况给出中文提示，不抛堆栈
- **简历匹配度评分**（可选）：`match_score.py` 用雷达图直观展示简历与岗位的匹配维度

## 📦 安装
- 方式一（推荐）：下载 `law-student-resume.skill`，在 WorkBuddy 中导入安装。
- 方式二：克隆本仓库到 `~/.workbuddy/skills/law-student-resume/`（Windows：`C:\Users\<你>\.workbuddy\skills\law-student-resume\`）。

## 🚀 使用
准备一份 `resume.json` 数据，然后运行：

```bash
python scripts/generate_resume.py --data resume.json --template modern --lang zh
```

常用参数：
- `--template`：minimal / classic / compact / centered / modern
- `--lang`：zh / en
- `--accent`：主题色，如 `navy`、`#1a3a5c`
- `--force-color`：强制标题文字颜色 white / black
- `--output`：输出路径

查看全部参数：`python scripts/generate_resume.py --help`

## 📁 目录结构
```
law-student-resume/
├── SKILL.md                # Skill 说明（WorkBuddy 读取）
├── scripts/                # 生成脚本
│   ├── generate_resume.py  # 主生成器
│   ├── match_score.py      # 匹配度评分
│   └── ...
├── references/             # 法律术语 / 简历写作参考
└── assets/                 # 示例数据、模板预览图
```

## 🔧 依赖
- 必需：`python-docx`
- 可选（PDF / 评分）：`PyMuPDF`、`matplotlib`

```bash
pip install python-docx pymupdf matplotlib
```

## ⚠️ 免责声明
仓库内 `assets/sample_resume.json` 及预览图仅为**示例数据，纯属虚构**，请勿直接投递。生成真实简历时请填写本人真实信息。

## 📄 License
MIT
