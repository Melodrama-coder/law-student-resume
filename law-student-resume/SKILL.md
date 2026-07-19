---
name: law-student-resume
description: Create a polished, one-page Chinese law-student resume in Word, PDF, and optional LaTeX. Use when the user asks for help writing a 法学生简历, 法律求职简历, 红圈所/央企法务简历, 法学生求职简历, resume optimization for law students, or any task involving generating a standardized Chinese law-school resume. Supports 5 layout templates (classic/minimal/centered/modern/compact) with visual preview, black/white/gray/skyblue palette + custom accent, font choices (仿宋/微软雅黑/宋体/楷体) with Times New Roman, automatic one-page font-fitting, configurable language styles (aggressive/moderate/conservative), Chinese/English/bilingual output, multi-version generation for one-resume-multiple-submissions (law/legal/civil), interactive patch editing, overflow trimming suggestions, data-integrity validation with no-internship fallback, a phrase/example library, a job-fit radar chart, and breakpoint resume via JSON export/import.
---

# 法学生求职简历生成 Skill

帮助法学生通过**逐模块、口语友好**的结构化提问，整理真实经历，生成一页 A4 简历。支持 Word / PDF / LaTeX，中文/英文/中英双版，5 种版式模板及可视化预览，黑白灰/天蓝配色，自动字号适配，并提供数据校验、溢出精简建议、断点续传、交互式微调与岗位匹配雷达图。

## 核心交互原则（务必遵守）

1. **一次只问一个模块**。不要在一条消息里同时索取全部信息——法学生（尤其口语化的）会遗漏。按下方「模块顺序」逐个推进。
2. **口语友好**。明确告诉用户：“你只需用大白话说你做了什么，措辞我来帮你提炼。”不要求用户自己写成简历语言。
3. **每模块结束先回显小结**。用表格或要点复述你记录到的内容，请用户确认“对不对 / 补不补”，确认后再进入下一模块。
4. **边收集边追问细节**。每段经历都要深挖“具体做了什么 + 有无数据佐证”（见 `references/resume_guide.md`）。
5. **进度可见**。每步开头标明“第 X / N 步”，让用户知道还剩多少。
6. **诚实第一**。全程只做措辞提炼，绝不虚构单位、岗位、获奖、分数、数据；不确定的表述请用户确认或删除。
7. **断点续传**。每完成一个模块的回显小结，附上一段「当前汇总 JSON」（` ```json ` 代码块），请用户复制保存；下次粘贴该 JSON 即可续填，或用 `scripts/merge_resume.py` 合并多段碎片。
8. **必填校验**。生成前强制确认 姓名、学校、联系方式 非空；缺失则先追问，不生成空简历。无实习时主动提供「模拟课题/模拟法庭」替代方案（见 `references/resume_guide.md` §10）。
9. **语言风格**。第 0 步询问风格（`aggressive` 红圈所 / `moderate` 综合 / `conservative` 体制内），影响口语→书面的改写措辞（动词库见 `references/resume_guide.md` §11）。
10. **一稿多投**。收集完基础信息、确认真实性后，可询问是否生成**多版本**（律所版/法务版/体制版，`--versions law,legal`），不同版本自动重排模块顺序（见 §12）。
11. **中英双版**。第 0 步询问是否需要英文简历（`zh` / `en` / `both`）。英文内容存放在 `data.en` 块，章节标题与联系标签已自动英译，不确定术语参考 `references/legal_terms.md`。
12. **微调模式**。交付后若用户只想改个别字段，引导使用 `scripts/patch_resume.py`（见 `references/resume_guide.md` §8.4），改完重新生成即可。
13. **岗位匹配雷达图**。交付后可主动提供 `scripts/match_score.py` 岗位匹配诊断（红圈所/法务/外所 6 维雷达图 + 提升建议）。

## 模块顺序（逐个进行，不要合并）

> 开场先做「第 0 步」，然后每完成一个模块→回显小结→确认→进入下一个。

**第 0 步 · 版式、风格、配色与模块盘点**

先明确两件事，后面的收集才不会跑偏：

**① 必填 vs 可选模块（生成前强制校验）**
- **必填（缺任一项会中止生成并提示补齐）**：个人信息（姓名/学校/联系方式）、教育经历、实习/工作经历、专业技能。
- **可选（有才问，无则跳过）**：科研、竞赛、项目、个人总结。
- ⚠️ 若用户**完全没有实习/工作经历**：不编造，转而用「模拟法庭 / 课程课题 / 法律援助」替代（见 `references/resume_guide.md` §10），并提示这是替代素材。

**② 版式 / 风格 / 配色 / 语言（让用户可视化选择）**
- **先看模板预览**：打开 `assets/templates/preview.html`（同目录有 `classic/minimal/centered/modern/compact` 五张 PNG 缩略图），让用户一眼对比差异再选。
  - 红圈所/综合 → `classic`；央企法务/公检法 → `minimal` 或 `centered`（配仿宋）；新锐律所/互联网法务 → `modern`；信息密度高（offer6 风格）→ `compact`。
- **语言**：中文 `zh` / 英文 `en` / 中英双版 `both`（`--language`）。英文内容存在 `data.en` 块，未提供时自动回退中文并提示补充；术语参考 `references/legal_terms.md`。
- **语言风格**（影响改写措辞）：`aggressive` 红圈所（突出主导/结果）/ `moderate` 综合（平衡）/ `conservative` 体制内（严谨/合规）。见 §11 动词库。
- **配色**：默认 **黑、白、灰、天蓝**；可用 `--accent` / `meta.accent` 指定其他颜色（`navy`/`#C0392B` 等）。`compact` 还原参考图用 `accent: black`。
- **字体**：中文 仿宋/微软雅黑/宋体/楷体（默认微软雅黑；体制内推荐仿宋）；英文固定 Times New Roman。
- **一稿多投？** 询问是否要生成多版本（`--versions law,legal,civil`），不同版本自动重排模块顺序、突出不同侧重（见 §12）。
- **输出格式？** 默认 Word+PDF；若目标外所/学术岗，可问是否需要 LaTeX（`--format latex`）。

**③ 求职方向 + 模块清单**：问清求职方向（律所/央企法务/公检法/其他）与 TA 具体有哪些可选模块，没有的后面直接跳过。

**第 1 步 · 个人信息（必填）** — 姓名、电话、邮箱、出生年月/性别、籍贯、政治面貌、求职意向、证件照（可选，让用户自行在 Word 版添加或提供图片路径）。

**第 2 步 · 教育经历（必填）** — 逐段：学校、专业、学位层次、时间、GPA/均分、排名、核心课程、奖学金/荣誉；法考、语言成绩可在此或技能模块记录。

**第 3 步 · 实习/工作经历（必填，简历核心）** — **一段一段来**。每段：单位、部门+岗位、时间；然后深挖：
- 你独立/参与做了哪些**具体案件或项目**？（领域、类型）
- 产出了什么**成果**？（尽调报告、法律意见书、检索报告、合同审查、文书起草……）
- 有无**可量化数据**？（合同 X 份、检索 X 个课题/判例、案卷 X 卷、项目金额）
- 有无**正向反馈/转化**？（留用、被合伙人采纳、客户直接采用）
- 混合叙述要帮用户**拆分成独立段落**，并按倒序（最近在前）排列。

**第 4 步 · 可选模块（有才问，无则跳过）** — 科研经历、竞赛经历、项目经历、个人总结。逐个确认有无，有则同样深挖数据。

**第 5 步 · 专业技能（必填）** — 职业资格（法考 A/C 证、其他证书）、语言（雅思/托福/CET 分数、法律英语能力）、检索与办公工具（北大法宝、威科先行、把手案例、Westlaw/LexisNexis、企查查/天眼查、Office）。主动剔除用户明确不写的项。

**第 6 步 · 真实性确认（必须、单独一步）** — 逐条请用户确认“全部属实、约数未夸大、愿意负责”。确认后才生成。

## 模板与配色

配色原则：简历以专业、稳重为基调，默认推荐 **黑、白、灰、天蓝色**。**所有模板均支持自定义配色**，通过 `--accent` 或 `meta.accent` 指定任意颜色。

### 模板列表（推荐配色可自定义）

| 模板 | 推荐配色（默认） | 适配方向 | 配色说明 |
|------|------------------|----------|----------|
| `classic` | 天蓝色 `skyblue` | 红圈所 / 综合投递 | 经典稳重，标题栏与分隔线使用 accent |
| `minimal` | 黑色 `black` | 央企法务 / 公检法 | 极简黑白，灰色细线；accent 仅用于标题强调 |
| `centered` | 深灰色 `darkgray` | 体制内正式风（配仿宋） | 居中庄重，accent 用于姓名下划线 |
| `modern` | 天蓝色 `skyblue` | 新锐律所 / 互联网法务 | 现代感强，accent 用于色块与标题图标 |
| `compact` | 黑色 `black` | 信息密度高（offer6 风格） | 黑底白字标题栏，accent 作为标题底色；⚠️ 推荐深色 |

> **自定义配色示例**：
> ```bash
> # 使用预设名
> python scripts/generate_resume.py --data resume.json --template classic --accent navy
>
> # 使用十六进制颜色
> python scripts/generate_resume.py --data resume.json --template compact --accent #C0392B
>
> # 在 JSON 中指定
> # meta.accent: "teal" 或 meta.accent: "#2E7CC2"
> ```

### 支持的配色值

| 类型 | 示例 |
|------|------|
| 预设名（9 种） | `skyblue`、`black`、`gray`、`darkgray`、`navy`、`teal`、`burgundy`、`green`、`purple` |
| 自定义十六进制 | `#2E7CC2`、`#C0392B`、`#333333` |

> 📌 **提示**：`compact`（及 `modern` 色块头）若使用浅色 accent（如 `#EEEEEE`），脚本会**自动将标题/姓名文字反色为黑色**以保证可读性（基于感知亮度阈值）。推荐使用深色系（`black`/`navy`/`burgundy`/`darkgray`）。

建议先用当前数据生成用户选中的模板；若用户犹豫，可一次生成多种供对比（渲染成图见 `references/resume_guide.md`）。

## §13 配色建议与最佳实践

### 13.1 按求职方向推荐配色

| 求职方向 | 推荐配色 | 理由 |
|----------|----------|------|
| 红圈所 / 精品所 | `navy`、`skyblue`、`teal` | 体现专业、理性、国际化 |
| 央企法务 / 国企 | `black`、`darkgray`、`navy` | 体现稳重、规范、保守 |
| 公检法 / 体制内 | `black`、`darkgray` | 极简庄重，避免花哨 |
| 互联网大厂法务 | `purple`、`burgundy`、`teal` | 体现创新、现代感，仍保持专业 |
| 外所 / 涉外岗 | `navy`、`skyblue` | 与国际律所品牌色（蓝/藏青）一致 |

### 13.2 颜色可读性规则

脚本内置以下智能校验：
- **浅色背景 + 白色文字**：若 `accent` 感知亮度（加权 `0.299·R + 0.587·G + 0.114·B`）> 160，标题文字自动切换为黑色（由 `get_contrast_text_color()` 计算，见 §13.4）
- **深色背景 + 白色文字**：若 `accent` 亮度 ≤ 160，保持白色文字
- 用户也可用 `--force-color white|black` 强制覆盖自动反色（不推荐，除非自定义 accent 下反色不理想）

### 13.3 避免使用的颜色

- ❌ 荧光色（`#FF00FF`、`#00FF00`、`#FFFF00`）——不专业
- ❌ 过浅色（`#F5F5F5` 等）——与白底无区分
- ❌ 纯红色（`#FF0000`）——暗示错误/警示
- ❌ 过于花哨的多色混搭——简历应保持 1 种强调色

### 13.4 颜色解析函数（技术参考）

`scripts/generate_resume.py` 内置 `parse_color()` 函数：
- 支持 9 种预设名（见 §13.5）
- 支持 `#RRGGBB` 和 `#RGB` 格式
- 返回 RGB 元组 `(r, g, b)` 供 `python-docx` 使用
- 无效输入抛出 `ColorParseError`（由 `scripts/exceptions.py` 定义），并在命令行给出友好提示，不会抛出堆栈

### 13.5 预设颜色对照表

| 预设名 | 色值 | 预览 | 适用场景 |
|--------|------|------|----------|
| `skyblue` | `#87CEEB` | 🔵 天蓝 | 律所通用 |
| `navy` | `#000080` | ⚫ 藏青 | 红圈/外所 |
| `teal` | `#008080` | 🟢 墨绿 | 综合/创新 |
| `burgundy` | `#800020` | 🔴 酒红 | 精品所/互联网 |
| `purple` | `#800080` | 🟣 紫色 | 新锐/创意 |
| `green` | `#008000` | 🟢 森林绿 | 环保/公益法务 |
| `black` | `#000000` | ⚫ 纯黑 | 体制内/极简 |
| `gray` | `#808080` | ⚪ 灰色 | 低调/学术 |
| `darkgray` | `#404040` | ⬛ 深灰 | 央企/正式 |

> 提示：以上预设名即 `COLOR_PRESETS` 的键；自定义色值请用 `#RRGGBB`。若想要回之前更深的品牌蓝，可用 `--accent #2E7CC2`。

## 生成简历

1. 按 `references/resume_guide.md` 的 JSON 格式整理数据（含 `meta.font`、`meta.template`、`meta.accent`、`meta.style`、`meta.versions` 可选；见 §7）。
2. 调用脚本：
   ```bash
   # Word + PDF（默认）
   python scripts/generate_resume.py \
     --data path/to/resume.json \
     --out ./output \
     --name 姓名_简历 \
     [--font fang|yahei|song|kaiti] \
     [--template classic|minimal|centered|modern|compact] \
     [--accent skyblue|black|gray|navy|#RRGGBB] \
     [--style aggressive|moderate|conservative] \
     [--language zh|en|both] \
     [--versions law,legal,civil]      # 一稿多投：可省略

   # LaTeX 输出（外所/学术岗）
   python scripts/generate_resume.py \
     --data path/to/resume.json \
     --out ./output --name 姓名_简历 --format latex
   ```
3. 脚本先**强制校验必填项**（姓名/学校/联系方式/实习或替代素材），缺失则中止并提示；通过后自动从 **12pt 起**逐档尝试字号，选能单页容纳的最大字号。若最小字号仍超页，自动打印**两份精简建议（保守/激进）**供选择。
4. 生成后建议渲染 PDF 首页为 PNG 自检版面；英文版可请用户/母语者审阅。所有 PDF 已写入生成时间戳元数据。

## 持久化与断点续传（重要）

- **每次回显小结**都附一段「当前汇总 JSON」（` ```json ` 代码块），请用户**复制保存**。
- 下次会话用户粘贴该 JSON 即可续填；若有多段碎片，用：
  ```bash
  python scripts/merge_resume.py --base 已保存.json --patch 新片段.json --out 合并.json
  ```
  合并规则：标量字段新覆盖旧；数组（教育/实习等）按内容去重追加，不丢历史。

## 交付

- 交付 `.docx`（可继续编辑、插证件照）与 `.pdf`（投递用）。多版本时会分别交付 `_law` / `_legal` / `_civil` 后缀的文件；双语时会交付 `_zh` / `_en` 后缀的文件。
- 若用户要对比模板，可交付多套或一张拼版对比图；也可直接给 `assets/templates/preview.html` 链接。
- 生成失败通常是缺少 Microsoft Word 或 LibreOffice（用于 DOCX→PDF），或缺少 Python 依赖；脚本会友好提示 `pip install` 命令。

## 关键原则

- **诚实第一**：不编造任何信息；只做措辞提炼与合理约数（约数也须用户确认）。
- **一页原则**：所有内容压到一页 A4；过多时先删弱项、合并同类、精简 bullet。
- **量化优先**：用数字/结果/具体项目替代空泛形容词。
- **个性化**：按求职方向调整侧重，详见 `references/resume_guide.md`。

## 文件说明

- `scripts/generate_resume.py`：核心生成脚本（多模板 / 多语言 / 多版本 / 风格 / 校验 / 溢出精简 / LaTeX / PDF 时间戳 / 自定义配色 `--accent` / 强制文字色 `--force-color`），输入 JSON，输出单页 DOCX + PDF + 可选 .tex。日志写入 `resume.log`。
- `scripts/exceptions.py`：自定义异常 `ColorParseError` / `DataValidationError` / `TemplateNotFoundError`（配色或数据校验失败时抛出，命令行给出友好提示）。
- `scripts/__version__.py`：版本号 `__version__`（当前 `2.0`）。
- `scripts/patch_resume.py`：交互式微调——按路径定向修改 JSON（`--set` / `--add` / `--del` / `--patch`），改完重新生成即可。
- `scripts/merge_resume.py`：断点续传——合并「已保存 JSON 片段」与「新补充 JSON」（`--base / --patch / --out`）。
- `scripts/match_score.py`：岗位匹配度 6 维雷达图 + 文本诊断（红圈所 / 企业法务 / 外所）。
- `scripts/render_thumbnails.py`：渲染 5 种模板缩略图到 `assets/templates/`（用于预览）。
- `references/resume_guide.md`：模块化提问模板、口语转书面、真实性校验、一页裁剪决策树、风格引擎、多版本、JSON 格式、LaTeX/微调/雷达图用法。
- `references/examples.md`：案例库 / 话术库——10 个「口语→bullet」脱敏示例 + 各领域量化指标。
- `references/legal_terms.md`：法律简历常见中英术语对照表，供英文版翻译参考。
- `assets/templates/`：`classic/minimal/centered/modern/compact` 五张 PNG 缩略图 + `preview.html` 对比页。
- `assets/sample_resume.json`：可运行的匿名示例数据（含中英双块）。
