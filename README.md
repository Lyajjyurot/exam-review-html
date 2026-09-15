# exam-review-html

把 PDF/Word 考试提纲转换成**单文件交互式 HTML 复习页**。
<img width="3840" height="2160" alt="23" src="https://github.com/user-attachments/assets/ddcf3e57-4087-401b-be11-ed88e71cef42" />


## 功能

-  按科目分标签切换，支持多科目
-  知识要点带批注与分级（核心必背/重点掌握/一般了解）
-  速背速记卡（口诀/对比表/公式卡，可遮罩自测）
-  模拟题在线作答即时判分，附参考答案与解析
-  易错点与考前自测清单
-  全局搜索、深浅色切换、进度本地存档、打印支持
-  产出物为单个自包含 HTML，不联网、不依赖 CDN

## 快速开始

### 1. 安装依赖

```bash
pip install pypdf python-docx
```

### 2. 提取文档内容

```bash
python scripts/extract_document.py <PDF/Word文件或目录> -o ./extracted --json
```

### 3. 生成复习内容 JSON

根据提取的内容，按照 `references/content-model.md` 的格式生成 `review_content.json`。

### 4. 构建 HTML

```bash
python scripts/build_review_html.py --content review_content.json --out 期末复习.html
```

## 项目结构

```
exam-review-html/
├── SKILL.md                  # 技能说明文档
├── assets/
│   ├── sample_content.json   # 示例内容 JSON
│   └── template.html         # HTML 模板
├── references/
│   ├── content-model.md      # 内容模型（JSON 字段规范）
│   └── pedagogy.md           # 批注·速记·命题方法论
└── scripts/
    ├── build_review_html.py  # 构建脚本
    └── extract_document.py   # 文档提取脚本
```

## 文档格式

- `.docx`：优先使用 `python-docx`，缺失时用标准库解析
- `.pdf`：需要安装 `pypdf`（或 `pdfplumber`/`PyMuPDF`）
- `.txt`/`.md`：原样读取

## 输出示例

生成的 HTML 包含：
- 顶部科目切换导航
- 知识要点卡片（带分级标签和批注）
- 速记卡（支持遮罩自测）
- 模拟题（选择题即时判分，主观题显示参考答案）
- 易错点清单和自测清单

## 许可证

[MIT License](LICENSE)
