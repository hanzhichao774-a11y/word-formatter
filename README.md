# 排版侠

Word 文档一键排版工具。上传 `.docx`，选择模板，自动完成正文格式化，封面、声明、目录等前置页面原样保留。

## 功能

- **三套内置模板**

| 模板 | 适用场景 | 正文字体 | 正文字号 | 行距 |
|------|----------|---------|---------|------|
| 通用论文 | 高校毕业论文 | 宋体 + Times New Roman | 小四 (12pt) | 1.5 倍 |
| 国标期刊 | 中文核心期刊投稿 | 宋体 + Times New Roman | 五号 (10.5pt) | 1.25 倍 |
| 简洁商务 | 公司报告、商务文档 | 微软雅黑 + Arial | 11pt | 1.5 倍 |

- **自动识别文档结构** — 基于 Word 段落样式 + 正则模式双重检测，区分标题、一/二/三级标题、正文
- **前置页面保护** — 智能定位目录位置，自动跳过封面、诚信声明、目录页，只格式化正文内容
- **三线表转换** — 表格自动应用三线表样式（论文/期刊模板）
- **页边距 & 页码** — 按模板规范设置页边距，页脚居中添加页码
- **中英文混排** — 通过 `w:eastAsia` 属性分别设置中文字体和英文字体

## 快速开始

```bash
# 克隆仓库
git clone https://github.com/hanzhichao774-a11y/word-formatter.git
cd word-formatter

# 安装依赖
pip install -r requirements.txt

# 启动
python app.py
```

浏览器打开 http://127.0.0.1:5000 ，上传文件即可使用。

## 项目结构

```
word-formatter/
├── app.py              # Flask 后端（上传 / 格式化 / 下载）
├── formatter.py        # 排版引擎核心
├── templates/
│   └── index.html      # 前端界面
├── requirements.txt    # Python 依赖
└── CHANGELOG.md        # 变更记录
```

## 排版引擎工作流程

```
上传 .docx
    │
    ▼
find_content_start()  ──→  定位"目录"标题
    │                       跳过目录条目（点线/页码识别）
    │                       匹配第一个正文标题模式
    ▼
前置区域 (封面/声明/目录)  ──→  完全跳过，不做任何修改
    │
    ▼
正文区域  ──→  detect_role() 识别每段角色
    │          ├── title    → 标题格式
    │          ├── heading1 → 一级标题格式
    │          ├── heading2 → 二级标题格式
    │          ├── heading3 → 三级标题格式
    │          └── body     → 正文格式
    ▼
表格处理  ──→  三线表转换（仅正文区域内的表格）
    │
    ▼
页面设置  ──→  页边距 + 页码
    │
    ▼
输出 .docx
```

## 支持的标题模式

引擎通过以下正则模式识别标题层级：

| 层级 | 匹配模式示例 |
|------|-------------|
| 一级标题 | `第一章 XXX`、`1 XXX`、`摘要`、`Abstract`、`参考文献`、`致谢`、`附录` |
| 二级标题 | `1.1 XXX`、`第一节 XXX` |
| 三级标题 | `1.1.1 XXX` |

同时也识别 Word 内置样式名（Heading 1 / Heading 2 / Heading 3 / Title）。

## 作为 Python 库使用

```python
from formatter import format_document

result = format_document(
    input_path="input.docx",
    output_path="output.docx",
    template_key="通用论文",  # 或 "国标期刊" / "简洁商务"
)

print(result)
# {
#     "template": "通用论文模板",
#     "changes": {"title": 0, "heading1": 5, "heading2": 8, "heading3": 2, "body": 30, "skipped": 27},
#     "total_paragraphs": 73,
#     "content_start": 27,
#     "total_tables": 3,
# }
```

## 技术栈

- **后端**: Python 3 + Flask
- **文档处理**: python-docx
- **前端**: 原生 HTML/CSS/JS（无框架依赖）

## License

MIT
