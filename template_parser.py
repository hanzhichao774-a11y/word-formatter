import json
import re
from docx import Document
from openai import OpenAI
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from config import KIMI_API_KEY, KIMI_BASE_URL, KIMI_MODEL

FONT_SIZE_MAP = {
    "小初": 36, "一号": 26, "小一": 24,
    "二号": 22, "小二": 18, "三号": 16,
    "小三": 15, "四号": 14, "小四": 12,
    "五号": 10.5, "小五": 9,
}

ALIGNMENT_MAP = {
    "居中": WD_ALIGN_PARAGRAPH.CENTER,
    "左对齐": WD_ALIGN_PARAGRAPH.LEFT,
    "右对齐": WD_ALIGN_PARAGRAPH.RIGHT,
    "两端对齐": WD_ALIGN_PARAGRAPH.JUSTIFY,
}

SYSTEM_PROMPT = """你是一个排版规范解析器。用户会提供一份Word文档中的排版格式要求文本。
请从中提取以下7个元素的排版参数，返回严格的JSON格式。

需要提取的元素：title（题目）、heading1（一级标题）、heading2（二级标题）、heading3（三级标题）、body（正文）、caption（图表标题，如"图 1-1"、"表 2-3"的格式要求）、reference（参考文献条目）

每个元素需要提取以下字段：
- font_cn: 中文字体名（如"黑体"、"宋体"、"楷体"、"仿宋"、"微软雅黑"）
- font_en: 英文字体名（如"Times New Roman"、"Arial"），如未提及默认"Times New Roman"
- size: 字号名称（如"小二"、"小三"、"四号"、"小四"、"五号"）
- bold: 是否加粗（true/false）
- alignment: 对齐方式（"居中"、"左对齐"、"右对齐"、"两端对齐"）
- line_spacing: 行距，数字表示倍数（如1.5），或 "固定值20磅" 这样的描述
- space_before: 段前间距（磅数，如6、12、0）
- space_after: 段后间距（磅数，如6、12、0）
- first_line_indent: 首行缩进字符数（如2、0），如果是左边缩进也算
- hanging_indent: 悬挂缩进字符数（如2、0），仅reference可能用到

同时提取页面设置（如有）：
- margins: 页边距 { top, bottom, left, right }（厘米）

如果某个元素在文本中未提及，用null表示整个元素。
如果某个字段在文本中未提及，用null表示该字段。

只返回JSON，不要返回任何其他文字。格式如下：
{
  "title": { "font_cn": "黑体", "font_en": "Times New Roman", "size": "小二", "bold": true, "alignment": "居中", "line_spacing": 1.5, "space_before": 12, "space_after": 12, "first_line_indent": 0 },
  "heading1": { ... },
  "heading2": { ... },
  "heading3": { ... },
  "body": { ... },
  "caption": { ... },
  "reference": { ... },
  "margins": { "top": 2.54, "bottom": 2.54, "left": 3.17, "right": 3.17 }
}"""


def extract_text_from_docx(filepath):
    """Read all text from a .docx file."""
    doc = Document(filepath)
    lines = []
    for para in doc.paragraphs:
        text = para.text.strip()
        if text:
            lines.append(text)
    return "\n".join(lines)


def call_kimi(text):
    """Send text to Kimi API and get structured JSON back."""
    client = OpenAI(api_key=KIMI_API_KEY, base_url=KIMI_BASE_URL)
    response = client.chat.completions.create(
        model=KIMI_MODEL,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": f"请从以下排版格式要求中提取参数：\n\n{text}"},
        ],
        temperature=0.6,
        extra_body={"thinking": {"type": "disabled"}},
    )
    content = response.choices[0].message.content.strip()
    match = re.search(r"\{[\s\S]*\}", content)
    if match:
        return json.loads(match.group())
    return json.loads(content)


def _parse_line_spacing(value):
    """Convert line_spacing from API response to python-docx compatible value."""
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return value
    s = str(value)
    m = re.search(r"固定值?\s*(\d+(?:\.\d+)?)\s*磅", s)
    if m:
        return Pt(float(m.group(1)))
    m = re.search(r"(\d+(?:\.\d+)?)\s*倍", s)
    if m:
        return float(m.group(1))
    try:
        return float(s)
    except ValueError:
        return None


def build_template(parsed_json):
    """Convert Kimi's parsed JSON into a formatter-compatible template dict."""
    tpl = {
        "name": "自定义模板",
        "description": "根据上传的格式要求文档自动生成",
        "margins": {
            "top": Cm(2.54), "bottom": Cm(2.54),
            "left": Cm(3.17), "right": Cm(3.17),
        },
        "page_number": True,
        "header_footer_size": Pt(9),
    }

    margins_data = parsed_json.get("margins")
    if margins_data:
        for key in ("top", "bottom", "left", "right"):
            val = margins_data.get(key)
            if val is not None:
                tpl["margins"][key] = Cm(float(val))

    for role in ("title", "heading1", "heading2", "heading3", "body", "caption", "reference"):
        elem = parsed_json.get(role)
        if not elem:
            continue

        size_name = elem.get("size", "小四")
        size_pt = FONT_SIZE_MAP.get(size_name, 12)

        alignment_str = elem.get("alignment", "左对齐")
        alignment = ALIGNMENT_MAP.get(alignment_str, WD_ALIGN_PARAGRAPH.LEFT)

        line_spacing = _parse_line_spacing(elem.get("line_spacing", 1.5))

        indent = elem.get("first_line_indent")
        if indent is not None:
            indent = int(indent)
        elif role == "body":
            indent = 2
        else:
            indent = 0

        hanging = elem.get("hanging_indent")
        hanging = int(hanging) if hanging else 0

        space_before = elem.get("space_before", 0)
        space_after = elem.get("space_after", 0)

        role_cfg = {
            "font_cn": elem.get("font_cn", "宋体"),
            "font_en": elem.get("font_en", "Times New Roman"),
            "size": Pt(size_pt),
            "bold": bool(elem.get("bold", False)),
            "alignment": alignment,
            "line_spacing": line_spacing,
            "first_line_indent": indent,
            "space_before": Pt(float(space_before)) if space_before else Pt(0),
            "space_after": Pt(float(space_after)) if space_after else Pt(0),
        }
        if hanging > 0:
            role_cfg["hanging_indent"] = hanging
        tpl[role] = role_cfg

    if "table" not in tpl:
        body_cfg = tpl.get("body", {})
        tpl["table"] = {
            "font_cn": body_cfg.get("font_cn", "宋体"),
            "font_en": body_cfg.get("font_en", "Times New Roman"),
            "size": Pt(FONT_SIZE_MAP.get("五号", 10.5)),
            "three_line": True,
        }

    if "caption" not in tpl:
        body_cfg = tpl.get("body", {})
        tpl["caption"] = {
            "font_cn": body_cfg.get("font_cn", "宋体"),
            "font_en": body_cfg.get("font_en", "Times New Roman"),
            "size": Pt(FONT_SIZE_MAP.get("五号", 10.5)),
            "bold": False,
            "alignment": WD_ALIGN_PARAGRAPH.CENTER,
            "line_spacing": body_cfg.get("line_spacing", 1.5),
            "first_line_indent": 0,
            "space_before": Pt(6), "space_after": Pt(6),
        }

    if "reference" not in tpl:
        body_cfg = tpl.get("body", {})
        tpl["reference"] = {
            "font_cn": body_cfg.get("font_cn", "宋体"),
            "font_en": body_cfg.get("font_en", "Times New Roman"),
            "size": Pt(FONT_SIZE_MAP.get("五号", 10.5)),
            "bold": False,
            "alignment": WD_ALIGN_PARAGRAPH.JUSTIFY,
            "line_spacing": body_cfg.get("line_spacing", 1.25),
            "first_line_indent": 0,
            "hanging_indent": 2,
            "space_before": Pt(0), "space_after": Pt(0),
        }

    return tpl


def parse_template_from_docx(filepath):
    """Full pipeline: read docx → call Kimi → build template."""
    text = extract_text_from_docx(filepath)
    parsed = call_kimi(text)
    template = build_template(parsed)
    return template, parsed
