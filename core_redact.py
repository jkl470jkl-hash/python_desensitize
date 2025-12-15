import json
import os
import re
import datetime
from typing import Dict, List, Tuple, Optional, TYPE_CHECKING

if TYPE_CHECKING:
    from docx import Document

CONFIG_DIR = os.path.join(os.path.dirname(__file__), "config")
OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "output")


WORDLIST_DEFAULTS = {
    "company_suffix.txt": [
        "有限公司",
        "有限责任公司",
        "股份有限公司",
        "集团有限公司",
        "集团",
        "公司",
        "事务所",
        "研究院",
        "中心",
        "基金会",
    ],
    "role_words.txt": [
        "甲方",
        "乙方",
        "丙方",
        "丁方",
        "买方",
        "卖方",
        "承包人",
        "发包人",
        "委托方",
        "受托方",
        "出租方",
        "承租方",
        "服务方",
        "客户方",
    ],
    "address_field_words.txt": [
        "地址",
        "联系地址",
        "注册地址",
        "办公地址",
        "经营地址",
        "通讯地址",
        "住所地",
    ],
    "address_component_words.txt": [
        "省",
        "市",
        "区",
        "县",
        "镇",
        "乡",
        "路",
        "街",
        "巷",
        "号",
        "栋",
        "幢",
        "楼",
        "室",
        "大厦",
        "园区",
        "小区",
        "工业园",
        "科技园",
    ],
    "person_family_names.txt": [
        "赵",
        "钱",
        "孙",
        "李",
        "周",
        "吴",
        "郑",
        "王",
        "冯",
        "陈",
        "卫",
        "蒋",
        "沈",
        "韩",
        "杨",
        "朱",
        "秦",
        "尤",
        "许",
        "何",
        "吕",
        "施",
        "张",
        "孔",
        "曹",
        "严",
        "华",
        "金",
        "魏",
        "陶",
        "姜",
        "戚",
        "谢",
        "邹",
        "喻",
        "柏",
        "水",
        "窦",
        "章",
        "云",
        "苏",
        "潘",
        "葛",
        "奚",
        "范",
        "彭",
        "郎",
        "鲁",
        "韦",
        "昌",
        "马",
        "苗",
        "凤",
        "花",
        "方",
        "俞",
        "任",
        "袁",
        "柳",
        "酆",
        "鲍",
        "史",
        "唐",
        "费",
        "廉",
        "岑",
        "薛",
        "雷",
        "贺",
        "倪",
        "汤",
        "滕",
        "殷",
        "罗",
        "毕",
        "郝",
        "邬",
        "安",
        "常",
        "乐",
        "于",
        "时",
        "傅",
        "皮",
        "卞",
        "齐",
        "康",
        "伍",
        "余",
        "元",
        "卜",
        "顾",
        "孟",
        "平",
        "黄",
        "和",
        "穆",
        "萧",
        "尹",
        "欧阳",
        "司马",
        "上官",
        "诸葛",
        "夏侯",
        "东方",
        "端木",
        "独孤",
        "南宫",
        "万俟",
        "闻人",
        "皇甫",
    ],
}


ROLE_PLACEHOLDER_MAP = {
    "甲方": "PARTY_A",
    "乙方": "PARTY_B",
    "丙方": "PARTY_C",
    "丁方": "PARTY_D",
    "买方": "PARTY_A",
    "卖方": "PARTY_B",
    "承包人": "PARTY_A",
    "发包人": "PARTY_B",
    "委托方": "PARTY_A",
    "受托方": "PARTY_B",
    "出租方": "PARTY_A",
    "承租方": "PARTY_B",
    "服务方": "PARTY_B",
    "客户方": "PARTY_A",
}


class PlaceholderAllocator:
    """为不同类型的原始值分配稳定占位符。"""

    def __init__(self):
        self.existing: Dict[Tuple[str, str], str] = {}
        self.counters: Dict[str, int] = {}

    def get_placeholder(self, category: str, value: str) -> str:
        key = (category, value)
        if key in self.existing:
            return self.existing[key]
        self.counters.setdefault(category, 0)
        self.counters[category] += 1
        placeholder = f"[[{category}_{self.counters[category]}]]"
        self.existing[key] = placeholder
        return placeholder


class RedactionResult:
    def __init__(self, text: str, redaction_map: Dict[str, Dict[str, str]], replacements: List[Tuple[int, int, str]], stats: Dict[str, int]):
        self.text = text
        self.redaction_map = redaction_map
        self.replacements = replacements
        self.stats = stats


def ensure_wordlists() -> Dict[str, List[str]]:
    os.makedirs(CONFIG_DIR, exist_ok=True)
    wordlists: Dict[str, List[str]] = {}
    for filename, defaults in WORDLIST_DEFAULTS.items():
        path = os.path.join(CONFIG_DIR, filename)
        if not os.path.exists(path):
            with open(path, "w", encoding="utf-8") as f:
                f.write("\n".join(defaults))
        with open(path, "r", encoding="utf-8") as f:
            lines = [line.strip() for line in f if line.strip()]
            wordlists[filename] = lines
    return wordlists


def build_patterns(wordlists: Dict[str, List[str]]):
    company_suffix = wordlists.get("company_suffix.txt", [])
    suffix_regex = "|".join(map(re.escape, sorted(company_suffix, key=len, reverse=True)))

    family_names = wordlists.get("person_family_names.txt", [])
    family_regex = "|".join(map(re.escape, sorted(family_names, key=len, reverse=True)))

    phone_pattern = re.compile(r"(?<!\d)(?:\+?86[- ]?)?(1[3-9][0-9]{1}[- ]?\d{4}[- ]?\d{4})(?!\d)")
    landline_pattern = re.compile(r"(?<!\d)((?:0\d{2,3}-\d{7,8})|(?:400|800)-?\d{3}-?\d{4})(?!\d)")
    email_pattern = re.compile(r"([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})")
    bank_pattern = re.compile(r"(?<!\d)((?:\d[ \-]?){12,30}\d)(?!\d)")
    uscc_pattern = re.compile(r"([0-9A-Z]{18})")
    id_pattern = re.compile(r"([0-9]{6}\d{8}(?:\d|X|x))")

    amount_pattern = re.compile(
        r"((?:人民币|RMB|CNY|¥|￥)\s*[0-9]{1,3}(?:[,\s]?[0-9]{3})*(?:\.\d+)?(?:\s*(?:元|万元|亿元))?"
        r"|\d+(?:\.\d+)?\s*(?:元|万元|亿元)"
        r"|[壹贰叁肆伍陆柒捌玖拾佰仟万亿]+元(?:整|正)?)"
    )
    date_pattern = re.compile(r"(\d{4}[-/.年]\d{1,2}[-/.月]\d{1,2}[日]?|\d{1,2}月\d{1,2}日)")
    doc_no_pattern = re.compile(r"(合同编号|协议编号)[：:]?\s*([A-Za-z0-9\-_/]{3,})")
    project_no_pattern = re.compile(r"(项目编号)[：:]?\s*([A-Za-z0-9\-_/]{3,})")

    name_keyword_pattern = re.compile(r"(?:联系人|法定代表人|负责人|项目负责人|项目经理|经办人|签字人|授权代表)[：:]?\s*([\u4e00-\u9fa5]{2,3})")
    honorific_pattern = re.compile(r"(?:(?<=\s)|(?<=：)|(?<=:)|(?<=，)|(?<=。)|^)([\u4e00-\u9fa5]{2,3})(先生|女士|经理|老师)")

    company_pattern = re.compile(rf"([\u4e00-\u9fa5A-Za-z0-9（）()·\s]{{2,40}}?(?:{suffix_regex}))") if suffix_regex else None

    patterns = {
        "PHONE": [phone_pattern, landline_pattern],
        "EMAIL": [email_pattern],
        "BANK_ACCOUNT": [bank_pattern],
        "USCC": [uscc_pattern],
        "ID": [id_pattern],
        "AMOUNT": [amount_pattern],
        "DATE": [date_pattern],
        "DOC_NO": [doc_no_pattern],
        "PROJECT_NO": [project_no_pattern],
        "NAME_KEYWORD": [name_keyword_pattern],
    }
    if honorific_pattern:
        patterns["NAME_HONORIFIC"] = [honorific_pattern]
    if company_pattern:
        patterns["COMPANY"] = [company_pattern]
    return patterns


def detect_party(text: str, allocator: PlaceholderAllocator, patterns, wordlists):
    matches = []
    role_words = wordlists.get("role_words.txt", [])
    if not role_words:
        return matches
    company_suffix = wordlists.get("company_suffix.txt", [])
    suffix_regex = "|".join(map(re.escape, sorted(company_suffix, key=len, reverse=True)))
    role_regex = "|".join(map(re.escape, role_words))
    pattern = re.compile(rf"({role_regex})[：: ]+([^\n\r]{{2,40}}?(?:{suffix_regex}))") if suffix_regex else None
    if not pattern:
        return matches
    for m in pattern.finditer(text):
        role = m.group(1)
        company_name = m.group(2).strip()
        category = ROLE_PLACEHOLDER_MAP.get(role, "PARTY_OTHER")
        placeholder = allocator.get_placeholder(category, company_name)
        matches.append((m.start(2), m.end(2), placeholder, category, company_name))
    return matches


def detect_addresses(text: str, allocator: PlaceholderAllocator, wordlists):
    matches = []
    fields = wordlists.get("address_field_words.txt", [])
    components = wordlists.get("address_component_words.txt", [])
    if not fields or not components:
        return matches
    field_regex = "|".join(map(re.escape, fields))
    pattern = re.compile(rf"({field_regex})[：: ]*([^\n\r]{{6,}})")
    for m in pattern.finditer(text):
        address_body = m.group(2)
        if any(comp in address_body for comp in components):
            placeholder = allocator.get_placeholder("ADDRESS", address_body.strip())
            matches.append((m.start(2), m.end(2), placeholder, "ADDRESS", address_body.strip()))
    return matches


def detect_names(text: str, allocator: PlaceholderAllocator, patterns):
    matches = []
    name_keywords = ["联系人", "法定代表人", "负责人", "项目负责人", "项目经理", "经办人", "签字人", "授权代表"]
    for kw in name_keywords:
        pattern = re.compile(
            rf"{re.escape(kw)}[：:]?\s*([\u4e00-\u9fa5]{{2,3}})(?=(?:先生|女士|经理|老师)?[\s，,。；;]|$)"
        )
        for m in pattern.finditer(text):
            person = m.group(1)
            placeholder = allocator.get_placeholder("PERSON", person)
            matches.append((m.start(1), m.end(1), placeholder, "PERSON", person))
    honorific = patterns.get("NAME_HONORIFIC", [])
    for p in honorific:
        for m in p.finditer(text):
            person = m.group(1)
            placeholder = allocator.get_placeholder("PERSON", person)
            matches.append((m.start(1), m.end(1), placeholder, "PERSON", person))
    return matches


def detect_generic(text: str, allocator: PlaceholderAllocator, patterns):
    type_map = {
        "PHONE": "PHONE",
        "EMAIL": "EMAIL",
        "ID": "ID",
        "USCC": "USCC",
        "BANK_ACCOUNT": "BANK_ACCOUNT",
        "AMOUNT": "AMOUNT",
        "DATE": "DATE",
        "DOC_NO": "DOC_NO",
        "PROJECT_NO": "PROJECT_NO",
    }
    matches = []
    for key, category in type_map.items():
        for pat in patterns.get(key, []):
            for m in pat.finditer(text):
                if key in {"DOC_NO", "PROJECT_NO"}:
                    value = m.group(2)
                    start, end = m.start(2), m.end(2)
                else:
                    value = m.group(1)
                    start, end = m.start(1), m.end(1)
                placeholder = allocator.get_placeholder(category, value)
                matches.append((start, end, placeholder, category, value))
    return matches


def detect_companies(text: str, allocator: PlaceholderAllocator, patterns):
    matches = []
    for pat in patterns.get("COMPANY", []):
        for m in pat.finditer(text):
            value = m.group(1).strip()
            placeholder = allocator.get_placeholder("PARTY_OTHER", value)
            matches.append((m.start(1), m.end(1), placeholder, "PARTY_OTHER", value))
    return matches


def redact_text(text: str) -> RedactionResult:
    """对给定文本执行脱敏，返回结果和映射。"""
    wordlists = ensure_wordlists()
    patterns = build_patterns(wordlists)
    allocator = PlaceholderAllocator()

    found: List[Tuple[int, int, str, str, str]] = []
    found.extend(detect_party(text, allocator, patterns, wordlists))
    found.extend(detect_addresses(text, allocator, wordlists))
    found.extend(detect_names(text, allocator, patterns))
    found.extend(detect_generic(text, allocator, patterns))
    found.extend(detect_companies(text, allocator, patterns))

    # 去重和避免重叠
    found.sort(key=lambda x: x[0])
    filtered: List[Tuple[int, int, str, str, str]] = []
    last_end = -1
    for start, end, placeholder, category, value in found:
        if start < last_end:
            continue
        filtered.append((start, end, placeholder, category, value))
        last_end = end

    redaction_map: Dict[str, Dict[str, str]] = {}
    stats: Dict[str, int] = {}

    result_parts = []
    cursor = 0
    replacements: List[Tuple[int, int, str]] = []
    for start, end, placeholder, category, value in filtered:
        if start < cursor:
            continue
        result_parts.append(text[cursor:start])
        result_parts.append(placeholder)
        replacements.append((start, end, placeholder))
        cursor = end
        redaction_map[placeholder] = {"type": category, "value": value}
        stats[category] = stats.get(category, 0) + 1
    result_parts.append(text[cursor:])
    redacted_text = "".join(result_parts)

    return RedactionResult(redacted_text, redaction_map, replacements, stats)


def document_to_text_and_mapping(doc: "Document") -> Tuple[str, List[Tuple[int, int, object]]]:
    """抽取全文字符串并记录每个 run 的字符区间。"""
    text_parts: List[str] = []
    run_spans: List[Tuple[int, int, object]] = []
    cursor = 0

    def process_paragraph(paragraph):
        nonlocal cursor
        for run in paragraph.runs:
            run_text = run.text
            start, end = cursor, cursor + len(run_text)
            run_spans.append((start, end, run))
            text_parts.append(run_text)
            cursor = end
        text_parts.append("\n")
        cursor += 1

    for block in iter_block_items(doc):
        if isinstance(block, type(doc.paragraphs[0])):
            process_paragraph(block)
        else:
            # table
            for row in block.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        process_paragraph(paragraph)
    return "".join(text_parts), run_spans


def iter_block_items(parent):
    from docx import Document
    from docx.table import _Cell, Table
    from docx.text.paragraph import Paragraph

    if isinstance(parent, Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        return []

    for child in parent_elm.iterchildren():
        if child.tag.endswith('}p'):
            yield Paragraph(child, parent)
        elif child.tag.endswith('}tbl'):
            yield Table(child, parent)


def apply_replacements_to_docx(doc: "Document", replacements: List[Tuple[int, int, str]]):
    _, run_spans = document_to_text_and_mapping(doc)
    replacements = sorted(replacements, key=lambda x: x[0])

    for start, end, placeholder in replacements:
        for run_start, run_end, run in run_spans:
            if end <= run_start or start >= run_end:
                continue
            local_start = max(0, start - run_start)
            local_end = min(run_end, end) - run_start
            new_text = run.text
            new_text = new_text[:local_start] + placeholder + new_text[local_end:]
            run.text = new_text
    return doc


def write_outputs(original_path: str, original_text: str, result: RedactionResult, doc: Optional["Document"], output_root: str = OUTPUT_DIR) -> str:
    os.makedirs(output_root, exist_ok=True)
    base_name = os.path.basename(original_path)
    name, ext = os.path.splitext(base_name)
    folder_name = f"{datetime.datetime.now():%Y%m%d_%H%M%S}_{name}_脱敏"
    folder_path = os.path.join(output_root, folder_name)
    os.makedirs(folder_path, exist_ok=True)

    txt_path = os.path.join(folder_path, f"{name}_redacted.txt")
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write(result.text)

    if doc is not None:
        redacted_doc = apply_replacements_to_docx(doc, result.replacements)
        docx_path = os.path.join(folder_path, f"{name}_redacted.docx")
        redacted_doc.save(docx_path)
    else:
        docx_path = None

    map_path = os.path.join(folder_path, f"{name}_redaction_map.json")
    with open(map_path, "w", encoding="utf-8") as f:
        json.dump(result.redaction_map, f, ensure_ascii=False, indent=2)

    return folder_path


def process_docx(path: str, output_root: str = OUTPUT_DIR) -> Tuple[str, RedactionResult]:
    from docx import Document

    doc = Document(path)
    text, _ = document_to_text_and_mapping(doc)
    result = redact_text(text)
    folder = write_outputs(path, text, result, doc, output_root=output_root)
    return folder, result


def process_txt(path: str, output_root: str = OUTPUT_DIR) -> Tuple[str, RedactionResult]:
    with open(path, "r", encoding="utf-8", errors="ignore") as f:
        text = f.read()
    result = redact_text(text)
    folder = write_outputs(path, text, result, None, output_root=output_root)
    return folder, result


def run_self_test():
    """构造典型样例，快速验证规则效果。"""
    sample = (
        "甲方：北京星河生物科技有限公司，注册地址：北京市海淀区知春路88号星辉大厦20层\n"
        "乙方：上海光速信息技术有限公司（下称乙方），法定代表人：李雷，联系人王芳女士，电话：138-1234-5678\n"
        "合同编号：GX-2025-009，项目编号：PRJ-7788\n"
        "开户行账号：6222 1234 5678 9090 123，统一社会信用代码：91310000MA1K123X8\n"
        "身份证号：110101199912123456，联系地址：上海市浦东新区世纪大道100号世纪汇广场15层\n"
        "付款金额人民币 1,200,000 元，签署日期：2025年12月15日"
    )
    result = redact_text(sample)
    print("原文:\n", sample)
    print("\n脱敏后:\n", result.text)
    print("\n映射示例:")
    for k, v in list(result.redaction_map.items())[:5]:
        print(k, v)
    print("\n统计:", result.stats)


if __name__ == "__main__":
    run_self_test()
