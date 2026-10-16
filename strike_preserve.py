# -*- coding: utf-8 -*-
"""删除线（富文本）保真模块。

背景：openpyxl 在 load_workbook 读入时，会把单元格内带 run 级格式的富文本
（inline `<is>` 或共享字符串 `<si>`，例如“部分文字带删除线”）压平成普通字符串，
导致最终导出的 xlsx 丢失删除线等格式。cell 级整格删除线（font.strikethrough）
不受影响，但很多平台把“整格删除线”也写成单个富文本 run，同样会被压平。

方案：在 openpyxl 处理【之前】从原始 xlsx 中提取每个富文本单元格的 run 结构
（含 rPr 原始 XML），在筛选处理（删行/删列）【之后】按坐标映射把 run 结构
原样写回输出文件，从而完整保住删除线及其他 run 级格式。本模块不修改
filter_core 的筛选规则，只负责“提取 → 回写”。
"""
import io
import re
import zipfile
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape

from openpyxl.utils.cell import column_index_from_string, get_column_letter

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships"

_CELL_RE = re.compile(r'<c\b[^>]*\br="([A-Z]+)(\d+)"[^>]*>')


# ---------------------------------------------------------------- 提取
def _parse_sheet_map(z: zipfile.ZipFile) -> dict:
    """{页名: xl 内 xml 路径}。"""
    try:
        wb_root = ET.fromstring(z.read("xl/workbook.xml"))
        rel_root = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    except (KeyError, ET.ParseError):
        return {}
    rid_by_name = {}
    for sh in wb_root.findall(f"{{{NS_MAIN}}}sheets/{{{NS_MAIN}}}sheet"):
        rid_by_name[sh.get("name")] = sh.get(f"{{{NS_REL}}}id")
    target_by_rid = {}
    for rel in rel_root.findall(f"{{{NS_PKG}}}Relationship"):
        target_by_rid[rel.get("Id")] = rel.get("Target")
    result = {}
    for name, rid in rid_by_name.items():
        t = target_by_rid.get(rid)
        if not t:
            continue
        if t.startswith("/"):
            t = t.lstrip("/")
        else:
            t = "xl/" + t.lstrip("/")
        result[name] = t
    return result


def _clean_fragment(frag: str) -> str:
    """去掉 ET 序列化时可能加的命名空间前缀与声明，让片段直接嵌入 worksheet。"""
    frag = re.sub(r'\sxmlns:ns\d+="[^"]*"', "", frag)
    frag = re.sub(r'</ns\d+:', "</", frag)          # 闭合标签：</ns0:rPr> -> </rPr>
    frag = re.sub(r'<ns\d+:', "<", frag)            # 开标签：<ns0:rPr> -> <rPr>
    frag = re.sub(r'\sxmlns="[^"]*"', "", frag)
    return frag


def _runs_from_r_list(r_list, si=None) -> list or None:
    """把一组 <r> run 转成 [(text, rPr原始XML或None)]；无 run 返回 None。"""
    runs = []
    for r in r_list:
        rpr = r.find(f"{{{NS_MAIN}}}rPr")
        t = r.find(f"{{{NS_MAIN}}}t")
        text = ""
        if t is not None and t.text:
            text = t.text
        rpr_xml = None
        if rpr is not None:
            rpr_xml = _clean_fragment(ET.tostring(rpr, encoding="unicode"))
        runs.append((text, rpr_xml))
    return runs or None


def _shared_runs_list(z: zipfile.ZipFile):
    """共享字符串表：下标 -> run 列表（含 rPr）；普通字符串或空 -> None。"""
    try:
        root = ET.fromstring(z.read("xl/sharedStrings.xml"))
    except (KeyError, ET.ParseError):
        return []
    out = []
    for si in root.findall(f"{{{NS_MAIN}}}si"):
        out.append(_runs_from_r_list(si.findall(f"{{{NS_MAIN}}}r")))
    return out


def _sheet_rich_cells(z: zipfile.ZipFile, xml_path: str, shared_runs) -> dict:
    """{单元格引用: [(text, rPrXML或None), ...]}，只含带 run 的富文本单元格。"""
    try:
        root = ET.fromstring(z.read(xml_path))
    except (KeyError, ET.ParseError):
        return {}
    result = {}
    for c in root.iter(f"{{{NS_MAIN}}}c"):
        ref = c.get("r")
        if not ref:
            continue
        isel = c.find(f"{{{NS_MAIN}}}is")
        if isel is not None:
            runs = _runs_from_r_list(isel.findall(f"{{{NS_MAIN}}}r"))
            if runs:
                result[ref] = runs
            continue
        if c.get("t") == "s":
            v = c.find(f"{{{NS_MAIN}}}v")
            if v is not None and v.text and shared_runs:
                idx = int(v.text)
                if idx < len(shared_runs) and shared_runs[idx]:
                    result[ref] = shared_runs[idx]
    return result


def extract_rich_cells(data: bytes) -> dict:
    """{页名: {单元格引用: [(text, rPrXML), ...]}}。无富文本时返回空 dict。"""
    try:
        z = zipfile.ZipFile(io.BytesIO(data))
    except zipfile.BadZipFile:
        return {}
    with z:
        sheet_map = _parse_sheet_map(z)
        if not sheet_map:
            return {}
        shared_runs = _shared_runs_list(z)
        result = {}
        for name, path in sheet_map.items():
            cells = _sheet_rich_cells(z, path, shared_runs)
            if cells:
                result[name] = cells
    return result


# ---------------------------------------------------------------- 坐标映射
def _map_ref(ref: str, del_cols, del_rows) -> str or None:
    """原始引用 -> 处理后新引用。del_cols: 原坐标 [(c0,c1)]；del_rows: 原行号 [r]。

    单元格本身落在被删列组或被删行内 → 返回 None（该单元格已不存在）；
    否则返回新坐标（向左移过被删列组、向上移过被删行）。
    """
    m = re.match(r"^([A-Z]+)(\d+)$", ref)
    if not m:
        return None
    col = column_index_from_string(m.group(1))
    row = int(m.group(2))
    shift = 0
    for c0, c1 in del_cols:
        if c0 <= col <= c1:
            return None                      # 单元格在被删项目列组内 → 已被删除
        if c1 < col:                         # 被删项目列组整体在左侧 → 左移
            shift += c1 - c0 + 1
    if row in set(del_rows):
        return None                          # 整行被删除 → 单元格不存在
    new_col = col - shift
    new_row = row - sum(1 for d in del_rows if d < row)
    if new_col < 1 or new_row < 1:
        return None
    return f"{get_column_letter(new_col)}{new_row}"


# ---------------------------------------------------------------- 回写
def _replace_cell(xml_text: str, ref: str, new_el: str) -> str:
    """把 sheet XML 中 r=ref 的 <c> 元素整体替换为 new_el（保留 s/cm/vm/ph 属性）。"""
    head_pat = re.compile(r'<c\b[^>]*\br="%s"[^>]*>' % re.escape(ref))
    m = head_pat.search(xml_text)
    if not m:
        return None
    start = m.start()
    end = xml_text.find("</c>", m.end())
    if end == -1:
        return None
    end += len("</c>")
    return xml_text[:start] + new_el + xml_text[end:]


def _cell_attrs_to_keep(head: str) -> str:
    """从原 <c> 开头标签里带出样式属性（s/cm/vm/ph），避免替换后丢样式。"""
    attrs = ""
    for a in ("s", "cm", "vm", "ph"):
        m = re.search(r'\b%s="(\d+)"' % a, head)
        if m:
            attrs += ' %s="%s"' % (a, m.group(1))
    return attrs


def _build_cell(ref: str, head_attrs: str, runs) -> str:
    """构造带富文本 run 的内联字符串单元格 <c r=".." s=".." t="inlineStr"><is>..</is></c>。"""
    parts = [f'<c r="{ref}"{head_attrs} t="inlineStr"><is>']
    for text, rpr in runs:
        parts.append("<r>")
        if rpr:
            parts.append(rpr)
        parts.append(f'<t xml:space="preserve">{escape(text)}</t></r>')
    parts.append("</is></c>")
    return "".join(parts)


def restore_rich_text(out_bytes: bytes, rich_cells: dict, trace: dict) -> bytes:
    """把处理后的 xlsx 字节流里对应单元格回写成带删除线的富文本。

    trace: filter_core.process_workbook 记录的筛选结果：
        {页名: {'deleted': True} 或 {'cols': [(c0,c1)...], 'rows': [r...]}}
    """
    if not rich_cells:
        return out_bytes
    try:
        zin = zipfile.ZipFile(io.BytesIO(out_bytes))
    except zipfile.BadZipFile:
        return out_bytes
    with zin:
        sheet_map = _parse_sheet_map(zin)
        patches = {}                       # xml路径 -> {新引用: 新元素}
        for sheet, cells in rich_cells.items():
            info = trace.get(sheet)
            xml_path = sheet_map.get(sheet)
            if info is None or info.get("deleted") or not xml_path:
                continue
            del_cols = info.get("cols", [])
            del_rows = info.get("rows", [])
            for ref, runs in cells.items():
                new_ref = _map_ref(ref, del_cols, del_rows)
                if new_ref is None:
                    continue
                patches.setdefault(xml_path, {})[new_ref] = runs
        if not patches:
            return out_bytes

        out = io.BytesIO()
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename in patches:
                    xml_text = data.decode("utf-8")
                    changed = False
                    for ref, runs in patches[item.filename].items():
                        head_pat = re.compile(r'<c\b[^>]*\br="%s"[^>]*>' % re.escape(ref))
                        m = head_pat.search(xml_text)
                        if not m:
                            continue
                        attrs = _cell_attrs_to_keep(m.group(0))
                        new_el = _build_cell(ref, attrs, runs)
                        replaced = _replace_cell(xml_text, ref, new_el)
                        if replaced is None:
                            continue
                        xml_text = replaced
                        changed = True
                    if changed:
                        data = xml_text.encode("utf-8")
                zout.writestr(item, data)
        return out.getvalue()
