# -*- coding: utf-8 -*-
"""
Vercel 部署版：Flask + docx 生成
关键修复：
1) 所有路径改为基于 __file__ 的绝对路径（Vercel 函数工作目录不是仓库根）
2) / 路由从 public/index.html 返回
3) output 写到 /tmp（Vercel 允许写的目录）
"""

import os
import re
import uuid
import datetime
import zipfile

from flask import Flask, request, jsonify, send_from_directory, abort
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# ============ 路径：全部用绝对路径（Vercel 必须） ============
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

PUBLIC_DIR = os.path.join(BASE_DIR, "public")
DOCX_TEMPLATE_ROOT = os.path.join(BASE_DIR, "templates")

# Vercel 函数只能写 /tmp
OUTPUT_ROOT = os.path.join("/tmp", "fund-docx-output")
os.makedirs(OUTPUT_ROOT, exist_ok=True)

app = Flask(__name__)

PH_RE = re.compile(r"{{(.*?)}}")

# ============ 可选：下拉框配置（按需自行改） ============
SELECT_OPTIONS = {
    "基金类型": [
        "股票型基金", "混合型基金", "债券型基金", "货币市场基金",
        "基金中基金（FOF）", "QDII", "商品型基金", "REITs", "其他",
    ],
    "上市交易所": ["上海证券交易所", "深圳证券交易所", "北京证券交易所", "香港交易所", "其他"],
    "ETF类型": ["跨市场ETF", "单市场ETF", "跨境ETF", "债券ETF", "商品ETF", "其他"],
}

LONG_TEXT_KEYWORDS = ["风险", "揭示", "策略", "分析", "简介", "介绍", "说明", "情况", "内容"]


# ============ 工具函数 ============
def _safe_join_under_root(root: str, user_path: str) -> str:
    user_path = (user_path or "").strip()
    joined = os.path.abspath(os.path.join(root, user_path))
    root_abs = os.path.abspath(root)
    if not (joined == root_abs or joined.startswith(root_abs + os.sep)):
        raise ValueError("非法路径")
    return joined


def safe_filename(name: str) -> str:
    name = (name or "").replace("\\", "_").replace("/", "_")
    name = re.sub(r'[<>:"|?*]', "_", name)
    name = re.sub(r"\s+", " ", name).strip()
    return (name or "output")[:180]


def list_docx_files(folder: str):
    if not os.path.isdir(folder):
        return []
    return sorted([
        f for f in os.listdir(folder)
        if f.lower().endswith(".docx") and not f.startswith("~$")
    ])


def get_product_folder(product_type: str) -> str:
    if not product_type or product_type == "__root__":
        return DOCX_TEMPLATE_ROOT
    return _safe_join_under_root(DOCX_TEMPLATE_ROOT, product_type)


def list_product_types():
    types = []
    if os.path.isdir(DOCX_TEMPLATE_ROOT):
        for name in sorted(os.listdir(DOCX_TEMPLATE_ROOT)):
            full = os.path.join(DOCX_TEMPLATE_ROOT, name)
            if os.path.isdir(full) and not name.startswith("."):
                if list_docx_files(full):
                    types.append({"id": name, "name": name})
    if list_docx_files(DOCX_TEMPLATE_ROOT):
        types.insert(0, {"id": "__root__", "name": "默认（templates根目录）"})
    return types


# ============ Word：XML 级别替换（尽量保留格式） ============
def _set_text_preserve_space(t_elem, text: str):
    t_elem.text = text
    if text and (text[0].isspace() or text[-1].isspace() or "  " in text):
        t_elem.set(qn("xml:space"), "preserve")
    else:
        attr = qn("xml:space")
        if attr in t_elem.attrib:
            del t_elem.attrib[attr]


def _insert_after(parent, ref_child, new_child):
    parent.insert(parent.index(ref_child) + 1, new_child)


def _append_text_with_breaks(first_t, text_to_insert: str):
    run = first_t.getparent()  # w:r
    segments = (text_to_insert or "").split("\n")

    existing = first_t.text or ""
    _set_text_preserve_space(first_t, existing + segments[0])

    last_elem = first_t
    for seg in segments[1:]:
        br = OxmlElement("w:br")
        _insert_after(run, last_elem, br)
        last_elem = br
        if seg != "":
            tnew = OxmlElement("w:t")
            _set_text_preserve_space(tnew, seg)
            _insert_after(run, last_elem, tnew)
            last_elem = tnew


def replace_placeholders_in_element(element, mapping: dict) -> int:
    replaced = 0
    for p in element.xpath(".//w:p"):
        guard = 0
        while True:
            guard += 1
            if guard > 300:
                break

            t_elems = p.xpath(".//w:t")
            if not t_elems:
                break

            full_text = "".join([t.text or "" for t in t_elems])
            matches = list(PH_RE.finditer(full_text))

            target = None
            for m in matches:
                key = (m.group(1) or "").strip()
                if key in mapping:
                    target = m
                    break
            if not target:
                break

            key = target.group(1).strip()
            replacement = "" if mapping.get(key) is None else str(mapping.get(key))
            start, end = target.span()

            cum = []
            pos = 0
            for t in t_elems:
                l = len(t.text or "")
                cum.append((pos, pos + l))
                pos += l

            def locate(posi: int):
                for i, (s, e) in enumerate(cum):
                    if s <= posi < e:
                        return i, posi - s
                return None, None

            first_i, start_off = locate(start)
            last_i, _ = locate(end - 1)
            if first_i is None or last_i is None:
                break

            end_off = (end - 1 - cum[last_i][0]) + 1

            first_t = t_elems[first_i]
            last_t = t_elems[last_i]
            first_text = first_t.text or ""
            last_text = last_t.text or ""

            prefix = first_text[:start_off]
            suffix = last_text[end_off:]

            if first_i == last_i:
                _set_text_preserve_space(first_t, prefix)
                _append_text_with_breaks(first_t, replacement + suffix)
            else:
                _set_text_preserve_space(first_t, prefix)
                _append_text_with_breaks(first_t, replacement)
                for j in range(first_i + 1, last_i):
                    _set_text_preserve_space(t_elems[j], "")
                _set_text_preserve_space(last_t, suffix)

            replaced += 1
    return replaced


def replace_placeholders_in_doc(doc: Document, mapping: dict):
    replace_placeholders_in_element(doc._part._element, mapping)
    for section in doc.sections:
        parts = [
            section.header, section.footer,
            section.first_page_header, section.first_page_footer,
            section.even_page_header, section.even_page_footer,
        ]
        for part in parts:
            if part is not None:
                replace_placeholders_in_element(part._element, mapping)


def scan_placeholders_and_longflags(doc: Document):
    keys = set()
    long_candidate = set()

    def _scan_element(el):
        nonlocal keys, long_candidate
        for p in el.xpath(".//w:p"):
            text = "".join([t.text or "" for t in p.xpath(".//w:t")])
            text_stripped = text.strip()

            found = [m.group(1).strip() for m in PH_RE.finditer(text)]
            for k in found:
                keys.add(k)

            if len(found) == 1 and text_stripped == "{{%s}}" % found[0]:
                long_candidate.add(found[0])

    _scan_element(doc._part._element)
    for section in doc.sections:
        parts = [
            section.header, section.footer,
            section.first_page_header, section.first_page_footer,
            section.even_page_header, section.even_page_footer,
        ]
        for part in parts:
            if part is not None:
                _scan_element(part._element)

    return keys, long_candidate


def build_schema(product_type: str):
    folder = get_product_folder(product_type)
    tpl_files = list_docx_files(folder)
    if not tpl_files:
        return {"product_type": product_type, "templates": [], "fields": []}

    all_keys = set()
    long_flags = set()

    for f in tpl_files:
        p = os.path.join(folder, f)
        doc = Document(p)
        keys, longs = scan_placeholders_and_longflags(doc)
        all_keys |= keys
        long_flags |= longs

    fields = []
    for key in sorted(all_keys):
        ftype = "text"
        if key in SELECT_OPTIONS:
            ftype = "select"
        elif key in long_flags or any(kw in key for kw in LONG_TEXT_KEYWORDS):
            ftype = "textarea"

        item = {"key": key, "label": key, "type": ftype}
        if ftype == "select":
            item["options"] = SELECT_OPTIONS.get(key, [])
        fields.append(item)

    return {"product_type": product_type, "templates": tpl_files, "fields": fields}


def apply_placeholders_to_filename(filename: str, mapping: dict, blank_unfilled: bool = True) -> str:
    def _rep(m):
        k = (m.group(1) or "").strip()
        if k in mapping:
            return "" if mapping[k] is None else str(mapping[k])
        return "" if blank_unfilled else m.group(0)

    name = PH_RE.sub(_rep, filename)
    name = re.sub(r"{{.*?}}", "" if blank_unfilled else "", name)
    if not name.lower().endswith(".docx"):
        name += ".docx"
    return safe_filename(name)


def create_zip_from_folder(folder: str, zip_path: str):
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        for root, _, files in os.walk(folder):
            for fn in files:
                if fn.startswith("~$"):
                    continue
                full = os.path.join(root, fn)
                rel = os.path.relpath(full, folder)
                z.write(full, arcname=rel)


# ============ 路由 ============
@app.route("/")
def index():
    # 关键：用绝对路径 PUBLIC_DIR
    return send_from_directory(PUBLIC_DIR, "index.html")


@app.route("/api/product_types")
def api_product_types():
    if not os.path.isdir(DOCX_TEMPLATE_ROOT):
        return jsonify({"ok": False, "error": "templates 目录不存在"}), 400
    return jsonify({"ok": True, "data": list_product_types()})


@app.route("/api/schema")
def api_schema():
    product_type = request.args.get("product_type", "__root__")
    try:
        folder = get_product_folder(product_type)
    except Exception:
        return jsonify({"ok": False, "error": "product_type 非法"}), 400

    if not os.path.isdir(folder):
        return jsonify({"ok": False, "error": f"未找到产品类型目录：{product_type}"}), 404

    schema = build_schema(product_type)
    return jsonify({"ok": True, "data": schema})


@app.route("/api/generate", methods=["POST"])
def api_generate():
    payload = request.get_json(force=True, silent=False) or {}

    product_type = payload.get("product_type", "__root__")
    values = payload.get("values", {}) or {}
    mode = payload.get("mode", "all")  # all / selected
    selected = payload.get("selected_templates", []) or []
    blank_unfilled = bool(payload.get("blank_unfilled", True))

    try:
        folder = get_product_folder(product_type)
    except Exception:
        return jsonify({"ok": False, "error": "product_type 非法"}), 400

    tpl_files = list_docx_files(folder)
    if not tpl_files:
        return jsonify({"ok": False, "error": f"该产品类型下没有 .docx 模板：{product_type}"}), 400

    if mode == "selected":
        tpl_files = [f for f in tpl_files if f in selected]
        if not tpl_files:
            return jsonify({"ok": False, "error": "未选择任何模板"}), 400

    schema = build_schema(product_type)
    all_keys = [f["key"] for f in schema.get("fields", [])]

    mapping = {}
    for k in all_keys:
        if k in values:
            v = values.get(k)
            mapping[k] = "" if v is None else str(v)
        else:
            mapping[k] = "" if blank_unfilled else None

    run_name = f"{product_type}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:6]}"
    run_dir = os.path.join(OUTPUT_ROOT, run_name)
    os.makedirs(run_dir, exist_ok=True)

    generated_files = []
    for tpl in tpl_files:
        tpl_path = os.path.join(folder, tpl)
        try:
            doc = Document(tpl_path)
            replace_placeholders_in_doc(doc, mapping)

            out_name = apply_placeholders_to_filename(tpl, mapping, blank_unfilled=blank_unfilled)
            out_path = os.path.join(run_dir, out_name)
            doc.save(out_path)
            generated_files.append(out_name)
        except Exception as e:
            return jsonify({"ok": False, "error": f"生成失败：{tpl}；原因：{repr(e)}"}), 500

    zip_name = f"{run_name}.zip"
    zip_path = os.path.join(OUTPUT_ROOT, zip_name)
    create_zip_from_folder(run_dir, zip_path)

    # 这里先返回一个 /download 路径（注意：在 Vercel 上跨请求下载可能不稳定，后续如需稳定下载再上 Blob）
    return jsonify({
        "ok": True,
        "data": {
            "job_id": run_name,
            "zip_name": zip_name,
            "zip_url": f"/download/{zip_name}",
            "generated_files": generated_files,
        }
    })


@app.route("/download/<path:filename>")
def download(filename):
    safe_rel = os.path.normpath(filename)
    if safe_rel.startswith("..") or os.path.isabs(safe_rel):
        abort(400)

    full = os.path.join(OUTPUT_ROOT, safe_rel)
    if not os.path.isfile(full):
        abort(404)

    return send_from_directory(OUTPUT_ROOT, os.path.basename(full), as_attachment=True)

