import streamlit as st
import pandas as pd
import re
import os
import zipfile
import tempfile
import docx
from docx.document import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph
import pdfplumber

# --- 核心辅助函数 ---

def clean_filename(text):
    text = str(text).strip()
    return re.sub(r'[\\/:*?"<>|\n\r]', '', text)[:80]

def get_level(text):
    text = str(text).strip()
    # 【核心修复1】：暴力清除开头的各种 Word 特殊项目符号和空格
    match_text = re.sub(r'^[\s▪•►◆■●○➢\-\·]+', '', text)
    
    # 4级: 8-1-1, 1.1.1
    if re.match(r'^\d+[-_\.]\d+[-_\.]\d+', match_text): return 4
    # 3级: 1-1, 1.1
    elif re.match(r'^\d+[-_\.]\d+', match_text): return 3
    # 2级: 1、 1. (1) （1） 
    elif re.match(r'^\d+[、\.\s]', match_text) or re.match(r'^[\(（]\d+[\)）]', match_text) or re.match(r'^\d+[）\)]', match_text): return 2
    # 1级 (传统): 一、 第一部分
    elif re.match(r'^第?[一二三四五六七八九十百]+[、\.\s部分章]', match_text) or re.match(r'^[\(（][一二三四五六七八九十百]+[\)）]', match_text): return 1
    # 【核心修复2】：专属法律底稿识别规则。即使“三、”被 Word 自动编号隐藏，依然能识别《审核问询函》问题、反馈等
    elif re.match(r'^《[^》]+》(?:问题|回复|补充|反馈|关注|部分|章节|关于)', match_text): return 1
    elif re.match(r'^(?:问题|附件|反馈|答复)\s*[\d一二三四五六七八九十]', match_text): return 1
    
    return 0

def is_valid_flat_folder(text):
    if len(text) < 2: return False
    if text in ["原件", "复印件", "卷宗号", "卷宗", "纸质", "电子", "扫描件", "文件资料", "查验内容"]: return False
    if re.match(r'^[\d\s\-\.,/]+$', text): return False 
    return True

def analyze_header(matrix):
    seq_col = -1
    skip_cols = set()
    start_row = 0

    for i, row in enumerate(matrix[:5]):
        row_str = "".join([str(c) for c in row if c]).replace(" ", "").replace("\n", "")
        if not row_str: continue

        if ("序号" in row_str and ("文件" in row_str or "名称" in row_str or "资料" in row_str or "查验" in row_str)) or ("原件" in row_str and "复印件" in row_str):
            start_row = max(start_row, i + 1)
            for j, cell in enumerate(row):
                cell_str = str(cell).replace(" ", "").replace("\n", "")
                if "序号" in cell_str or "编号" in cell_str:
                    seq_col = j
                elif "备注" in cell_str or "索引" in cell_str:
                    skip_cols.add(j)

    return seq_col, skip_cols, start_row

class PathBuilder:
    def __init__(self):
        self.stack = {1: None, 2: None, 3: None, 4: None}
        self.paths = []
        self.seen = set()
        
    def add(self, text, level, allow_flat):
        clean_name = clean_filename(text)
        if not clean_name: return
            
        if level > 0:
            self.stack[level] = clean_name
            for i in range(level + 1, 5): self.stack[i] = None
            parts = [self.stack[i] for i in range(1, level + 1) if self.stack[i]]
            full_path = "/".join(parts)
        else:
            if not allow_flat: return
            active_level = 0
            for i in range(4, 0, -1):
                if self.stack[i]:
                    active_level = i
                    break
            if active_level > 0:
                parts = [self.stack[i] for i in range(1, active_level + 1) if self.stack[i]]
                parts.append(clean_name)
                full_path = "/".join(parts)
            else:
                full_path = clean_name
                
        if full_path and full_path not in self.seen:
            self.seen.add(full_path)
            self.paths.append(full_path)

def process_matrix(matrix, builder, allow_flat):
    seq_col, skip_cols, start_row = analyze_header(matrix)

    for i in range(start_row, len(matrix)):
        row = matrix[i]
        seq_prefix = ""

        for j, cell in enumerate(row):
            if j in skip_cols:
                continue

            cell_text = str(cell).strip() if pd.notna(cell) and cell is not None else ""
            if not cell_text or cell_text.lower() == "nan":
                continue

            if cell_text.endswith(".0"): 
                cell_text = cell_text[:-2]

            lvl = get_level(cell_text)

            if j == seq_col and lvl == 0 and len(cell_text) <= 5:
                seq_prefix = cell_text
                continue

            if lvl > 0:
                builder.add(cell_text, lvl, allow_flat)
                seq_prefix = "" 
            else:
                if seq_prefix:
                    combined = f"{seq_prefix} {cell_text}".strip()
                    comb_lvl = get_level(combined)
                    if comb_lvl > 0:
                        builder.add(combined, comb_lvl, allow_flat)
                        seq_prefix = ""
                    else:
                        if allow_flat and is_valid_flat_folder(cell_text):
                            builder.add(combined, 0, allow_flat)
                            seq_prefix = ""
                else:
                    if allow_flat and is_valid_flat_folder(cell_text):
                        builder.add(cell_text, 0, allow_flat)

def iter_block_items(parent):
    if isinstance(parent, Document): parent_elm = parent.element.body
    elif isinstance(parent, _Cell): parent_elm = parent._tc
    else: return
    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P): yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl): yield Table(child, parent)

# --- 🌲 树状目录可视化 ---
def build_tree_dict(paths):
    tree = {}
    for path in paths:
        parts = path.split('/')
        current_node = tree
        for part in parts:
            if part not in current_node: current_node[part] = {}
            current_node = current_node[part]
    return tree

def generate_tree_string(tree_dict, indent=""):
    result = ""
    items = list(tree_dict.items())
    for i, (folder_name, sub_tree) in enumerate(items):
        is_last = (i == len(items) - 1)
        prefix = "└── " if is_last else "├── "
        result += f"{indent}{prefix}📁 {folder_name}\n"
        next_indent = indent + ("    " if is_last else "│   ")
        result += generate_tree_string(sub_tree, next_indent)
    return result

# --- 三大文件解析引擎 ---

def parse_excel(file, allow_flat):
    df = pd.read_excel(file, header=None)
    matrix = df.values.tolist()
    builder = PathBuilder()
    process_matrix(matrix, builder, allow_flat)
    return builder.paths

def parse_word(file, allow_flat):
    doc = docx.Document(file)
    builder = PathBuilder()
    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            text = block.text.strip()
            if text:
                level = get_level(text)
                if level > 0: builder.add(text, level, allow_flat=False) 
        elif isinstance(block, Table):
            matrix = [[cell.text.strip() for cell in row.cells] for row in block.rows]
            process_matrix(matrix, builder, allow_flat)
    return builder.paths

def parse_pdf(file, allow_flat):
    builder = PathBuilder()
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                for line in text.split('\n'):
                    line = line.strip()
                    level = get_level(line)
                    if level > 0: builder.add(line, level, allow_flat=False)
            tables = page.extract_tables()
            for table in tables:
                if not table: continue
                matrix = [[str(cell).replace('\n', ' ').strip() if cell else "" for cell in row] for row in table]
                process_matrix(matrix, builder, allow_flat)
    return builder.paths


# --- Streamlit 网页界面 ---
st.set_page_config(page_title="自动化多层级文件夹生成", layout="wide")
st.title("📂 自动化多层级文件夹生成器")

allow_flat = st.checkbox("开启宽松模式 (仅针对表格内容有效：若表格内的文字没有编号，自动将其放入最近一级的父文件夹中)", value=True)

uploaded_files = st.file_uploader("请上传包含底稿目录的文件（支持直接拖拽多个文件）", type=["xlsx", "docx", "pdf", "doc"], accept_multiple_files=True)

if uploaded_files:
    for uploaded_file in uploaded_files:
        paths_to_create = []
        
        if uploaded_file.name.endswith(".doc"):
            st.error(f"⚠️ 文件 {uploaded_file.name} 是老旧的 `.doc` 格式。请按 `F12` 另存为 `.docx` 格式后重新上传。")
            continue
            
        if uploaded_file.name.endswith(".xlsx"): paths_to_create = parse_excel(uploaded_file, allow_flat)
        elif uploaded_file.name.endswith(".docx"): paths_to_create = parse_word(uploaded_file, allow_flat)
        elif uploaded_file.name.endswith(".pdf"): paths_to_create = parse_pdf(uploaded_file, allow_flat)
                
        if not paths_to_create:
            st.warning(f"⚠️ 未能从文件 {uploaded_file.name} 中解析出结构，请检查目录格式。")
        else:
            st.success(f"✅ 文件 {uploaded_file.name} 解析成功！共生成 {len(paths_to_create)} 个路径。")
            
            with st.expander(f"🌲 点击预览 {uploaded_file.name} 的层级结构树", expanded=True):
                tree_dict = build_tree_dict(paths_to_create)
                tree_visual_str = generate_tree_string(tree_dict)
                st.code(tree_visual_str, language="text")
            
            with tempfile.TemporaryDirectory() as tmpdirname:
                zip_filename = os.path.join(tmpdirname, f"folders_{uploaded_file.name}.zip")
                with zipfile.ZipFile(zip_filename, 'w') as zipf:
                    for path in paths_to_create:
                        full_path = os.path.join(tmpdirname, path)
                        os.makedirs(full_path, exist_ok=True)
                        zipf.writestr(zipfile.ZipInfo(path + "/"), "")
                        
                with open(zip_filename, "rb") as f:
                    st.download_button(
                        label=f"⬇️ 下载 {uploaded_file.name} 对应的文件夹 (ZIP)",
                        data=f,
                        file_name=f"目录结构_{uploaded_file.name.rsplit('.', 1)[0]}.zip",
                        mime="application/zip",
                        key=uploaded_file.name
                    )