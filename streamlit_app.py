import streamlit as st
import re
import random
import zipfile
import io
import os
from xml.dom import minidom
import pandas as pd

# ==================== PHẦN 1: LOGIC XỬ LÝ (CORE) ====================

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

def parse_range_string(s):
    res = set()
    if not s: return res
    parts = str(s).split(',')
    for part in parts:
        part = part.strip()
        if not part: continue
        if '-' in part:
            try:
                start, end = map(int, part.split('-'))
                res.update(range(start, end + 1))
            except: pass
        else:
            try:
                res.add(int(part))
            except: pass
    return res

def escape_xml(text):
    if not text: return ""
    return str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;").replace("'", "&apos;")

def is_correct_option(block):
    """Kiểm tra xem block (đoạn văn) có chứa dấu hiệu đáp án đúng không (gạch chân hoặc đỏ)"""
    r_nodes = block.getElementsByTagNameNS(W_NS, "r")
    for r in r_nodes:
        # Kiểm tra gạch chân
        rPr_list = r.getElementsByTagNameNS(W_NS, "rPr")
        for rPr in rPr_list:
            u_list = rPr.getElementsByTagNameNS(W_NS, "u")
            if u_list:
                val = u_list[0].getAttributeNS(W_NS, "val")
                if val and val != "none": return True
            
            # Kiểm tra màu đỏ
            color_list = rPr.getElementsByTagNameNS(W_NS, "color")
            if color_list:
                val = color_list[0].getAttributeNS(W_NS, "val")
                # Các mã màu đỏ thường gặp trong Word
                if val and val.upper() in ["FF0000", "RED", "C00000", "FF3333"]: return True
    return False

def extract_short_answer_key(question_blocks):
    key = ""
    clean_blocks = []
    for block in question_blocks:
        txt = get_text(block)
        m = re.match(r'^\s*(?:Đáp án|DA|Lời giải|HD|Hướng dẫn)\s*[:\.]?\s*(.*)', txt, re.IGNORECASE)
        if m:
            key = m.group(1).strip()
            continue
        clean_blocks.append(block)
    return clean_blocks, key

def get_text(block):
    texts = []
    t_nodes = block.getElementsByTagNameNS(W_NS, "t")
    for t in t_nodes:
        if t.firstChild and t.firstChild.nodeValue:
            texts.append(t.firstChild.nodeValue)
    return "".join(texts).strip()

# --- AUTO-SPLIT MERGED OPTIONS LOGIC ---

def split_paragraph_at_text_index(p, split_idx):
    """Chia paragraph p thành 2 paragraph tại vị trí text index"""
    doc = p.ownerDocument
    
    # 1. Map toàn bộ text node và vị trí của nó
    t_nodes = []
    curr_len = 0
    
    def walk_t_nodes(node):
        nonlocal curr_len
        if node.nodeType == node.ELEMENT_NODE and node.localName == 't' and node.namespaceURI == W_NS:
            txt = node.firstChild.nodeValue if node.firstChild else ""
            t_nodes.append({
                "node": node,
                "start": curr_len,
                "end": curr_len + len(txt),
                "text": txt,
                "parent_run": node.parentNode
            })
            curr_len += len(txt)
        elif node.hasChildNodes():
            for child in node.childNodes:
                walk_t_nodes(child)
                
    walk_t_nodes(p)
    
    if split_idx <= 0 or split_idx >= curr_len:
        return None # Không cần split

    # Tìm node chứa điểm cắt
    target_info = None
    for info in t_nodes:
        if info["start"] <= split_idx < info["end"]:
            target_info = info
            break
            
    if not target_info: return None
    
    # 2. Clone paragraph mới
    p_new = p.cloneNode(True)
    p.parentNode.insertBefore(p_new, p.nextSibling)
    
    # 3. Xử lý P cũ (Giữ phần đầu, xóa phần sau)
    # Cần xác định node cắt trong P cũ để xóa các node sau nó
    # Logic đơn giản hóa: Duyệt lại t_nodes của P cũ, cắt text tại target, xóa các t_node sau target
    # Tuy nhiên cấu trúc DOM phức tạp (Run > Text). 
    
    # Giải pháp an toàn hơn: 
    # - P cũ: Cắt text tại split_point. Xóa nội dung text sau đó. (Các run sau đó sẽ rỗng text, nhưng vẫn còn style -> chấp nhận được hoặc cleanup sau)
    # - P mới: Cắt text tại split_point (lấy phần sau). Xóa nội dung text trước đó.
    
    # Xử lý P cũ (Left)
    rel_idx = split_idx - target_info["start"]
    target_info["node"].firstChild.nodeValue = target_info["text"][:rel_idx] # Cắt text
    
    # Xóa nội dung của các text node SAU node cắt trong P cũ
    found_split = False
    
    def clear_text_after(node, stop_node):
        nonlocal found_split
        if node == stop_node:
            found_split = True
            return
        
        if node.nodeType == node.ELEMENT_NODE and node.localName == 't' and node.namespaceURI == W_NS:
            if found_split:
                if node.firstChild: node.firstChild.nodeValue = ""
        
        if node.hasChildNodes():
            for child in node.childNodes:
                clear_text_after(child, stop_node)
                
    clear_text_after(p, target_info["node"])
    
    # Xử lý P mới (Right)
    # Tìm lại node tương ứng trong P mới (do cloneNode)
    # Vì clone hoàn toàn nên cấu trúc y hệt. Ta duyệt tương tự để tìm node đối ứng.
    
    t_nodes_new = []
    def walk_t_nodes_new(node):
        if node.nodeType == node.ELEMENT_NODE and node.localName == 't' and node.namespaceURI == W_NS:
            t_nodes_new.append(node)
        elif node.hasChildNodes():
            for child in node.childNodes:
                walk_t_nodes_new(child)
    
    walk_t_nodes_new(p_new)
    
    # Index của node cắt trong danh sách t_nodes là giống nhau
    target_idx_in_list = t_nodes.index(target_info)
    target_t_new = t_nodes_new[target_idx_in_list]
    
    # Cắt text p mới (Lấy phần sau)
    target_t_new.firstChild.nodeValue = target_info["text"][rel_idx:]
    
    # Xóa nội dung các text node TRƯỚC node cắt trong P mới
    for i in range(target_idx_in_list):
        t_node = t_nodes_new[i]
        if t_node.firstChild: t_node.firstChild.nodeValue = ""
        
    return p_new

def fix_merged_options(dom):
    """Tự động tách các đáp án A. B. C. D. nằm chung 1 dòng thành các dòng riêng"""
    body = dom.getElementsByTagNameNS(W_NS, "body")[0]
    blocks = []
    for child in list(body.childNodes):
        if child.nodeType == child.ELEMENT_NODE and child.localName == "p":
            blocks.append(child)
            
    fixed_count = 0
    
    # Regex tìm B., C., D. nằm giữa dòng (có khoảng trắng phía trước)
    # Group 1: Whitespace, Group 2: Letter (B-D), Group 3: Dot/Paren
    # VD: " ...  B. "
    pattern = re.compile(r'(\s+)([B-D])([\.\)])')
    
    i = 0
    while i < len(blocks):
        block = blocks[i]
        txt = get_text(block)
        
        # Chỉ xử lý nếu dòng này có vẻ là dòng đáp án (chứa A. hoặc a.)
        if not re.match(r'^\s*[A-Da-d][\.\)]', txt):
            i += 1
            continue
            
        # Tìm vị trí cần cắt
        match = pattern.search(txt)
        if match:
            # Vị trí cắt là bắt đầu của chữ cái (B, C, D)
            # match.start(2) là vị trí của ký tự B/C/D
            split_idx = match.start(2)
            
            # Thực hiện tách
            new_block = split_paragraph_at_text_index(block, split_idx)
            
            if new_block:
                # Chèn block mới vào danh sách để duyệt tiếp (vì block mới có thể chứa C, D tiếp)
                blocks.insert(i + 1, new_block)
                fixed_count += 1
                
                # Không tăng i, để vòng lặp sau kiểm tra lại block hiện tại 
                # (thực ra block hiện tại đã bị cắt ngắn, block mới nằm sau)
                # Logic đúng: block hiện tại đã mất phần sau. Block sau (new_block) chứa phần sau.
                # Cần kiểm tra tiếp new_block xem còn C. D. không.
                # Nên ta tăng i để qua block hiện tại, xử lý block kế tiếp (new_block)
                i += 1 
            else:
                i += 1
        else:
            i += 1
            
    return fixed_count

# --- END AUTO-SPLIT LOGIC ---

def set_paragraph_tabs(paragraph, tab_positions):
    doc = paragraph.ownerDocument
    pPr_list = paragraph.getElementsByTagNameNS(W_NS, "pPr")
    if not pPr_list:
        pPr = doc.createElementNS(W_NS, "w:pPr")
        paragraph.insertBefore(pPr, paragraph.firstChild)
    else: pPr = pPr_list[0]
    tabs_list = pPr.getElementsByTagNameNS(W_NS, "tabs")
    for tabs in tabs_list: pPr.removeChild(tabs)
    w_tabs = doc.createElementNS(W_NS, "w:tabs")
    for pos in tab_positions:
        w_tab = doc.createElementNS(W_NS, "w:tab")
        w_tab.setAttributeNS(W_NS, "w:val", "left")
        w_tab.setAttributeNS(W_NS, "w:pos", str(pos))
        w_tabs.appendChild(w_tab)
    pPr.appendChild(w_tabs)

def merge_paragraphs(p_dest, p_src):
    doc = p_dest.ownerDocument
    r_tab = doc.createElementNS(W_NS, "w:r")
    tab = doc.createElementNS(W_NS, "w:tab")
    r_tab.appendChild(tab)
    p_dest.appendChild(r_tab)
    children = []
    for child in p_src.childNodes:
        if child.localName not in ["pPr", "proofErr", "bookmarkStart", "bookmarkEnd"]:
            children.append(child)
    for child in children: p_dest.appendChild(child)
    return p_dest

def format_mcq_layout(question_blocks):
    option_indices = []
    for i, block in enumerate(question_blocks):
        if re.match(r'^\s*[A-D][\.\)]', get_text(block), re.IGNORECASE):
            option_indices.append(i)
    if len(option_indices) != 4: return question_blocks
    opt_blocks = [question_blocks[i] for i in option_indices]
    lengths = [len(get_text(b)) for b in opt_blocks]
    max_len = max(lengths)
    layout_mode = 1
    if max_len < 20: layout_mode = 4
    elif max_len < 45: layout_mode = 2
    else: layout_mode = 1
    if layout_mode == 1: return question_blocks
    new_question_blocks = []
    for i in range(option_indices[0]): new_question_blocks.append(question_blocks[i])
    if layout_mode == 4:
        p_root = opt_blocks[0]
        merge_paragraphs(p_root, opt_blocks[1])
        merge_paragraphs(p_root, opt_blocks[2])
        merge_paragraphs(p_root, opt_blocks[3])
        set_paragraph_tabs(p_root, [3000, 6000, 9000])
        new_question_blocks.append(p_root)
    elif layout_mode == 2:
        row1 = opt_blocks[0]
        merge_paragraphs(row1, opt_blocks[1])
        set_paragraph_tabs(row1, [6000])
        new_question_blocks.append(row1)
        row2 = opt_blocks[2]
        merge_paragraphs(row2, opt_blocks[3])
        set_paragraph_tabs(row2, [6000])
        new_question_blocks.append(row2)
    last_opt_idx = option_indices[-1]
    for i in range(last_opt_idx + 1, len(question_blocks)): new_question_blocks.append(question_blocks[i])
    return new_question_blocks

def style_run_blue_bold(run):
    doc = run.ownerDocument
    rPr_list = run.getElementsByTagNameNS(W_NS, "rPr")
    if rPr_list: rPr = rPr_list[0]
    else:
        rPr = doc.createElementNS(W_NS, "w:rPr")
        run.insertBefore(rPr, run.firstChild)
    color_list = rPr.getElementsByTagNameNS(W_NS, "color")
    if color_list: color_el = color_list[0]
    else:
        color_el = doc.createElementNS(W_NS, "w:color")
        rPr.appendChild(color_el)
    color_el.setAttributeNS(W_NS, "w:val", "0000FF")
    b_list = rPr.getElementsByTagNameNS(W_NS, "b")
    if not b_list:
        b_el = doc.createElementNS(W_NS, "w:b")
        rPr.appendChild(b_el)

def update_mcq_label(paragraph, new_label):
    t_nodes = paragraph.getElementsByTagNameNS(W_NS, "t")
    if not t_nodes: return
    new_letter = new_label[0].upper()
    for i, t in enumerate(t_nodes):
        if not t.firstChild: continue
        txt = t.firstChild.nodeValue
        m = re.match(r'^(\s*)([A-D])([\.\)])?', txt, re.IGNORECASE)
        if not m: continue
        leading_space = m.group(1) or ""
        old_punct = m.group(3) or ""
        after_match = txt[m.end():]
        t.firstChild.nodeValue = leading_space + new_letter + ("." if not old_punct else old_punct) + " " + after_match.strip()
        run = t.parentNode
        if run and run.localName == "r": style_run_blue_bold(run)
        for j in range(i + 1, len(t_nodes)):
            t2 = t_nodes[j]
            if not t2.firstChild: continue
            val2 = t2.firstChild.nodeValue
            if re.match(r'^[\s\.]+$', val2): t2.firstChild.nodeValue = ""
            elif re.match(r'^\.', val2): 
                t2.firstChild.nodeValue = val2[1:]
                break
            else: break
        break

def update_tf_label(paragraph, new_label):
    t_nodes = paragraph.getElementsByTagNameNS(W_NS, "t")
    if not t_nodes: return
    new_letter = new_label[0].lower()
    for i, t in enumerate(t_nodes):
        if not t.firstChild: continue
        txt = t.firstChild.nodeValue
        m = re.match(r'^(\s*)([a-d])(\))?', txt, re.IGNORECASE)
        if not m: continue
        leading_space = m.group(1) or ""
        after_match = txt[m.end():]
        t.firstChild.nodeValue = leading_space + new_letter + ")" + after_match
        run = t.parentNode
        if run and run.localName == "r": style_run_blue_bold(run)
        for j in range(i + 1, len(t_nodes)):
            t2 = t_nodes[j]
            if not t2.firstChild: continue
            val2 = t2.firstChild.nodeValue
            if re.match(r'^[\s\)]+$', val2): t2.firstChild.nodeValue = ""
            elif re.match(r'^\s*\)', val2):
                t2.firstChild.nodeValue = re.sub(r'^\s*\)', '', val2, count=1)
                break
            else: break
        break

def update_question_label(paragraph, new_label):
    t_nodes = paragraph.getElementsByTagNameNS(W_NS, "t")
    if not t_nodes: return
    for i, t in enumerate(t_nodes):
        if not t.firstChild: continue
        txt = t.firstChild.nodeValue
        m = re.match(r'^(\s*)(Câu\s*)(\d+)(\.)?', txt, re.IGNORECASE)
        if not m: continue
        leading_space = m.group(1) or ""
        after_match = txt[m.end():]
        t.firstChild.nodeValue = leading_space + new_label + after_match
        run = t.parentNode
        if run and run.localName == "r": style_run_blue_bold(run)
        for j in range(i + 1, len(t_nodes)):
            t2 = t_nodes[j]
            if not t2.firstChild: continue
            if re.match(r'^[\s0-9\.]*$', t2.firstChild.nodeValue): t2.firstChild.nodeValue = ""
            else: break
        break

def find_part_index(blocks, part_number):
    pattern = re.compile(rf'PHẦN\s*{part_number}\b', re.IGNORECASE)
    for i, block in enumerate(blocks):
        if pattern.search(get_text(block)): return i
    return -1

def parse_questions_in_range(blocks, start, end):
    part_blocks = blocks[start:end]
    items = [] 
    intro = []
    i = 0
    while i < len(part_blocks):
        text = get_text(part_blocks[i])
        if re.match(r'^Câu\s*\d+\b', text, re.IGNORECASE): break
        if "@BẮT ĐẦU DÙNG CHUNG@" in text.upper(): break
        intro.append(part_blocks[i])
        i += 1
    while i < len(part_blocks):
        block = part_blocks[i]
        text = get_text(block)
        if "@BẮT ĐẦU DÙNG CHUNG@" in text.upper():
            cluster_header = []
            cluster_questions = []
            i += 1 
            while i < len(part_blocks):
                b_curr = part_blocks[i]
                t_curr = get_text(b_curr)
                if "@KẾT THÚC DÙNG CHUNG@" in t_curr.upper():
                    i += 1 
                    break
                if re.match(r'^Câu\s*\d+\b', t_curr, re.IGNORECASE):
                    one_q = [b_curr]
                    i += 1
                    while i < len(part_blocks):
                        b_next = part_blocks[i]
                        t_next = get_text(b_next)
                        if "@KẾT THÚC DÙNG CHUNG@" in t_next.upper(): break
                        if re.match(r'^Câu\s*\d+\b', t_next, re.IGNORECASE): break
                        one_q.append(b_next)
                        i += 1
                    cluster_questions.append(one_q)
                else:
                    if cluster_questions: cluster_questions[-1].append(b_curr)
                    else: cluster_header.append(b_curr)
                    i += 1
            items.append({"type": "cluster", "header": cluster_header, "questions": cluster_questions})
            continue
        if re.match(r'^Câu\s*\d+\b', text, re.IGNORECASE):
            group = [block]
            i += 1
            while i < len(part_blocks):
                t2 = get_text(part_blocks[i])
                if re.match(r'^Câu\s*\d+\b', t2, re.IGNORECASE): break
                if "@BẮT ĐẦU DÙNG CHUNG@" in t2.upper(): break
                if re.match(r'^PHẦN\s*\d\b', t2, re.IGNORECASE): break
                group.append(part_blocks[i])
                i += 1
            items.append({"type": "question", "blocks": group})
        else:
            if items and items[-1]["type"] == "question": items[-1]["blocks"].append(block)
            elif not items: intro.append(block)
            i += 1
    return intro, items

def shuffle_array(arr):
    out = arr.copy()
    for i in range(len(out) - 1, 0, -1):
        j = random.randint(0, i)
        out[i], out[j] = out[j], out[i]
    return out

# --- NEW: VALIDATION FUNCTION WITH AUTO-FIX ---
def check_exam_structure(file_bytes):
    """Kiểm tra cấu trúc đề (Phần 1) trước khi trộn, có tự động sửa dòng"""
    input_buffer = io.BytesIO(file_bytes)
    messages = []
    is_valid = True
    
    try:
        with zipfile.ZipFile(input_buffer, 'r') as zin:
            doc_xml = zin.read("word/document.xml").decode('utf-8')
            dom = minidom.parseString(doc_xml)
            
            # 1. AUTO FIX: Tách các đáp án dính liền
            fixed_cnt = fix_merged_options(dom)
            if fixed_cnt > 0:
                messages.append(f"✅ Đã tự động tách {fixed_cnt} dòng đáp án bị dính liền.")
            
            body = dom.getElementsByTagNameNS(W_NS, "body")[0]
            blocks = []
            for child in list(body.childNodes):
                if child.nodeType == child.ELEMENT_NODE and child.localName in ["p", "tbl"]:
                    blocks.append(child)
            
            # Tìm phần 1
            p1 = find_part_index(blocks, 1)
            p2 = find_part_index(blocks, 2)
            
            start = 0
            end = len(blocks)
            
            if p1 >= 0:
                start = p1 + 1
                if p2 >= 0: end = p2
            elif p2 >= 0:
                end = p2
            
            _, items = parse_questions_in_range(blocks, start, end)
            
            if not items:
                messages.append("⚠️ Không tìm thấy câu hỏi trắc nghiệm nào (Phần 1). Hãy kiểm tra lại từ khóa 'Câu ...'.")
                return False, messages

            q_count = 0
            for item in items:
                if item["type"] == "question":
                    q_count += 1
                    q_blocks = item["blocks"]
                    
                    # 1. Kiểm tra số lượng đáp án
                    opt_blocks = []
                    correct_count = 0
                    
                    q_text_header = get_text(q_blocks[0])
                    
                    for b in q_blocks:
                        txt = get_text(b)
                        if re.match(r'^\s*[A-D][\.\)]', txt, re.IGNORECASE):
                            opt_blocks.append(b)
                            if is_correct_option(b):
                                correct_count += 1
                    
                    # Cảnh báo nếu không đủ 4 đáp án
                    if len(opt_blocks) < 4:
                        is_valid = False
                        messages.append(f"❌ {q_text_header[:10]}...: Chỉ tìm thấy {len(opt_blocks)} đáp án (A,B,C,D). Có thể do định dạng tab chưa chuẩn.")
                    
                    # 2. Kiểm tra đáp án đúng
                    if correct_count == 0:
                        is_valid = False
                        messages.append(f"❌ {q_text_header[:10]}...: Chưa chọn đáp án đúng (Chưa gạch chân hoặc tô đỏ).")
                    elif correct_count > 1:
                        is_valid = False
                        messages.append(f"❌ {q_text_header[:10]}...: Có {correct_count} đáp án được đánh dấu đúng (Chỉ được phép có 1).")
                
                elif item["type"] == "cluster":
                     messages.append(f"ℹ️ Phát hiện nhóm câu hỏi dùng chung. Hệ thống chưa hỗ trợ kiểm tra chi tiết bên trong nhóm này, nhưng vẫn sẽ trộn bình thường.")

            if q_count == 0:
                 messages.append("⚠️ Không tìm thấy câu hỏi nào bắt đầu bằng 'Câu ...'.")
                 is_valid = False

    except Exception as e:
        return False, [f"Lỗi khi đọc file: {str(e)}"]

    return is_valid, messages

# --- HELPER FUNCTIONS FOR WORD XML GENERATION ---
def create_header_xml(doc, info):
    so_gd = escape_xml(info.get("so_gd", "").upper())
    truong = escape_xml(info.get("truong", ""))
    ky_thi = escape_xml(info.get("ky_thi", "").upper())
    mon_thi = escape_xml(info.get("mon_thi", "").upper())
    thoi_gian = escape_xml(info.get("thoi_gian", ""))
    nam_hoc = escape_xml(info.get("nam_hoc", ""))
    xml_str = f"""
    <w:tbl xmlns:w="{W_NS}">
        <w:tblPr>
            <w:tblW w:w="0" w:type="auto"/>
            <w:jc w:val="center"/>
            <w:tblBorders>
                <w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>
                <w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>
                <w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>
                <w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>
                <w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>
                <w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>
            </w:tblBorders>
        </w:tblPr>
        <w:tr>
            <w:tc>
                <w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr>
                <w:p>
                    <w:pPr><w:jc w:val="center"/></w:pPr>
                    <w:r><w:rPr><w:b/></w:rPr><w:t>{so_gd}</w:t></w:r>
                </w:p>
                <w:p>
                    <w:pPr><w:jc w:val="center"/></w:pPr>
                    <w:r><w:rPr><w:b/></w:rPr><w:t>{truong}</w:t></w:r>
                </w:p>
                <w:p>
                    <w:pPr><w:jc w:val="center"/></w:pPr>
                    <w:r><w:t>------------------</w:t></w:r>
                </w:p>
            </w:tc>
            <w:tc>
                <w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr>
                <w:p>
                    <w:pPr><w:jc w:val="center"/></w:pPr>
                    <w:r><w:rPr><w:b/></w:rPr><w:t>{ky_thi}</w:t></w:r>
                </w:p>
                <w:p>
                    <w:pPr><w:jc w:val="center"/></w:pPr>
                    <w:r><w:rPr><w:b/></w:rPr><w:t>MÔN: {mon_thi}</w:t></w:r>
                </w:p>
                <w:p>
                    <w:pPr><w:jc w:val="center"/></w:pPr>
                    <w:r><w:t>Thời gian làm bài: {thoi_gian}</w:t></w:r>
                </w:p>
                 <w:p>
                    <w:pPr><w:jc w:val="center"/></w:pPr>
                    <w:r><w:t>(Năm học: {nam_hoc})</w:t></w:r>
                </w:p>
            </w:tc>
        </w:tr>
    </w:tbl>
    """
    return minidom.parseString(xml_str).documentElement

def create_footer_xml_content(ma_de):
    xml_str = f"""
    <w:ftr xmlns:w="{W_NS}">
        <w:p>
            <w:pPr>
                <w:pStyle w:val="Footer"/>
                <w:jc w:val="right"/>
                <w:pBdr>
                    <w:top w:val="single" w:sz="6" w:space="1" w:color="auto"/>
                </w:pBdr>
            </w:pPr>
            <w:r>
                <w:t xml:space="preserve">Mã đề {ma_de} - Trang </w:t>
            </w:r>
            <w:fldSimple w:instr="PAGE"/>
        </w:p>
    </w:ftr>
    """
    return xml_str.strip()

def add_header_to_body(dom, body, header_info):
    if not header_info.get("enable", False): return
    try:
        tbl_node = create_header_xml(dom, header_info)
        if body.firstChild: body.insertBefore(tbl_node, body.firstChild)
        else: body.appendChild(tbl_node)
        p_empty = dom.createElementNS(W_NS, "w:p")
        if body.childNodes.length > 1: body.insertBefore(p_empty, body.childNodes[1])
    except: pass

def relabel_mcq_options(question_blocks):
    letters = ["A", "B", "C", "D"]
    count = 0
    for block in question_blocks:
        if re.match(r'^\s*[A-D][\.\)]', get_text(block), re.IGNORECASE):
            l = letters[count] if count < 4 else "D"
            update_mcq_label(block, f"{l}.")
            count += 1

def relabel_tf_options(question_blocks):
    letters = ["a", "b", "c", "d"]
    count = 0
    for block in question_blocks:
        if re.match(r'^\s*[a-d]\)', get_text(block), re.IGNORECASE):
            l = letters[count] if count < 4 else "d"
            update_tf_label(block, f"{l})")
            count += 1

def shuffle_mcq_options(question_blocks, allow_shuffle=True):
    indices = []
    correct_indices_before = []
    for i, block in enumerate(question_blocks):
        if re.match(r'^\s*[A-D][\.\)]', get_text(block), re.IGNORECASE):
            indices.append(i)
            if is_correct_option(block): correct_indices_before.append(i)
    if len(indices) < 2: return question_blocks, ""
    options = [question_blocks[idx] for idx in indices]
    perm = list(range(len(options)))
    if allow_shuffle: random.shuffle(perm)
    shuffled_options = [options[p] for p in perm]
    new_correct_char = ""
    if correct_indices_before:
        orig_correct_idx_in_options = -1
        for k, val in enumerate(indices):
            if val == correct_indices_before[0]:
                orig_correct_idx_in_options = k
                break
        if orig_correct_idx_in_options != -1:
            for new_pos, old_pos in enumerate(perm):
                if old_pos == orig_correct_idx_in_options:
                    letters = ["A", "B", "C", "D", "E", "F"]
                    if new_pos < len(letters): new_correct_char = letters[new_pos]
                    break
    min_idx, max_idx = min(indices), max(indices)
    before = question_blocks[:min_idx]
    after = question_blocks[max_idx + 1:]
    return before + shuffled_options + after, new_correct_char

def shuffle_tf_options(question_blocks, allow_shuffle=True):
    option_indices = {}
    for i, block in enumerate(question_blocks):
        m = re.match(r'^\s*([a-d])\)', get_text(block), re.IGNORECASE)
        if m: option_indices[m.group(1).lower()] = i
    abc_idx = [option_indices.get(k) for k in ["a", "b", "c"] if option_indices.get(k) is not None]
    if len(abc_idx) < 2: return question_blocks, ["", "", "", ""]
    abc_nodes = [question_blocks[idx] for idx in abc_idx]
    if allow_shuffle: shuffled_abc = shuffle_array(abc_nodes)
    else: shuffled_abc = abc_nodes.copy()
    all_vals = [v for v in option_indices.values() if v is not None]
    min_idx, max_idx = min(all_vals), max(all_vals)
    before = question_blocks[:min_idx]
    after = question_blocks[max_idx + 1:]
    d_node = question_blocks[option_indices["d"]] if "d" in option_indices else None
    middle = shuffled_abc.copy()
    if d_node: middle.append(d_node)
    current_key_status = []
    for block in middle:
        status = "D" if is_correct_option(block) else "S"
        current_key_status.append(status)
    return before + middle + after, current_key_status

def process_single_question_logic(q, part_type, allow_shuffle_opt):
    new_block = []
    key = ""
    if part_type == "PHAN1":
        new_block, key = shuffle_mcq_options(q, allow_shuffle_opt)
    elif part_type == "PHAN2":
        new_block, key = shuffle_tf_options(q, allow_shuffle_opt)
    elif part_type == "PHAN3":
        new_block, key = extract_short_answer_key(q)
    else:
        new_block = q.copy()
    return new_block, key

def process_part(blocks, start, end, part_type, global_q_idx_start, config):
    intro, items = parse_questions_in_range(blocks, start, end)
    processed_items = []
    current_q_counter = global_q_idx_start
    fixed_pos_set = config.get("fixed_pos_set", set())
    fixed_opt_set = config.get("fixed_opt_set", set())
    fix_group_pos = config.get("fix_group_pos", False)
    
    for item in items:
        if item["type"] == "question":
            q_idx = current_q_counter + 1
            allow_opt = config.get("shuffle_opt_global", True)
            if q_idx in fixed_opt_set: allow_opt = False
            new_q, key = process_single_question_logic(item["blocks"], part_type, allow_opt)
            processed_items.append({"type": "question", "blocks": new_q, "keys": [key], "original_idx": q_idx})
            current_q_counter += 1
        elif item["type"] == "cluster":
            header = item["header"]
            sub_qs = item["questions"]
            sub_items_data = []
            sub_keys = []
            for sub_q_blocks in sub_qs:
                q_idx = current_q_counter + 1
                allow_opt = config.get("shuffle_opt_global", True)
                if q_idx in fixed_opt_set: allow_opt = False
                new_q, key = process_single_question_logic(sub_q_blocks, part_type, allow_opt)
                sub_items_data.append((new_q, key))
                current_q_counter += 1
            if config.get("shuffle_pos_global", True): random.shuffle(sub_items_data)
            cluster_final_blocks = header.copy()
            for sq, k in sub_items_data:
                cluster_final_blocks.extend(sq)
                sub_keys.append(k)
            processed_items.append({
                "type": "cluster",
                "blocks": cluster_final_blocks,
                "keys": sub_keys,
                "original_idx": current_q_counter - len(sub_qs) + 1 
            })

    fixed_map = {}
    movable = []
    for i, item_data in enumerate(processed_items):
        is_fixed = False
        if not config.get("shuffle_pos_global", True): is_fixed = True
        if item_data["original_idx"] in fixed_pos_set: is_fixed = True
        if fix_group_pos and item_data["type"] == "cluster": is_fixed = True
        if is_fixed: fixed_map[i] = item_data
        else: movable.append(item_data)
    random.shuffle(movable)
    final_blocks = intro.copy()
    final_keys = []
    movable_idx = 0
    total_items = len(processed_items)
    final_item_list = []
    for i in range(total_items):
        if i in fixed_map: final_item_list.append(fixed_map[i])
        else:
            final_item_list.append(movable[movable_idx])
            movable_idx += 1
    
    q_counter = 0
    def flush_q_group(group, p_type):
        if not group: return []
        if p_type == "PHAN1":
            relabel_mcq_options(group)
            return format_mcq_layout(group)
        elif p_type == "PHAN2":
            relabel_tf_options(group)
            return group
        return group

    for item in final_item_list:
        final_keys.extend(item["keys"])
        if item["type"] == "question":
            q_blocks = item["blocks"]
            if q_blocks:
                q_counter += 1
                update_question_label(q_blocks[0], f"Câu {q_counter}.")
                formatted_blocks = flush_q_group(q_blocks, part_type)
                final_blocks.extend(formatted_blocks)
        elif item["type"] == "cluster":
            c_blocks = item["blocks"]
            current_sub_q = []
            for blk in c_blocks:
                txt = get_text(blk)
                if re.match(r'^Câu\s*\d+\b', txt):
                    if current_sub_q:
                        final_blocks.extend(flush_q_group(current_sub_q, part_type))
                        current_sub_q = []
                    q_counter += 1
                    update_question_label(blk, f"Câu {q_counter}.")
                    current_sub_q.append(blk)
                else:
                    if current_sub_q: current_sub_q.append(blk)
                    else: final_blocks.append(blk)
            if current_sub_q: final_blocks.extend(flush_q_group(current_sub_q, part_type))
    return final_blocks, final_keys

def shuffle_docx_logic(file_bytes, shuffle_mode, header_info, ma_de_str="", config=None):
    if config is None: config = {}
    input_buffer = io.BytesIO(file_bytes)
    keys_by_part = {}
    with zipfile.ZipFile(input_buffer, 'r') as zin:
        doc_xml = zin.read("word/document.xml").decode('utf-8')
        dom = minidom.parseString(doc_xml)
        
        # --- AUTO FIX: Tách các đáp án dính liền trước khi trộn ---
        fix_merged_options(dom)
        
        body = dom.getElementsByTagNameNS(W_NS, "body")[0]
        blocks = []
        other_nodes = []
        for child in list(body.childNodes):
            if child.nodeType == child.ELEMENT_NODE and child.localName in ["p", "tbl"]: blocks.append(child)
            elif child.nodeType == child.ELEMENT_NODE: other_nodes.append(child)
            body.removeChild(child)
        new_blocks = []
        p1 = find_part_index(blocks, 1)
        p2 = find_part_index(blocks, 2)
        p3 = find_part_index(blocks, 3)
        if shuffle_mode != "auto" or (p1 == -1 and p2 == -1 and p3 == -1):
            p_type = "PHAN1" if shuffle_mode == "mcq" or shuffle_mode == "auto" else "PHAN2"
            nb, k = process_part(blocks, 0, len(blocks), p_type, 0, config)
            new_blocks = nb
            keys_by_part['MCQ_ALL' if p_type == "PHAN1" else 'TF_ALL'] = k
        else:
            cursor = 0
            current_global_q_idx = 0 
            if p1 >= 0:
                new_blocks.extend(blocks[cursor:p1+1])
                cursor = p1 + 1
                end1 = p2 if p2 >= 0 else len(blocks)
                nb, k = process_part(blocks, cursor, end1, "PHAN1", current_global_q_idx, config)
                new_blocks.extend(nb)
                keys_by_part['PHAN1'] = k
                current_global_q_idx += len(k)
                cursor = end1
            if p2 >= 0:
                new_blocks.append(blocks[p2])
                cursor = p2 + 1
                end2 = p3 if p3 >= 0 else len(blocks)
                nb, k = process_part(blocks, cursor, end2, "PHAN2", current_global_q_idx, config)
                new_blocks.extend(nb)
                keys_by_part['PHAN2'] = k
                current_global_q_idx += len(k)
                cursor = end2
            if p3 >= 0:
                new_blocks.append(blocks[p3])
                cursor = p3 + 1
                nb, k = process_part(blocks, cursor, len(blocks), "PHAN3", current_global_q_idx, config)
                new_blocks.extend(nb)
                keys_by_part['PHAN3'] = k

        if ma_de_str:
            p_ma = dom.createElementNS(W_NS, "w:p")
            p_ma_pr = dom.createElementNS(W_NS, "w:pPr")
            jc = dom.createElementNS(W_NS, "w:jc")
            jc.setAttributeNS(W_NS, "w:val", "right")
            p_ma_pr.appendChild(jc)
            p_ma.appendChild(p_ma_pr)
            r = dom.createElementNS(W_NS, "w:r")
            t = dom.createElementNS(W_NS, "w:t")
            rPr = dom.createElementNS(W_NS, "w:rPr")
            b = dom.createElementNS(W_NS, "w:b")
            rPr.appendChild(b)
            r.appendChild(rPr)
            t.appendChild(dom.createTextNode(f"Mã đề: {ma_de_str}"))
            r.appendChild(t)
            p_ma.appendChild(r)
            add_header_to_body(dom, body, header_info)
            if header_info.get("enable"):
                if body.childNodes.length > 1: body.insertBefore(p_ma, body.childNodes[1])
                else: body.appendChild(p_ma)
            else:
                if body.firstChild: body.insertBefore(p_ma, body.firstChild)
                else: body.appendChild(p_ma)
        else:
            add_header_to_body(dom, body, header_info)

        footer_rel_id = "rIdFooterNew"
        footer_fname = "word/footer_new.xml"
        sectPrs = body.getElementsByTagNameNS(W_NS, "sectPr")
        if sectPrs: sectPr = sectPrs[-1]
        else:
            sectPr = dom.createElementNS(W_NS, "w:sectPr")
            body.appendChild(sectPr)
        for child in list(sectPr.childNodes):
            if child.localName == "footerReference": sectPr.removeChild(child)
        fr = dom.createElementNS(W_NS, "w:footerReference")
        fr.setAttributeNS(W_NS, "w:type", "default")
        fr.setAttributeNS(R_NS, "r:id", footer_rel_id)
        sectPr.appendChild(fr)

        for b in new_blocks: body.appendChild(b)
        for n in other_nodes: body.appendChild(n)
        
        output_buffer = io.BytesIO()
        with zipfile.ZipFile(output_buffer, 'w', zipfile.ZIP_DEFLATED) as zout:
            footer_xml = create_footer_xml_content(ma_de_str)
            zout.writestr(footer_fname, footer_xml.encode('utf-8'))
            for item in zin.infolist():
                if item.filename == "word/document.xml":
                    zout.writestr(item, dom.toxml().encode('utf-8'))
                elif item.filename == "[Content_Types].xml":
                    ct_xml = zin.read(item).decode('utf-8')
                    ct_dom = minidom.parseString(ct_xml)
                    types = ct_dom.getElementsByTagName("Types")[0]
                    ov = ct_dom.createElement("Override")
                    ov.setAttribute("PartName", "/word/footer_new.xml")
                    ov.setAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml")
                    types.appendChild(ov)
                    zout.writestr(item, ct_dom.toxml().encode('utf-8'))
                elif item.filename == "word/_rels/document.xml.rels":
                    rels_xml = zin.read(item).decode('utf-8')
                    rels_dom = minidom.parseString(rels_xml)
                    relationships = rels_dom.getElementsByTagName("Relationships")[0]
                    rel = rels_dom.createElement("Relationship")
                    rel.setAttribute("Id", footer_rel_id)
                    rel.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer")
                    rel.setAttribute("Target", "footer_new.xml")
                    relationships.appendChild(rel)
                    zout.writestr(item, rels_dom.toxml().encode('utf-8'))
                else:
                    zout.writestr(item, zin.read(item.filename))
        return output_buffer.getvalue(), keys_by_part

def generate_real_excel_xlsx(all_answers_dict):
    ma_des = sorted(list(all_answers_dict.keys()))
    if not ma_des: return b""
    headers = ["Đề \\ Câu"]
    headers.extend([str(i) for i in range(1, 41)])
    for q in range(1, 9):
        for char in ['a', 'b', 'c', 'd']: headers.append(f"{q}{char}")
    headers.extend([str(i) for i in range(1, 7)])
    rows_data = []
    for md in ma_des:
        row = [str(md)]
        keys = all_answers_dict[md]
        mcq_list = []
        if 'PHAN1' in keys: mcq_list = keys['PHAN1']
        elif 'MCQ_ALL' in keys: mcq_list = keys['MCQ_ALL']
        row.extend((mcq_list + [""] * 40)[:40])
        tf_data = []
        if 'PHAN2' in keys: tf_data = keys['PHAN2']
        elif 'TF_ALL' in keys: tf_data = keys['TF_ALL']
        tf_flat = []
        for i in range(8):
            if i < len(tf_data): tf_flat.extend((tf_data[i] + [""] * 4)[:4])
            else: tf_flat.extend(["", "", "", ""])
        row.extend(tf_flat)
        sa_list = []
        if 'PHAN3' in keys: sa_list = keys['PHAN3']
        row.extend((sa_list + [""] * 6)[:6])
        rows_data.append(row)
    
    df = pd.DataFrame(rows_data, columns=headers)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()

def create_summary_table_xml(all_answers_dict):
    ma_des = sorted(list(all_answers_dict.keys()))
    if not ma_des: return None
    mcq_keys_map = {}
    tf_keys_map = {}
    sa_keys_map = {}
    for md in ma_des:
        k = all_answers_dict[md]
        if 'PHAN1' in k: mcq_keys_map[md] = k['PHAN1']
        elif 'MCQ_ALL' in k: mcq_keys_map[md] = k['MCQ_ALL']
        if 'PHAN2' in k: tf_keys_map[md] = k['PHAN2']
        elif 'TF_ALL' in k: tf_keys_map[md] = k['TF_ALL']
        if 'PHAN3' in k: sa_keys_map[md] = k['PHAN3']
    def make_p(text, bold=False, align='center', size=None):
        sz_tag = f'<w:sz w:val="{size}"/>' if size else ''
        b_tag = '<w:b/>' if bold else ''
        safe_text = str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        return f'<w:p><w:pPr><w:jc w:val="{align}"/></w:pPr><w:r><w:rPr>{b_tag}{sz_tag}</w:rPr><w:t>{safe_text}</w:t></w:r></w:p>'
    def make_tc(content, width=None):
        w_tag = f'<w:tcW w:w="{width}" w:type="dxa"/>' if width else '<w:tcW w:w="0" w:type="auto"/>'
        return f'<w:tc><w:tcPr>{w_tag}</w:tcPr>{content}</w:tc>'
    body_content = ""
    if mcq_keys_map:
        num_mcq = len(mcq_keys_map[ma_des[0]])
        body_content += make_p("PHẦN I: TRẮC NGHIỆM", bold=True, align='left', size='28')
        row_cells = make_tc(make_p("Câu \\ Mã", bold=True), width=1200)
        for md in ma_des: row_cells += make_tc(make_p(str(md), bold=True), width=800)
        tbl1_rows = f'<w:tr>{row_cells}</w:tr>'
        for i in range(num_mcq):
            row_cells = make_tc(make_p(str(i+1), bold=True))
            for md in ma_des:
                ans = mcq_keys_map[md][i] if i < len(mcq_keys_map[md]) else ""
                row_cells += make_tc(make_p(ans))
            tbl1_rows += f'<w:tr>{row_cells}</w:tr>'
        body_content += f'<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders></w:tblPr>{tbl1_rows}</w:tbl><w:p/>'
    if tf_keys_map:
        body_content += make_p("PHẦN II: ĐÚNG SAI", bold=True, align='left', size='28')
        row_cells = ""
        headers = ["Mã đề", "Câu", "Ý a", "Ý b", "Ý c", "Ý d"]
        widths = [1000, 800, 800, 800, 800, 800]
        for idx, h in enumerate(headers): row_cells += make_tc(make_p(h, bold=True), width=widths[idx])
        tbl2_rows = f'<w:tr>{row_cells}</w:tr>'
        for md in ma_des:
            tf_data = tf_keys_map[md]
            for i, ans_list in enumerate(tf_data):
                md_text = str(md)
                row_cells = make_tc(make_p(md_text)) + make_tc(make_p(str(i+1), bold=True))
                for char_idx in range(4):
                    val = ans_list[char_idx] if char_idx < len(ans_list) else ""
                    row_cells += make_tc(make_p(val))
                tbl2_rows += f'<w:tr>{row_cells}</w:tr>'
        body_content += f'<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders></w:tblPr>{tbl2_rows}</w:tbl><w:p/>'
    if sa_keys_map:
        body_content += make_p("PHẦN III: TRẢ LỜI NGẮN", bold=True, align='left', size='28')
        row_cells = make_tc(make_p("Câu \\ Mã", bold=True), width=1200)
        for md in ma_des: row_cells += make_tc(make_p(str(md), bold=True), width=1500)
        tbl3_rows = f'<w:tr>{row_cells}</w:tr>'
        num_sa = len(sa_keys_map[ma_des[0]])
        for i in range(num_sa):
            row_cells = make_tc(make_p(str(i+1), bold=True))
            for md in ma_des:
                ans = sa_keys_map[md][i] if i < len(sa_keys_map[md]) else ""
                row_cells += make_tc(make_p(ans))
            tbl3_rows += f'<w:tr>{row_cells}</w:tr>'
        body_content += f'<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders></w:tblPr>{tbl3_rows}</w:tbl>'
    doc_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:document xmlns:w="{W_NS}">
        <w:body>
            <w:p><w:pPr><w:jc w:val="center"/><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="32"/></w:rPr><w:t>BẢNG ĐÁP ÁN TỔNG HỢP</w:t></w:r></w:p>
            {body_content}
        </w:body>
    </w:document>
    """
    return doc_xml

def generate_summary_docx(file_bytes, all_answers_dict):
    input_buffer = io.BytesIO(file_bytes)
    output_buffer = io.BytesIO()
    table_xml_str = create_summary_table_xml(all_answers_dict)
    if not table_xml_str: return io.BytesIO(b"") 
    with zipfile.ZipFile(input_buffer, 'r') as zin:
        with zipfile.ZipFile(output_buffer, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "word/document.xml":
                    zout.writestr(item, table_xml_str.encode('utf-8'))
                else:
                    zout.writestr(item, zin.read(item.filename))
    return output_buffer.getvalue()


# ==================== PHẦN 2: GIAO DIỆN WEB (STREAMLIT) ====================

st.set_page_config(page_title="Trộn Đề Word Pro - AIOMT Online", layout="wide", page_icon="📚")

# --- HEADER & AUTHOR INFO ---
st.title("🎓 HỆ THỐNG TRỘN ĐỀ TRẮC NGHIỆM THÔNG MINH")
st.markdown("### 🚀 Giải pháp trộn đề Word chuyên nghiệp")

col_info, col_link = st.columns([2, 1])

with col_info:
    st.markdown("""   
    **👤 Tác giả:** Nguyễn Thị Thanh Vân   
    
    **📱 Zalo:** 0972.777.872     
    
    **🏫 Đơn vị:** Trường THCS Tây Phú       
   
    """)

with col_link:
    st.link_button("📥 Tải Đề Mẫu Chuẩn (Word)", "https://docs.google.com/document/d/1lCSNGQgulPxcuu3QDEk24pDMPjahXDR_/edit?usp=sharing&ouid=102049743266128652284&rtpof=true&sd=true", help="Bấm để xem và tải file mẫu định dạng chuẩn")

st.markdown("---")

# --- SIDEBAR CONFIG ---
with st.sidebar:
    st.header("⚙️ Cấu Hình")
    
    st.subheader("1. Thông tin Tiêu đề (Header)")
    use_header = st.checkbox("Thêm bảng tiêu đề", value=True)
    so_gd = st.text_input("Xã (SởGD&ĐT)", "UBND XÃ TÂY PHÚ", disabled=not use_header)
    truong = st.text_input("Trường", "TRƯỜNG THCS TÂY PHÚ", disabled=not use_header)
    ky_thi = st.text_input("Kỳ Thi", "ĐỀ KIỂM TRA CUỐI KỲ I", disabled=not use_header)
    mon_thi = st.text_input("Môn Thi", "TOÁN 9", disabled=not use_header)
    thoi_gian = st.text_input("Thời gian", "90 phút", disabled=not use_header)
    nam_hoc = st.text_input("Năm học", "2025 - 2026", disabled=not use_header)

    st.subheader("2. Tùy chọn Trộn")
    shuffle_pos = st.checkbox("Trộn vị trí Câu hỏi", value=True)
    shuffle_opt = st.checkbox("Trộn vị trí Đáp án (A,B,C,D)", value=True)
    fix_group_pos = st.checkbox("Cố định nhóm câu hỏi dùng chung", value=True)

    st.subheader("3. Mã đề")
    ma_de_mode = st.radio("Cách tạo mã đề:", ["Tự động (Ngẫu nhiên)", "Tự nhập"], index=0)
    
    ma_de_list = []
    if ma_de_mode == "Tự động (Ngẫu nhiên)":
        num_ver = st.number_input("Số lượng đề muốn tạo:", min_value=1, max_value=50, value=4)
        start_code = 101
        ma_de_list = [str(start_code + i) for i in range(num_ver)]
    else:
        manual_str = st.text_input("Nhập mã đề (cách nhau dấu phẩy):", "101, 102, 103")
        if manual_str:
            ma_de_list = [s.strip() for s in manual_str.split(',') if s.strip()]

    st.subheader("4. Cố định (Nâng cao)")
    fixed_pos_str = st.text_input("Câu hỏi KHÔNG trộn vị trí (VD: 1, 40):")
    fixed_opt_str = st.text_input("Câu hỏi KHÔNG trộn đáp án (VD: 1-5):")

# --- MAIN CONTENT ---

uploaded_file = st.file_uploader("📂 Chọn file Word (.docx) đề gốc", type=["docx"])

if uploaded_file is not None:
    st.success(f"Đã tải lên: {uploaded_file.name}")
    
    # --- CHECK BUTTON ---
    col_check, col_run = st.columns([1, 1])
    
    with col_check:
        if st.button("🔍 KIỂM TRA CẤU TRÚC ĐỀ", type="secondary", use_container_width=True):
            with st.spinner("Đang phân tích cấu trúc đề..."):
                is_valid, messages = check_exam_structure(uploaded_file.getvalue())
                if is_valid and not messages:
                    st.success("✅ ĐỀ BẠN CHUẨN! Hãy tiến hành trộn đề.")
                elif is_valid and messages:
                    # Check if auto-fix happened
                    if any("Đã tự động tách" in msg for msg in messages):
                        st.success("✅ Đã tự động sửa lỗi định dạng! Đề bây giờ đã hợp lệ.")
                        for msg in messages:
                            st.write(msg)
                    else:
                        st.warning("⚠️ Đề có thể trộn được, nhưng có một số lưu ý:")
                        for msg in messages:
                            st.write(msg)
                else:
                    st.error("❌ PHÁT HIỆN LỖI CẤU TRÚC (Phần 1):")
                    for msg in messages:
                        st.write(msg)
                    st.info("💡 Gợi ý: Hãy sửa lại các lỗi trên trong file Word rồi tải lên lại.")

    with col_run:
        if st.button("🚀 BẮT ĐẦU TRỘN ĐỀ", type="primary", use_container_width=True):
            with st.spinner("Đang xử lý trộn đề..."):
                try:
                    # Đọc file upload
                    file_bytes = uploaded_file.getvalue()
                    
                    # Cấu hình
                    header_info = {
                        "enable": use_header,
                        "so_gd": so_gd, "truong": truong,
                        "ky_thi": ky_thi, "mon_thi": mon_thi,
                        "thoi_gian": thoi_gian, "nam_hoc": nam_hoc
                    }
                    
                    config = {
                        "shuffle_pos_global": shuffle_pos,
                        "shuffle_opt_global": shuffle_opt,
                        "fixed_pos_set": parse_range_string(fixed_pos_str),
                        "fixed_opt_set": parse_range_string(fixed_opt_str),
                        "fix_group_pos": fix_group_pos
                    }
                    
                    all_answers_summary = {}
                    zip_buffer = io.BytesIO()
                    
                    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zout:
                        # Trộn từng đề
                        for ma_de in ma_de_list:
                            out_bytes, keys_by_part = shuffle_docx_logic(file_bytes, "auto", header_info, ma_de, config)
                            all_answers_summary[ma_de] = keys_by_part
                            fname = f"De_Tron_Ma_{ma_de}.docx"
                            zout.writestr(fname, out_bytes)
                        
                        # Tạo file tổng hợp
                        try:
                            summary_bytes = generate_summary_docx(file_bytes, all_answers_summary)
                            zout.writestr("Dap_an_tong_hop.docx", summary_bytes)
                        except Exception as e:
                            st.error(f"Lỗi tạo file Word đáp án: {e}")

                        try:
                            excel_bytes = generate_real_excel_xlsx(all_answers_summary)
                            zout.writestr("Dap_an_Excel_Chuan.xlsx", excel_bytes)
                        except Exception as e:
                            st.error(f"Lỗi tạo file Excel: {e}")
                    
                    # Hoàn tất
                    st.success("✅ Đã trộn xong! Tải file kết quả bên dưới.")
                    
                    btn = st.download_button(
                        label="📥 TẢI VỀ FILE KẾT QUẢ (.ZIP)",
                        data=zip_buffer.getvalue(),
                        file_name="Ket_qua_tron_de.zip",
                        mime="application/zip"
                    )
                    
                except Exception as e:
                    st.error(f"Có lỗi xảy ra: {str(e)}")
else:
    st.info("👈 Vui lòng tải lên file đề gốc (.docx) để bắt đầu.")
    st.markdown("""
    **Hướng dẫn:**
    1. Chuẩn bị file Word đề thi trắc nghiệm theo định dạng chuẩn (Xem file mẫu ở trên).
    2. Đáp án đúng cần được **Gạch chân** hoặc **Tô đỏ**.
    3. Tải file lên và bấm nút **"Kiểm tra cấu trúc đề"** để rà soát lỗi.
    4. Bấm **"Bắt đầu trộn đề"** để nhận kết quả.
    """)
