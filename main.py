import os
import sys
import argparse
import fnmatch
import pyperclip
import json # Thêm thư viện xử lý ipynb

# --- CÁC THƯ VIỆN XỬ LÝ TÀI LIỆU ---
try:
    import fitz  # PyMuPDF (cho PDF)
except ImportError:
    fitz = None

try:
    import docx # python-docx (cho DOCX)
except ImportError:
    docx = None

try:
    import openpyxl # openpyxl (cho XLSX)
except ImportError:
    openpyxl = None

try:
    from pptx import Presentation # python-pptx (cho PPTX)
except ImportError:
    Presentation = None


# --- HÀM TRÍCH XUẤT NỘI DUNG TỪNG LOẠI FILE ---

def read_pdf(file_path):
    if not fitz:
        return "[LỖI] Thiếu thư viện PyMuPDF. Hãy chạy: pip install PyMuPDF"
    doc = fitz.open(file_path)
    content = []
    for page in doc:
        blocks = page.get_text("blocks")
        blocks.sort(key=lambda b: (b[1], b[0]))
        for b in blocks:
            block_type = b[6]
            if block_type == 0:  # Văn bản
                content.append(b[4].strip())
            elif block_type == 1:  # Hình ảnh
                content.append("[IMAGE PLACEHOLDER]")
    return "\n\n".join(content)

def read_docx(file_path):
    if not docx:
        return "[LỖI] Thiếu thư viện python-docx. Hãy chạy: pip install python-docx"
    doc = docx.Document(file_path)
    content = []
    
    for block in doc.element.body:
        if block.tag.endswith('p'): 
            p = docx.text.paragraph.Paragraph(block, doc)
            if 'w:drawing' in block.xml or 'pic:pic' in block.xml:
                content.append("[IMAGE PLACEHOLDER]")
            if p.text.strip():
                content.append(p.text.strip())
                
        elif block.tag.endswith('tbl'):
            t = docx.table.Table(block, doc)
            for row in t.rows:
                row_data = [cell.text.replace('\n', ' ').strip() for cell in row.cells]
                content.append("| " + " | ".join(row_data) + " |")
            content.append("")
            
    return "\n".join(content)

def read_excel(file_path):
    if not openpyxl:
        return "[LỖI] Thiếu thư viện openpyxl. Hãy chạy: pip install openpyxl"
    wb = openpyxl.load_workbook(file_path, data_only=True)
    content = []
    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        content.append(f"### Sheet: {sheet_name}")
        for row in sheet.iter_rows(values_only=True):
            if not any(cell is not None and str(cell).strip() != "" for cell in row):
                continue
            row_data = [str(cell) if cell is not None else "" for cell in row]
            while row_data and row_data[-1].strip() == "":
                row_data.pop()
            content.append("| " + " | ".join(row_data) + " |")
        content.append("\n")
    return "\n".join(content)

def read_pptx(file_path):
    if not Presentation:
        return "[LỖI] Thiếu thư viện python-pptx. Hãy chạy: pip install python-pptx"
    prs = Presentation(file_path)
    content = []
    for i, slide in enumerate(prs.slides):
        content.append(f"### Slide {i+1}")
        shapes_data = []
        for shape in slide.shapes:
            top = shape.top if hasattr(shape, "top") and shape.top is not None else 0
            left = shape.left if hasattr(shape, "left") and shape.left is not None else 0
            
            if hasattr(shape, "text_frame") and shape.text_frame:
                shapes_data.append((top, left, "text", shape.text))
            elif hasattr(shape, "image") or shape.shape_type == 13: 
                shapes_data.append((top, left, "image", "[IMAGE PLACEHOLDER]"))
                
        shapes_data.sort(key=lambda x: (x[0], x[1]))
        for s in shapes_data:
            if s[3].strip():
                content.append(s[3].strip())
        content.append("\n")
    return "\n".join(content)

def read_ipynb(file_path):
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            nb = json.load(f)
    except Exception as e:
        return f"[LỖI] Không thể đọc file ipynb: {e}"

    content = []
    for cell in nb.get('cells', []):
        cell_type = cell.get('cell_type', '')
        source = "".join(cell.get('source', []))

        if cell_type == 'markdown':
            content.append(source.strip())
            if 'attachments' in cell:
                content.append("[IMAGE PLACEHOLDER]")
                
        elif cell_type == 'code':
            if source.strip():
                content.append(f"```python\n{source}\n```")
            for output in cell.get('outputs', []):
                output_type = output.get('output_type', '')
                if output_type in ('display_data', 'execute_result'):
                    data = output.get('data', {})
                    if any(k.startswith('image/') for k in data.keys()):
                        content.append("[IMAGE PLACEHOLDER]")
                    elif 'text/plain' in data:
                        content.append("".join(data['text/plain']).strip())
                elif output_type == 'stream':
                    content.append("".join(output.get('text', [])).strip())
                    
    return "\n\n".join(content)

# --- HELPER: GIỚI HẠN SỐ DÒNG THEO FILTER ---
def apply_line_limit(content, rel_path, limit_args):
    if not limit_args:
        return content
        
    unix_path = rel_path.replace(os.sep, '/')
    basename = os.path.basename(unix_path)
    
    for pattern, max_lines_str in limit_args:
        if not max_lines_str.isdigit():
            continue
        max_lines = int(max_lines_str)
        
        # Logic match giống gitignore
        is_match = False
        if pattern.startswith('/'):
            if fnmatch.fnmatch(unix_path, pattern.lstrip('/')): is_match = True
        elif '/' not in pattern:
            if fnmatch.fnmatch(basename, pattern): is_match = True
        else:
            if fnmatch.fnmatch(unix_path, pattern) or fnmatch.fnmatch(unix_path, f"*/{pattern}"): is_match = True
            
        if is_match:
            lines = content.split('\n')
            if len(lines) > max_lines:
                return '\n'.join(lines[:max_lines]) + f"\n\n[... ĐÃ CẮT BỚT. CHỈ HIỂN THỊ {max_lines} DÒNG ĐẦU ...]"
            break # Áp dụng limit đầu tiên match được
            
    return content


# --- HÀM CHÍNH ---

def combine_files_to_clipboard(dir_path, ignore_patterns, args):
    if not os.path.isdir(dir_path):
        print(f"Lỗi: '{dir_path}' không tồn tại.")
        sys.exit(1)

    print(f"Đang quét: {dir_path}")
    if ignore_patterns:
        print(f"Danh sách loại trừ (Gitignore style): {ignore_patterns}")
    if args.limit:
        print(f"Luật giới hạn số dòng: {args.limit}")

    def is_ignored(rel_path, is_dir):
        unix_path = rel_path.replace(os.sep, '/')
        basename = os.path.basename(unix_path)
        for pattern in ignore_patterns:
            dir_only = pattern.endswith('/')
            p = pattern.rstrip('/') 
            if dir_only and not is_dir:
                continue 
            if p.startswith('/'):
                if fnmatch.fnmatch(unix_path, p.lstrip('/')): return True
            elif '/' not in p:
                if fnmatch.fnmatch(basename, p): return True
            else:
                if fnmatch.fnmatch(unix_path, p) or fnmatch.fnmatch(unix_path, f"*/{p}"): return True
        return False

    all_content = []

    try:
        for root, dirs, files in os.walk(dir_path, topdown=True):
            rel_root = os.path.relpath(root, dir_path)
            if rel_root == ".": rel_root = ""

            valid_dirs = []
            for d in dirs:
                d_rel_path = os.path.join(rel_root, d)
                if not is_ignored(d_rel_path, is_dir=True):
                    valid_dirs.append(d)
            dirs[:] = valid_dirs 

            for filename in files:
                f_rel_path = os.path.join(rel_root, filename)
                if is_ignored(f_rel_path, is_dir=False):
                    continue

                full_path = os.path.join(root, filename)
                display_path = f_rel_path.replace(os.sep, '/')
                ext = os.path.splitext(filename)[1].lower()

                all_content.append(f"=== FILE: {display_path} ===\n")
                
                raw_content = ""
                try:
                    if ext == '.docx' and args.docx:
                        raw_content = read_docx(full_path)
                    elif ext == '.pdf' and args.pdf:
                        raw_content = read_pdf(full_path)
                    elif ext in ['.xlsx', '.xls'] and args.excel:
                        raw_content = read_excel(full_path)
                    elif ext == '.pptx' and args.pptx:
                        raw_content = read_pptx(full_path)
                    elif ext == '.ipynb' and args.ipynb:
                        raw_content = read_ipynb(full_path)
                    else:
                        with open(full_path, 'r', encoding='utf-8') as infile:
                            raw_content = infile.read()
                            
                    # Cắt dòng nếu trúng luật -l
                    final_content = apply_line_limit(raw_content, f_rel_path, args.limit)
                    all_content.append(final_content)
                    
                except UnicodeDecodeError:
                    all_content.append("[WARNING] File nhị phân bị bỏ qua do không hỗ trợ hoặc không truyền cờ phù hợp.")
                except Exception as e:
                    all_content.append(f"[LỖI XỬ LÝ] {str(e)}")
                
                all_content.append("\n\n")

        final_output = "".join(all_content)
        
        if not final_output:
            print("Không có nội dung nào được copy.")
            return

        pyperclip.copy(final_output)
        print(f"\nThành công! Đã copy {len(final_output)} ký tự vào clipboard.")

    except Exception as e:
        print(f"Lỗi chương trình: {e}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Copy nội dung file vào clipboard, hỗ trợ các định dạng nâng cao và cú pháp loại trừ.")
    parser.add_argument("folder", help="Thư mục cần quét")
    parser.add_argument("-e", "--exclude", nargs='*', default=[], help="Danh sách loại trừ (VD: __pycache__/ *.pyc node_modules/ .env)")
    
    # Action 'append' + nargs=2 để nhập nhiều điều kiện (VD: -l *.csv 100 -l src/*.js 500)
    parser.add_argument("-l", "--limit", action="append", nargs=2, metavar=('PATTERN', 'LINES'), help="Giới hạn dòng (VD: -l *.csv 100)")
    
    parser.add_argument("-d", "--docx", action="store_true", help="Bật quét nội dung file .docx")
    parser.add_argument("-p", "--pdf", action="store_true", help="Bật quét nội dung file .pdf")
    parser.add_argument("-x", "--excel", action="store_true", help="Bật quét nội dung file .xlsx")
    parser.add_argument("-t", "--pptx", action="store_true", help="Bật quét nội dung file .pptx")
    parser.add_argument("-i", "--ipynb", action="store_true", help="Bật quét nội dung file .ipynb")
    
    args = parser.parse_args()
    combine_files_to_clipboard(args.folder, args.exclude, args)