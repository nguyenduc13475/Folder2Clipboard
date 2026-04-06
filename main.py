import os
import sys
import argparse
import fnmatch
import pyperclip

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
        # Lấy các block nội dung (văn bản và hình ảnh)
        blocks = page.get_text("blocks")
        # Sắp xếp từ trên xuống dưới (y0), từ trái qua phải (x0)
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
    
    # Duyệt qua các thành phần XML cốt lõi để giữ đúng thứ tự bảng/văn bản
    for block in doc.element.body:
        if block.tag.endswith('p'): # Paragraph
            p = docx.text.paragraph.Paragraph(block, doc)
            # Kiểm tra xem có chứa hình ảnh không (thông qua thẻ w:drawing hoặc pic:pic)
            if 'w:drawing' in block.xml or 'pic:pic' in block.xml:
                content.append("[IMAGE PLACEHOLDER]")
            if p.text.strip():
                content.append(p.text.strip())
                
        elif block.tag.endswith('tbl'): # Table
            t = docx.table.Table(block, doc)
            for row in t.rows:
                row_data = [cell.text.replace('\n', ' ').strip() for cell in row.cells]
                content.append("| " + " | ".join(row_data) + " |")
            content.append("") # Dòng trống sau bảng
            
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
            # Bỏ qua các dòng hoàn toàn rỗng (tất cả các ô đều là None hoặc chuỗi trống)
            if not any(cell is not None and str(cell).strip() != "" for cell in row):
                continue
                
            row_data = [str(cell) if cell is not None else "" for cell in row]
            
            # Xóa các ô rỗng thừa ở cuối mỗi dòng để giảm dung lượng chuỗi
            while row_data and row_data[-1].strip() == "":
                row_data.pop()
                
            # Tạo bảng markdown cơ bản
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
            # Lấy vị trí để sắp xếp (ưu tiên các shape không có vị trí rõ ràng lên trước hoặc bỏ qua)
            top = shape.top if hasattr(shape, "top") and shape.top is not None else 0
            left = shape.left if hasattr(shape, "left") and shape.left is not None else 0
            
            if hasattr(shape, "text_frame") and shape.text_frame:
                shapes_data.append((top, left, "text", shape.text))
            elif hasattr(shape, "image") or shape.shape_type == 13: # 13 là msoPICTURE
                shapes_data.append((top, left, "image", "[IMAGE PLACEHOLDER]"))
                
        # Sắp xếp: từ trên xuống dưới, từ trái qua phải
        shapes_data.sort(key=lambda x: (x[0], x[1]))
        
        for s in shapes_data:
            if s[3].strip():
                content.append(s[3].strip())
        content.append("\n")
    return "\n".join(content)


# --- HÀM CHÍNH ---

def combine_files_to_clipboard(dir_path, ignore_patterns, args):
    if not os.path.isdir(dir_path):
        print(f"Lỗi: '{dir_path}' không tồn tại.")
        sys.exit(1)

    print(f"Đang quét: {dir_path}")
    if ignore_patterns:
        print(f"Danh sách loại trừ (Gitignore style): {ignore_patterns}")

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

            # Lọc thư mục
            valid_dirs = []
            for d in dirs:
                d_rel_path = os.path.join(rel_root, d)
                if not is_ignored(d_rel_path, is_dir=True):
                    valid_dirs.append(d)
            dirs[:] = valid_dirs 

            # Lọc file
            for filename in files:
                f_rel_path = os.path.join(rel_root, filename)
                if is_ignored(f_rel_path, is_dir=False):
                    continue

                full_path = os.path.join(root, filename)
                display_path = f_rel_path.replace(os.sep, '/')
                ext = os.path.splitext(filename)[1].lower()

                all_content.append(f"=== FILE: {display_path} ===\n")
                
                try:
                    # Chặn logic đọc file theo extension
                    if ext == '.docx' and args.docx:
                        all_content.append(read_docx(full_path))
                    elif ext == '.pdf' and args.pdf:
                        all_content.append(read_pdf(full_path))
                    elif ext in ['.xlsx', '.xls'] and args.excel:
                        all_content.append(read_excel(full_path))
                    elif ext == '.pptx' and args.pptx:
                        all_content.append(read_pptx(full_path))
                    else:
                        # Fallback về đọc file text tiêu chuẩn (code, txt, md...)
                        with open(full_path, 'r', encoding='utf-8') as infile:
                            content = infile.read()
                            all_content.append(content)
                except UnicodeDecodeError:
                    # Bỏ qua nếu file là nhị phân mà không có cờ tương ứng
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
    
    # Sử dụng action="store_true" cho phép gom cờ lại với nhau (VD: -dpxt)
    parser.add_argument("-d", "--docx", action="store_true", help="Bật quét nội dung file .docx")
    parser.add_argument("-p", "--pdf", action="store_true", help="Bật quét nội dung file .pdf")
    parser.add_argument("-x", "--excel", action="store_true", help="Bật quét nội dung file .xlsx")
    parser.add_argument("-t", "--pptx", action="store_true", help="Bật quét nội dung file .pptx")
    
    args = parser.parse_args()
    combine_files_to_clipboard(args.folder, args.exclude, args)