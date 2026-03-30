import os
import fitz  # PyMuPDF
import re

def mask_name(name):
    """
    Mask name: 홍길동 -> 홍*동
    """
    if len(name) <= 1:
        return name
    if len(name) == 2:
        return name[0] + "*"
    return name[0] + "*" * (len(name) - 2) + name[-1]

def mask_cert_num(cert_num):
    """
    Mask certificate number: 1 digit masked for every 2 digits.
    Example: 12345678 -> 1*3*5*7*
    """
    chars = list(cert_num)
    for i in range(1, len(chars), 2):
        chars[i] = "*"
    return "".join(chars)

def mask_ssn(ssn):
    """
    Mask SSN: 990915-1555555 -> 990915-1******
    """
    if "-" in ssn:
        front, back = ssn.split("-")
        if len(back) > 0:
            return f"{front}-{back[0]}{'*' * (len(back)-1)}"
    return ssn

class PDFProcessor:
    def __init__(self, input_path, output_dir, log_callback=None):
        self.input_path = input_path
        self.output_dir = output_dir
        self.log_callback = log_callback

    def log(self, message):
        if self.log_callback:
            self.log_callback(message)
        else:
            print(message)

    def process(self):
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)

        doc = fitz.open(self.input_path)
        total_pages = len(doc)
        self.log(f"전체 {total_pages} 페이지를 처리합니다.")

        # Find a Korean font robustly
        ko_font_path = None
        windir = os.environ.get('WINDIR', 'C:/Windows')
        possible_fonts = [
            os.path.join(windir, 'Fonts', 'malgun.ttf'),     # Malgun Gothic
            os.path.join(windir, 'Fonts', 'gulim.ttc'),      # Gulim
            os.path.join(windir, 'Fonts', 'batang.ttc'),    # Batang
            "/usr/share/fonts/truetype/nanum/NanumGothic.ttf"# Linux Nanum
        ]
        
        for fpath in possible_fonts:
            if os.path.exists(fpath):
                ko_font_path = fpath
                self.log(f"상태: 한글 폰트를 찾았습니다 ({os.path.basename(fpath)})")
                break
        
        if not ko_font_path:
            self.log("경고: 시스템에서 한글 폰트를 찾을 수 없습니다. 마스킹된 이름이 제대로 표시되지 않을 수 있습니다.")

        for page_num in range(total_pages):
            # Create a separate document for this page to avoid side effects
            temp_doc = fitz.open()
            temp_doc.insert_pdf(doc, from_page=page_num, to_page=page_num)
            page = temp_doc[0]
            
            words = page.get_text("words")
            
            serial_no = ""
            name = ""
            cert_num = ""
            ssn = ""
            
            ssn_pattern = re.compile(r"\d{6}-\d{7}")
            cert_pattern = re.compile(r"\d{10,12}")

            # Find headers X coordinates
            header_y = 0
            name_x = 0
            serial_x = 0
            for w in words:
                word_text = w[4]
                if "연번" in word_text:
                    serial_x = (w[0] + w[2]) / 2
                    header_y = w[3]
                elif "성명" in word_text:
                    name_x = (w[0] + w[2]) / 2
            
            # Find data row and collect Name
            data_row_y = 0
            name_parts = []
            
            # First pass: find serial no and data row Y
            for w in words:
                if w[1] > header_y and not data_row_y:
                    # Tighten threshold from 30 to 15 to avoid name column
                    if abs((w[0] + w[2]) / 2 - serial_x) < 15:
                        serial_no = w[4]
                        data_row_y = (w[1] + w[3]) / 2
            
            # Second pass: collect all words in the Name column at that Y
            if data_row_y:
                for w in words:
                    # Narrow Y range and X range to avoid adjacent columns
                    if abs((w[1] + w[3]) / 2 - data_row_y) < 10:
                        # Tighten threshold from 40 to 20 to avoid serial column
                        if abs((w[0] + w[2]) / 2 - name_x) < 20: 
                            if w[4] != serial_no: # Safety check
                                name_parts.append(w[4])
            
            name = "".join(name_parts).strip()
            
            # Find SSN and Cert Num
            for w in words:
                text = w[4]
                if ssn_pattern.match(text):
                    ssn = text
                elif cert_pattern.match(text):
                    # Sometimes cert number is split, but let's assume it's one word for now
                    cert_num = text

            if not serial_no or not name:
                self.log(f"경고: {page_num + 1} 페이지에서 정보를 찾을 수 없습니다. (추출된 이름: '{name}')")
                temp_doc.close()
                continue

            # Masking
            masked_name_text = mask_name(name)
            masked_cert_text = mask_cert_num(cert_num) if cert_num else ""
            masked_ssn_text = mask_ssn(ssn) if ssn else ""
            
            self.log(f"[{page_num + 1}/{total_pages}] 처리 중: {serial_no}-{name} -> {masked_name_text}")

            # Redaction and Insertion
            # 1. Name
            # We search for the full name. If it's split in PDF, search_for might fail.
            # In that case, we search for each part and redact.
            # But the replacement should be the full masked name.
            name_rects = page.search_for(name)
            if not name_rects and name_parts:
                # If full name search fails, redact all parts
                for part in name_parts:
                    name_rects.extend(page.search_for(part))
            
            # We only want to insert the masked name ONCE at the first rect position
            if name_rects:
                target_rect = name_rects[0]
                for rect in name_rects:
                    page.add_redact_annot(rect, fill=(1, 1, 1))
            
            # 2. Cert Num
            cert_rects = []
            if cert_num:
                cert_rects = page.search_for(cert_num)
                for rect in cert_rects:
                    page.add_redact_annot(rect, fill=(1, 1, 1))
            
            # 3. SSN
            ssn_rects = []
            if ssn:
                ssn_rects = page.search_for(ssn)
                for rect in ssn_rects:
                    page.add_redact_annot(rect, fill=(1, 1, 1))

            page.apply_redactions()

            # Insert Masked Text with Korean font support
            font_size = 9 # Slightly larger font for better visibility
            
            if name_rects:
                rect = name_rects[0]
                # Bounding box bottom-left for better CJK alignment
                pos = (rect.x0, rect.y1 - 3) 
                
                try:
                    if ko_font_path:
                        # Use insert_font for proper CID font embedding
                        page.insert_font(fontname="ko-font", fontfile=ko_font_path)
                        page.insert_text(pos, masked_name_text, 
                                         fontsize=font_size, fontname="ko-font")
                    else:
                        page.insert_text(pos, masked_name_text, fontsize=font_size)
                except Exception as e:
                    self.log(f"텍스트 삽입 오류 (성명): {str(e)}")
            
            if cert_rects:
                for rect in cert_rects:
                    if ko_font_path:
                        page.insert_text((rect.x0, rect.y1 - 2), masked_cert_text, 
                                         fontsize=font_size, fontfile=ko_font_path)
                    else:
                        page.insert_text(rect.tl, masked_cert_text, fontsize=font_size)
            
            if ssn_rects:
                for rect in ssn_rects:
                    if ko_font_path:
                        page.insert_text((rect.x0, rect.y1 - 2), masked_ssn_text, 
                                         fontsize=font_size, fontfile=ko_font_path)
                    else:
                        page.insert_text(rect.tl, masked_ssn_text, fontsize=font_size)

            # Save individual page
            # Clean filename
            safe_name = "".join(x for x in name if x.isalnum())
            output_filename = f"{serial_no}-{safe_name}.pdf"
            output_path = os.path.join(self.output_dir, output_filename)
            
            temp_doc.save(output_path)
            temp_doc.close()
            self.log(f"저장 완료: {output_filename}")

        doc.close()
        self.log("\n모든 작업이 완료되었습니다!")

if __name__ == "__main__":
    # Test logic
    pass
