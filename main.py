import os
import sys
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk
from PIL import Image, ImageOps, ImageFilter, ImageEnhance
import pytesseract
from datetime import datetime, timedelta

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class OTCalculator(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("CSV Chart Viewer - OT Calculator (Producer: KI.Shin)")
        self.geometry("1100x750")
        ctk.set_appearance_mode("light")
        
        try:
            engine_root = resource_path("Tesseract-OCR")
            pytesseract.pytesseract.tesseract_cmd = os.path.join(engine_root, "tesseract.exe")
            os.environ["TESSDATA_PREFIX"] = os.path.join(engine_root, "tessdata")
        except: pass

        self.setup_ui()

    def setup_ui(self):
        self.btn_load = ctk.CTkButton(self, text="Load Shiftee Screenshot", 
                                      command=self.load_image, font=("Segoe UI", 18, "bold"),
                                      width=350, height=55)
        self.btn_load.pack(pady=20)

        style = ttk.Style()
        style.configure("Treeview.Heading", font=("Segoe UI", 11, "bold"))
        style.configure("Treeview", font=("Segoe UI", 10), rowheight=30) 

        self.tree = ttk.Treeview(self, columns=("Date", "Range", "Net", "OT15", "Hol", "Total"), show='headings')
        headers = [("Date", "날짜"), ("Range", "근무시간"), ("Net", "실근무"), ("OT15", "연장(1.5)"), 
                   ("Hol", "휴일(1.5~)"), ("Total", "환산합계")]
        
        for col, name in headers:
            self.tree.heading(col, text=name)
            self.tree.column(col, width=150, anchor="center")
        self.tree.pack(pady=10, fill="both", expand=True, padx=20)

        self.summary_box = ctk.CTkTextbox(self, height=120, font=("Segoe UI", 22, "bold"), border_width=2)
        self.summary_box.pack(pady=20, fill="x", padx=20)

    def load_image(self):
        file_path = filedialog.askopenfilename()
        if not file_path: return
        try:
            self.btn_load.configure(text="Deep Scanning Data...", state="disabled")
            self.update()
            
            # [인식률 개선 1] 이미지 강화 전처리
            img = Image.open(file_path).convert('L')
            img = ImageEnhance.Contrast(img).enhance(2.0) # 대비 2배 강화
            img = img.point(lambda x: 0 if x < 160 else 255) # 이진화
            
            # [인식률 개선 2] PSM 3(자동)과 PSM 4(표)의 장점을 섞기 위해 PSM 4 유지
            raw_text = pytesseract.image_to_string(img, lang='kor+eng', config='--psm 4')
            
            self.process_ot_data(raw_text)
        except Exception as e:
            messagebox.showerror("Error", str(e))
        finally:
            self.btn_load.configure(text="Load Shiftee Screenshot", state="normal")

    def process_ot_data(self, raw_text):
        # [인식률 개선 3] 더 유연한 정규식: 날짜와 시간, 휴게시간을 독립적으로 찾음
        # 요일 괄호 ( )가 OCR 오류로 깨지는 경우가 많아 생략하고 숫자 위주로 매칭
        pattern = re.compile(r'(\d{1,2}/\d{1,2}).*?(\d{2}:\d{2})\s*-\s*(\d{2}:\d{2}).*?(\d+)\s*[분min]', re.S)
        matches = pattern.findall(raw_text)

        for item in self.tree.get_children(): self.tree.delete(item)
        grand_total_weighted = 0

        if not matches:
            self.summary_box.delete("0.0", "end")
            self.summary_box.insert("0.0", "⚠️ 데이터를 읽지 못했습니다. 스크린샷의 글자가 선명한지 확인해주세요.")
            return

        for m in matches:
            date_val, start_s, end_s, rest_m = m
            try:
                st = datetime.strptime(start_s, "%H:%M")
                en = datetime.strptime(end_s, "%H:%M")
                if en < st: en += timedelta(days=1)
                
                # 실근무 = 총시간 - 휴게시간
                net_h = (en - st).total_seconds() / 3600 - (float(rest_m) / 60)
                
                # 평일 연장(8시간 초과분 1.5배 가산)
                ot_15 = max(0, net_h - 8)
                
                # 환산 합계 = 기본(1.0) + 연장가산(0.5)
                weighted_h = net_h + (ot_15 * 0.5)
                
                grand_total_weighted += weighted_h

                self.tree.insert("", "end", values=(
                    date_val, f"{start_s}-{end_s}", f"{net_h:.1f}h", 
                    f"{ot_15:.1f}h", "-", f"{weighted_h:.1f}h"
                ))
            except: continue

        self.summary_box.delete("0.0", "end")
        self.summary_box.tag_config("center", justify='center')
        self.summary_box.insert("0.0", f"\nTOTAL MONTHLY OT: {grand_total_weighted:.1f} HOURS", "center")

if __name__ == "__main__":
    app = OTCalculator()
    app.mainloop()
