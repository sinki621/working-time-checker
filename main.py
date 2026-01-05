import os
import sys
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk
from PIL import Image
import pytesseract
from datetime import datetime

# EXE 빌드 시 내부 리소스 경로를 찾는 핵심 함수
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class OTCalculator(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # 기본 설정
        self.title("CSV Chart Viewer - OT Calculator (Producer: KI.Shin)")
        self.geometry("1100x850")
        ctk.set_appearance_mode("light")
        
        # Tesseract 설정 (EXE 내부의 tessdata 폴더 참조)
        self.tessdata_path = resource_path("tessdata")
        os.environ["TESSDATA_PREFIX"] = self.tessdata_path

        # UI 구성
        self.setup_ui()

    def setup_ui(self):
        # 상단 버튼
        self.btn_load = ctk.CTkButton(self, text="Load shiftee screenshot (PNG, JPG)", 
                                      command=self.load_image, font=("Segoe UI", 14, "bold"),
                                      width=350, height=50)
        self.btn_load.pack(pady=20)

        self.lbl_info = ctk.CTkLabel(self, text="normal 8 hour work automatically excluded", 
                                     font=("Segoe UI", 13, "bold"))
        self.lbl_info.pack(pady=5)

        # 결과 리스트뷰 (표)
        self.style = ttk.Style()
        self.style.configure("Treeview", font=("Segoe UI", 10), rowheight=25)
        
        self.tree = ttk.Treeview(self, columns=("Date", "Range", "Work", "OT1.5", "Night", "Hol", "Total"), show='headings')
        cols = [("Date", 120), ("Range", 120), ("Work", 80), ("OT1.5", 80), ("Night", 80), ("Hol", 80), ("Total", 80)]
        for col, width in cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=width, anchor="center")
        self.tree.pack(pady=10, fill="both", expand=True, padx=20)

        # 하단 요약창
        self.summary_box = ctk.CTkTextbox(self, height=150, font=("Consolas", 12))
        self.summary_box.pack(pady=20, fill="x", padx=20)

    def load_image(self):
        file_path = filedialog.askopenfilename(filetypes=[("Image Files", "*.png;*.jpg;*.jpeg")])
        if not file_path:
            return

        try:
            self.btn_load.configure(text="Analyzing...", state="disabled")
            self.update()

            # OCR 실행
            text = pytesseract.image_to_string(Image.open(file_path), lang='kor+eng')
            self.process_data(text)
            
        except Exception as e:
            messagebox.showerror("OCR Error", f"인증 엔진 초기화 실패 또는 에러:\n{str(e)}")
        finally:
            self.btn_load.configure(text="Load shiftee screenshot (PNG, JPG)", state="normal")

    def process_data(self, raw_text):
        # 기존 C#과 동일한 정규식 패턴
        pattern = re.compile(r'(\d{1,2}/\d{1,2})\s*\(\s*(.)\s*\)\s*(\d{2}:\d{2})\s*[-~]\s*(\d{2}:\d{2}).*?(\d+)\s*(?:분|min)', re.S | re.I)
        matches = pattern.findall(raw_text)

        for item in self.tree.get_children():
            self.tree.delete(item)

        total_ot15, total_night, total_hol, grand_total = 0, 0, 0, 0

        for m in matches:
            date, day, start_str, end_str, rest_m = m
            fmt = "%H:%M"
            start_t = datetime.strptime(start_str, fmt)
            end_t = datetime.strptime(end_str, fmt)
            if end_t < start_t: end_t += timedelta(days=1)
            
            rest_h = float(rest_m) / 60.0
            net_h = (end_t - start_t).total_seconds() / 3600.0 - rest_h
            is_holiday = day in ["토", "일"] or day.upper().startswith("S")

            ot15 = max(0, net_h - 8)
            night05 = 0 # (간단화를 위해 0처리, 필요시 추가 구현 가능)
            hol05 = net_h if is_holiday else 0
            
            day_sum = (ot15 * 1.5) + (night05 * 0.5) + (hol05 * 0.5)
            if is_holiday and ot15 > 0: day_sum += (ot15 * 0.5)

            total_ot15 += ot15
            total_night += night05
            total_hol += hol05
            grand_total += day_sum

            self.tree.insert("", "end", values=(f"{date}({day})", f"{start_str}-{end_str}", f"{net_h:.1f}", 
                                               f"{ot15:.1f}", "-", f"{hol05:.1f}", f"{day_sum:.1f}"))

        summary = f"""[ Monthly Overtime Summary ]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
▶ Total Excess Hours (>8h): {total_ot15:.1f} h
▶ Total Holiday Hours: {total_hol:.1f} h
▶ [GRAND TOTAL] Weighted OT: {grand_total:.1f} h
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"""
        self.summary_box.delete("0.0", "end")
        self.summary_box.insert("0.0", summary)

if __name__ == "__main__":
    app = OTCalculator()
    app.mainloop()
