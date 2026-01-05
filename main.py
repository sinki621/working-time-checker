import os
import sys
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk
from PIL import Image
import pytesseract
from datetime import datetime, timedelta

# EXE 빌드 시 내부 리소스(엔진, 데이터) 경로를 찾는 함수
def resource_path(relative_path):
    try:
        # PyInstaller에 의해 생성된 임시 폴더 경로
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class OTCalculator(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # 프로그램 정보 설정
        self.title("CSV Chart Viewer - OT Calculator (Producer: KI.Shin)")
        self.geometry("1100(cite: 850")
        ctk.set_appearance_mode("light")
        
        try:
            # 1. Tesseract 엔진 실행 파일 경로 설정 (EXE 내부 Tesseract-OCR 폴더)
            engine_root = resource_path("Tesseract-OCR")
            tesseract_exe = os.path.join(engine_root, "tesseract.exe")
            pytesseract.pytesseract.tesseract_cmd = tesseract_exe
            
            # 2. 언어 데이터(tessdata) 경로 설정
            os.environ["TESSDATA_PREFIX"] = os.path.join(engine_root, "tessdata")
        except Exception as e:
            print(f"Engine Init Path Error: {e}")

        self.setup_ui()

    def setup_ui(self):
        # 상단 로드 버튼
        self.btn_load = ctk.CTkButton(self, text="Load shiftee screenshot (PNG, JPG)", 
                                      command=self.load_image, font=("Segoe UI", 14, "bold"),
                                      width=350, height=50, fg_color="#007ACC")
        self.btn_load.pack(pady=20)

        self.lbl_info = ctk.CTkLabel(self, text="normal 8 hour work automatically excluded", 
                                     font=("Segoe UI", 13, "bold"), text_color="#404040")
        self.lbl_info.pack(pady=5)

        # 결과 리스트뷰
        self.style = ttk.Style()
        self.style.configure("Treeview", font=("Segoe UI", 10), rowheight=25)
        
        self.tree = ttk.Treeview(self, columns=("Date", "Range", "Work", "OT1.5", "Night", "Hol", "Total"), show='headings')
        cols = [("Date", 120), ("Range", 120), ("Work", 80), ("OT1.5", 80), ("Night", 80), ("Hol", 80), ("Total", 80)]
        for col, width in cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=width, anchor="center")
        self.tree.pack(pady=10, fill="both", expand=True, padx=20)

        # 하단 요약 결과창
        self.summary_box = ctk.CTkTextbox(self, height=180, font=("Consolas", 12), border_width=1)
        self.summary_box.pack(pady=20, fill="x", padx=20)

    def load_image(self):
        file_path = filedialog.askopenfilename(filetypes=[("Image Files", "*.png;*.jpg;*.jpeg")])
        if not file_path:
            return

        try:
            self.btn_load.configure(text="Processing OCR Engine...", state="disabled")
            self.update()

            # 이미지 열기 및 OCR 수행
            img = Image.open(file_path)
            raw_text = pytesseract.image_to_string(img, lang='kor+eng')
            self.process_data(raw_text)
            
        except Exception as e:
            messagebox.showerror("OCR Error", f"Engine failed to start:\n{str(e)}\n\nCheck if Tesseract-OCR folder exists inside EXE.")
        finally:
            self.btn_load.configure(text="Load shiftee screenshot (PNG, JPG)", state="normal")

    def process_data(self, raw_text):
        # 정규식 패턴: 날짜, 요일, 시작시간, 종료시간, 휴게시간(분) 추출
        pattern = re.compile(r'(\d{1,2}/\d{1,2})\s*\(\s*(.)\s*\)\s*(\d{2}:\d{2})\s*[-~]\s*(\d{2}:\d{2}).*?(\d+)\s*(?:분|min)', re.S | re.I)
        matches = pattern.findall(raw_text)

        for item in self.tree.get_children():
            self.tree.delete(item)

        t_ot15, t_night, t_hol, g_total = 0, 0, 0, 0

        for m in matches:
            date, day, start_s, end_s, rest_m = m
            fmt = "%H:%M"
            st_t = datetime.strptime(start_s, fmt)
            en_t = datetime.strptime(end_s, fmt)
            if en_t < st_t: en_t += timedelta(days=1)
            
            rest_h = float(rest_m) / 60.0
            net_h = (en_t - st_t).total_seconds() / 3600.0 - rest_h
            is_hol = day in ["토", "일"] or day.upper().startswith("S")

            # 수당 계산 로직
            ot15 = max(0, net_h - 8)
            night05 = 0 # 심야 시간 계산 로직 추가 가능
            hol05 = net_h if is_hol else 0
            
            day_sum = (ot15 * 1.5) + (night05 * 0.5) + (hol05 * 0.5)
            if is_hol and ot15 > 0: day_sum += (ot15 * 0.5)

            t_ot15 += ot15; t_hol += hol05; g_total += day_sum

            self.tree.insert("", "end", values=(f"{date}({day})", f"{start_s}-{end_s}", f"{net_h:.1f}", 
                                               f"{ot15:.1f}", "-", f"{hol05:.1f}", f"{day_sum:.1f}"))

        summary_text = f""" [ Monthly Overtime Summary (Producer: KI.Shin) ]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  ▶ Total Excess Hours (> 8h)    : {t_ot15:.1f} hours
  ▶ Total Holiday Work Hours     : {t_hol:.1f} hours
  ▶ [GRAND TOTAL] Weighted OT    : {g_total:.1f} hours
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"""
        self.summary_box.delete("0.0", "end")
        self.summary_box.insert("0.0", summary_text)

if __name__ == "__main__":
    app = OTCalculator()
    app.mainloop()
