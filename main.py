import os
import sys
import subprocess

# [추가] 모듈이 없을 경우 자동 설치하는 함수
def install_requirements():
    required = ['customtkinter', 'holidays', 'pytesseract', 'Pillow']
    for lib in required:
        try:
            __import__(lib if lib != 'Pillow' else 'PIL')
        except ImportError:
            print(f"Installing {lib}...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", lib])

# 실행 시 최우선적으로 설치 확인
install_requirements()

import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk
from PIL import Image, ImageEnhance
import pytesseract
from datetime import datetime, timedelta
import holidays
import ctypes

# 윈도우 고해상도 DPI 인식 (웹 기반 데스크톱 환경 대응)
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except: pass

class OTCalculator(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.scaling = self.winfo_fpixels('1i') / 96.0
        self.title("CSV Chart Viewer - OT Calculator (Producer: KI.Shin)")
        self.geometry(f"{int(1150 * self.scaling)}x{int(800 * self.scaling)}")
        ctk.set_appearance_mode("light")
        self.kr_holidays = holidays.KR()
        self.setup_ui()

    def setup_ui(self):
        self.btn_load = ctk.CTkButton(self, text="Load Screenshot (KOR/ENG)", 
                                      command=self.load_image, font=("Segoe UI", int(18 * self.scaling), "bold"),
                                      width=int(400 * self.scaling), height=int(55 * self.scaling))
        self.btn_load.pack(pady=20)

        style = ttk.Style()
        style.configure("Treeview.Heading", font=("Segoe UI", int(12 * self.scaling), "bold"))
        style.configure("Treeview", font=("Segoe UI", int(11 * self.scaling)), rowheight=int(35 * self.scaling)) 

        self.tree = ttk.Treeview(self, columns=("Date", "Range", "Rest", "Net", "Type", "Total"), show='headings')
        headers = [("Date", "날짜(Date)"), ("Range", "시간(Time)"), ("Rest", "휴게(Break)"), 
                   ("Net", "실근무(Net)"), ("Type", "유형(Type)"), ("Total", "환산합계")]
        
        for col, name in headers:
            self.tree.heading(col, text=name)
            self.tree.column(col, width=int(150 * self.scaling), anchor="center")
        self.tree.pack(pady=10, fill="both", expand=True, padx=20)

        self.summary_box = ctk.CTkTextbox(self, height=int(120 * self.scaling), 
                                          font=("Segoe UI", int(24 * self.scaling), "bold"), border_width=2)
        self.summary_box.pack(pady=20, fill="x", padx=20)

    def load_image(self):
        file_path = filedialog.askopenfilename()
        if not file_path: return
        try:
            img = Image.open(file_path).convert('L')
            img = ImageEnhance.Contrast(img).enhance(2.2)
            img = img.point(lambda x: 0 if x < 155 else 255)
            # 한글(분)과 영어(m) 유연하게 인식
            raw_text = pytesseract.image_to_string(img, lang='kor+eng', config='--psm 4')
            self.process_ot_data(raw_text)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def process_ot_data(self, raw_text):
        pattern = re.compile(r'(\d{1,2}/\d{1,2}).*?(\d{2}:\d{2})\s*-\s*(\d{2}:\d{2}).*?(\d+)\s*(?:분|m|min)', re.S | re.I)
        matches = pattern.findall(raw_text)
        for item in self.tree.get_children(): self.tree.delete(item)
        total_w = 0
        curr_year = datetime.now().year

        for m in matches:
            d_val, s_s, e_s, r_m = m
            try:
                dt_obj = datetime.strptime(f"{curr_year}/{d_val}", "%Y/%m/%d")
                is_hol = dt_obj.weekday() >= 5 or dt_obj in self.kr_holidays
                day_n = ["월", "화", "수", "목", "금", "토", "일"][dt_obj.weekday()]
                st, en = datetime.strptime(s_s, "%H:%M"), datetime.strptime(e_s, "%H:%M")
                if en < st: en += timedelta(days=1)
                net_h = (en - st).total_seconds() / 3600 - (float(r_m) / 60)
                
                # 수당 계산 로직 (주말/공휴일 가산 포함)
                if is_hol:
                    type_s = f"Holiday({day_n})"
                    weighted = net_h * 1.5 if net_h <= 8 else (8 * 1.5) + ((net_h - 8) * 2.0)
                else:
                    type_s = "Weekday"
                    weighted = net_h + (max(0, net_h - 8) * 0.5)
                
                total_w += weighted
                self.tree.insert("", "end", values=(f"{d_val}({day_n})", f"{s_s}-{e_s}", f"{r_m}m", f"{net_h:.1f}h", type_s, f"{weighted:.1f}h"))
            except: continue
        self.summary_box.delete("0.0", "end")
        self.summary_box.insert("0.0", f"\nTOTAL WEIGHTED: {total_w:.1f} HOURS")

if __name__ == "__main__":
    app = OTCalculator()
    app.mainloop()
