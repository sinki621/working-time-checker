import os
import sys
import subprocess

# [1] 터미널 없이 라이브러리를 강제 설치하는 로직
def install_and_import(package, import_name=None):
    if import_name is None:
        import_name = package
    try:
        __import__(import_name)
    except ImportError:
        print(f"[{package}] 설치 중... 잠시만 기다려주세요.")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        # 설치 후 다시 불러오기
        __import__(import_name)

# 필수 라이브러리 목록 자동 설치
install_and_import('customtkinter')
install_and_import('holidays')
install_and_import('pytesseract')
install_and_import('Pillow', 'PIL')

import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk
from PIL import Image, ImageEnhance
import pytesseract
from datetime import datetime, timedelta
import holidays
import ctypes

# [2] holidays.countries 에러 방지를 위한 강제 로드
try:
    from holidays.countries import Korea
except ImportError:
    pass

# 윈도우 4K DPI 인식 설정
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except:
    try: ctypes.windll.user32.SetProcessDPIAware()
    except: pass

class OTCalculator(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # 해상도 배율 감지
        self.scaling = self.winfo_fpixels('1i') / 96.0
        self.title("CSV Chart Viewer - OT Calculator (Producer: KI.Shin)")
        self.geometry(f"{int(1200 * self.scaling)}x{int(850 * self.scaling)}")
        ctk.set_appearance_mode("light")
        
        # 대한민국 공휴일 정보
        self.kr_holidays = holidays.KR()
        
        self.setup_ui()

    def setup_ui(self):
        btn_f = int(18 * self.scaling)
        self.btn_load = ctk.CTkButton(self, text="Load Screenshot (한/영 자동)", 
                                      command=self.load_image, font=("Segoe UI", btn_f, "bold"),
                                      width=int(400 * self.scaling), height=int(60 * self.scaling))
        self.btn_load.pack(pady=20)

        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview.Heading", font=("Segoe UI", int(12 * self.scaling), "bold"))
        style.configure("Treeview", font=("Segoe UI", int(11 * self.scaling)), rowheight=int(35 * self.scaling)) 

        self.tree = ttk.Treeview(self, columns=("Date", "Range", "Rest", "Net", "Type", "Total"), show='headings')
        headers = [("Date", "날짜(Date)"), ("Range", "시간(Time)"), ("Rest", "휴게(Break)"), 
                   ("Net", "실근무(Net)"), ("Type", "유형(Type)"), ("Total", "환산합계")]
        
        for col, name in headers:
            self.tree.heading(col, text=name)
            self.tree.column(col, width=int(160 * self.scaling), anchor="center")
        self.tree.pack(pady=10, fill="both", expand=True, padx=20)

        self.summary_box = ctk.CTkTextbox(self, height=int(130 * self.scaling), 
                                          font=("Segoe UI", int(24 * self.scaling), "bold"), border_width=2)
        self.summary_box.pack(pady=20, fill="x", padx=20)

    def load_image(self):
        f_path = filedialog.askopenfilename()
        if not f_path: return
        try:
            img = Image.open(f_path).convert('L')
            img = ImageEnhance.Contrast(img).enhance(2.2)
            img = img.point(lambda x: 0 if x < 155 else 255)
            # 한국어(분)와 영어(m) 통합 정규식
            raw_text = pytesseract.image_to_string(img, lang='kor+eng', config='--psm 4')
            self.process_ot_data(raw_text)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def process_ot_data(self, raw_text):
        # 정규식 패턴: 날짜, 시간, 휴게시간(분/m)
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
                
                if is_hol:
                    type_s = "Holiday"
                    weighted = net_h * 1.5 if net_h <= 8 else (8 * 1.5) + ((net_h - 8) * 2.0)
                else:
                    type_s = "Weekday"
                    weighted = net_h + (max(0, net_h - 8) * 0.5)
                
                total_w += weighted
                self.tree.insert("", "end", values=(f"{d_val}({day_n})", f"{s_s}-{e_s}", f"{r_m}m", f"{net_h:.1f}h", type_s, f"{weighted:.1f}h"))
            except: continue
        self.summary_box.delete("0.0", "end")
        self.summary_box.insert("0.0", f"\nTOTAL WEIGHTED OT: {total_w:.1f} HOURS")

if __name__ == "__main__":
    app = OTCalculator()
    app.mainloop()
