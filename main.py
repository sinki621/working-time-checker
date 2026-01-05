import os
import sys
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk
from PIL import Image, ImageEnhance
import pytesseract
from datetime import datetime, timedelta
import holidays
import ctypes

# 1. 4K DPI 인식 설정
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except:
        pass

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class OTCalculator(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # 해상도 배율 계산
        self.scaling = self.winfo_fpixels('1i') / 96.0
        self.title("CSV Chart Viewer - OT Calculator (Producer: KI.Shin)")
        
        win_w = int(1150 * self.scaling)
        win_h = int(800 * self.scaling)
        self.geometry(f"{win_w}x{win_h}")
        ctk.set_appearance_mode("light")
        
        self.kr_holidays = holidays.KR()
        
        try:
            engine_root = resource_path("Tesseract-OCR")
            pytesseract.pytesseract.tesseract_cmd = os.path.join(engine_root, "tesseract.exe")
            os.environ["TESSDATA_PREFIX"] = os.path.join(engine_root, "tessdata")
        except: pass

        self.setup_ui()

    def setup_ui(self):
        btn_font = int(18 * self.scaling)
        self.btn_load = ctk.CTkButton(self, text="Load Screenshot (KOR/ENG)", 
                                      command=self.load_image, font=("Segoe UI", btn_font, "bold"),
                                      width=int(400 * self.scaling), height=int(55 * self.scaling))
        self.btn_load.pack(pady=20)

        # 표 글씨 크기 설정
        h_font = int(12 * self.scaling)
        b_font = int(11 * self.scaling)
        r_height = int(35 * self.scaling)

        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview.Heading", font=("Segoe UI", h_font, "bold"))
        style.configure("Treeview", font=("Segoe UI", b_font), rowheight=r_height) 

        self.tree = ttk.Treeview(self, columns=("Date", "Range", "Rest", "Net", "Type", "Total"), show='headings')
        headers = [("Date", "날짜(Date)"), ("Range", "시간(Time)"), ("Rest", "휴게(Break)"), 
                   ("Net", "실근무(Net)"), ("Type", "유형(Type)"), ("Total", "환산합계")]
        
        for col, name in headers:
            self.tree.heading(col, text=name)
            self.tree.column(col, width=int(150 * self.scaling), anchor="center")
        self.tree.pack(pady=10, fill="both", expand=True, padx=20)

        s_font = int(24 * self.scaling)
        self.summary_box = ctk.CTkTextbox(self, height=int(120 * self.scaling), 
                                          font=("Segoe UI", s_font, "bold"), border_width=2)
        self.summary_box.pack(pady=20, fill="x", padx=20)

    def load_image(self):
        file_path = filedialog.askopenfilename()
        if not file_path: return
        try:
            self.btn_load.configure(text="Detecting Language & Data...", state="disabled")
            self.update()
            
            img = Image.open(file_path).convert('L')
            img = ImageEnhance.Contrast(img).enhance(2.2)
            img = img.point(lambda x: 0 if x < 155 else 255)
            
            # 한글과 영어를 동시에 인식하도록 설정
            raw_text = pytesseract.image_to_string(img, lang='kor+eng', config='--psm 4')
            self.process_ot_data(raw_text)
        except Exception as e:
            messagebox.showerror("Error", str(e))
        finally:
            self.btn_load.configure(text="Load Screenshot (KOR/ENG)", state="normal")

    def process_ot_data(self, raw_text):
        # [업데이트] 한국어(분)와 영어(m)를 모두 추출하는 정규식
        # 날짜 / 시작시간 / 종료시간 / 휴게시간(숫자) 순서 매칭
        pattern = re.compile(r'(\d{1,2}/\d{1,2}).*?(\d{2}:\d{2})\s*-\s*(\d{2}:\d{2}).*?(\d+)\s*(?:분|m|min)', re.S | re.I)
        matches = pattern.findall(raw_text)

        for item in self.tree.get_children(): self.tree.delete(item)
        grand_total_weighted = 0
        current_year = datetime.now().year

        if not matches:
            self.summary_box.delete("0.0", "end")
            self.summary_box.insert("0.0", "⚠️ No data found. Please check image quality.")
            return

        for m in matches:
            date_val, start_s, end_s, rest_m = m
            try:
                # 날짜 및 휴일 판별
                full_date_str = f"{current_year}/{date_val}"
                date_obj = datetime.strptime(full_date_str, "%Y/%m/%d")
                is_holiday = date_obj.weekday() >= 5 or date_obj in self.kr_holidays
                day_name = ["월", "화", "수", "목", "금", "토", "일"][date_obj.weekday()]

                # 시간 계산
                st = datetime.strptime(start_s, "%H:%M")
                en = datetime.strptime(end_s, "%H:%M")
                if en < st: en += timedelta(days=1)
                
                net_h = (en - st).total_seconds() / 3600 - (float(rest_m) / 60)
                
                # 수당 계산
                if is_holiday:
                    type_str = "Holiday"
                    weighted_h = net_h * 1.5 if net_h <= 8 else (8 * 1.5) + ((net_h - 8) * 2.0)
                else:
                    type_str = "Weekday"
                    weighted_h = net_h + (max(0, net_h - 8) * 0.5)
                
                grand_total_weighted += weighted_h

                self.tree.insert("", "end", values=(
                    f"{date_val}({day_name})", f"{start_s}-{end_s}", f"{rest_m}m", 
                    f"{net_h:.1f}h", type_str, f"{weighted_h:.1f}h"
                ))
            except: continue

        self.summary_box.delete("0.0", "end")
        self.summary_box.tag_config("center", justify='center')
        self.summary_box.insert("0.0", f"\nTOTAL WEIGHTED: {grand_total_weighted:.1f} HOURS", "center")

if __name__ == "__main__":
    app = OTCalculator()
    app.mainloop()
