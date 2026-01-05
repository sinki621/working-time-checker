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
import ctypes # 4K DPI 인식을 위한 라이브러리

# 1. 윈도우 DPI 인식 설정 (4K 해상도 글자 흐림/작음 방지)
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
        
        # 2. 해상도별 상대적 크기 계산 (DPI Scale 기반)
        # 윈도우 배율이 150%, 200% 등일 때 이를 감지하여 UI 크기 조정
        self.scaling = self.winfo_fpixels('1i') / 96.0
        
        self.title("CSV Chart Viewer - OT Calculator (Producer: KI.Shin)")
        # 창 크기도 배율에 맞게 조절
        win_w = int(1100 * self.scaling)
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
        # 상단 버튼 (배율에 따른 폰트 크기 적용)
        btn_font_size = int(18 * self.scaling)
        self.btn_load = ctk.CTkButton(self, text="Load Shiftee Screenshot", 
                                      command=self.load_image, font=("Segoe UI", btn_font_size, "bold"),
                                      width=int(350 * self.scaling), height=int(55 * self.scaling))
        self.btn_load.pack(pady=20)

        # 3. 표(Treeview) 글씨 크기 가변 설정
        header_font_size = int(12 * self.scaling)
        body_font_size = int(11 * self.scaling)
        row_height = int(35 * self.scaling)

        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview.Heading", font=("Segoe UI", header_font_size, "bold"))
        style.configure("Treeview", font=("Segoe UI", body_font_size), rowheight=row_height) 

        self.tree = ttk.Treeview(self, columns=("Date", "Range", "Rest", "Net", "Type", "Total"), show='headings')
        headers = [("Date", "날짜"), ("Range", "근무시간"), ("Rest", "휴게"), ("Net", "실근무"), ("Type", "근무유형"), ("Total", "환산합계")]
        
        for col, name in headers:
            self.tree.heading(col, text=name)
            self.tree.column(col, width=int(140 * self.scaling), anchor="center")
        self.tree.pack(pady=10, fill="both", expand=True, padx=20)

        # 하단 합계창 폰트 가변 설정
        summary_font_size = int(24 * self.scaling)
        self.summary_box = ctk.CTkTextbox(self, height=int(120 * self.scaling), font=("Segoe UI", summary_font_size, "bold"), border_width=2)
        self.summary_box.pack(pady=20, fill="x", padx=20)

    def load_image(self):
        file_path = filedialog.askopenfilename()
        if not file_path: return
        try:
            self.btn_load.configure(text="Analyzing Holidays & OT...", state="disabled")
            self.update()
            
            img = Image.open(file_path).convert('L')
            img = ImageEnhance.Contrast(img).enhance(2.0)
            img = img.point(lambda x: 0 if x < 160 else 255)
            
            # 성공률 높았던 PSM 4 설정 유지
            raw_text = pytesseract.image_to_string(img, lang='kor+eng', config='--psm 4')
            self.process_ot_data(raw_text)
        except Exception as e:
            messagebox.showerror("Error", str(e))
        finally:
            self.btn_load.configure(text="Load Shiftee Screenshot", state="normal")

    def process_ot_data(self, raw_text):
        pattern = re.compile(r'(\d{1,2}/\d{1,2}).*?(\d{2}:\d{2})\s*-\s*(\d{2}:\d{2}).*?(\d+)\s*[분min]', re.S)
        matches = pattern.findall(raw_text)

        for item in self.tree.get_children(): self.tree.delete(item)
        grand_total_weighted = 0
        current_year = datetime.now().year

        for m in matches:
            date_val, start_s, end_s, rest_m = m
            try:
                full_date_str = f"{current_year}/{date_val}"
                date_obj = datetime.strptime(full_date_str, "%Y/%m/%d")
                
                is_holiday = date_obj.weekday() >= 5 or date_obj in self.kr_holidays
                holiday_name = self.kr_holidays.get(date_obj) if date_obj in self.kr_holidays else ""
                day_name = ["월", "화", "수", "목", "금", "토", "일"][date_obj.weekday()]

                st = datetime.strptime(start_s, "%H:%M")
                en = datetime.strptime(end_s, "%H:%M")
                if en < st: en += timedelta(days=1)
                
                net_h = (en - st).total_seconds() / 3600 - (float(rest_m) / 60)
                
                if is_holiday:
                    type_str = f"휴일({holiday_name if holiday_name else day_name})"
                    weighted_h = net_h * 1.5 if net_h <= 8 else (8 * 1.5) + ((net_h - 8) * 2.0)
                else:
                    type_str = "평일"
                    weighted_h = net_h + (max(0, net_h - 8) * 0.5)
                
                grand_total_weighted += weighted_h

                self.tree.insert("", "end", values=(
                    f"{date_val}({day_name})", f"{start_s}-{end_s}", f"{rest_m}분", 
                    f"{net_h:.1f}h", type_str, f"{weighted_h:.1f}h"
                ))
            except: continue

        self.summary_box.delete("0.0", "end")
        self.summary_box.tag_config("center", justify='center')
        self.summary_box.insert("0.0", f"\nTOTAL WEIGHTED OT: {grand_total_weighted:.1f} HOURS", "center")

if __name__ == "__main__":
    app = OTCalculator()
    app.mainloop()
