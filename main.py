import os
import sys
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk
from PIL import Image, ImageEnhance
import pytesseract
from datetime import datetime, timedelta
import holidays # 공휴일 판별을 위한 라이브러리

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
        
        # 대한민국 공휴일 정보 로드 (2025-2026년 포함)
        self.kr_holidays = holidays.KR()
        
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

        # 표 구성: 휴일 여부 확인을 위한 비고란 추가
        self.tree = ttk.Treeview(self, columns=("Date", "Range", "Rest", "Net", "Type", "Total"), show='headings')
        headers = [("Date", "날짜"), ("Range", "근무시간"), ("Rest", "휴게"), ("Net", "실근무"), ("Type", "근무유형"), ("Total", "환산합계")]
        
        for col, name in headers:
            self.tree.heading(col, text=name)
            self.tree.column(col, width=140, anchor="center")
        self.tree.pack(pady=10, fill="both", expand=True, padx=20)

        self.summary_box = ctk.CTkTextbox(self, height=120, font=("Segoe UI", 24, "bold"), border_width=2)
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
        current_year = datetime.now().year # 현재 연도 기준 (필요 시 수정)

        for m in matches:
            date_val, start_s, end_s, rest_m = m
            try:
                # 1. 공휴일 및 주말 판별
                full_date_str = f"{current_year}/{date_val}"
                date_obj = datetime.strptime(full_date_str, "%Y/%m/%d")
                
                # 주말이거나 법정 공휴일 이름이 holidays에 존재하면 True
                is_holiday = date_obj.weekday() >= 5 or date_obj in self.kr_holidays
                holiday_name = self.kr_holidays.get(date_obj) if date_obj in self.kr_holidays else ""
                day_name = ["월", "화", "수", "목", "금", "토", "일"][date_obj.weekday()]

                # 2. 시간 계산
                st = datetime.strptime(start_s, "%H:%M")
                en = datetime.strptime(end_s, "%H:%M")
                if en < st: en += timedelta(days=1)
                
                net_h = (en - st).total_seconds() / 3600 - (float(rest_m) / 60)
                
                # 3. 법정 수당 적용 (근로기준법)
                if is_holiday:
                    # 휴일 근로: 8시간까지 1.5배, 초과분 2.0배
                    type_str = f"휴일({holiday_name if holiday_name else day_name})"
                    if net_h <= 8:
                        weighted_h = net_h * 1.5
                    else:
                        weighted_h = (8 * 1.5) + ((net_h - 8) * 2.0)
                else:
                    # 평일 근로: 8시간 초과분 1.5배 (기본 1.0 + 가산 0.5)
                    type_str = "평일"
                    ot_h = max(0, net_h - 8)
                    weighted_h = net_h + (ot_h * 0.5)
                
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
