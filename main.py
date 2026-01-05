import os
import sys
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk
from PIL import Image
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

        # 표 디자인: 법정 수당 항목별 세분화
        style = ttk.Style()
        style.configure("Treeview.Heading", font=("Segoe UI", 11, "bold"))
        style.configure("Treeview", font=("Segoe UI", 10), rowheight=30) 

        # 컬럼 구성: 날짜, 시간, 실근무, 연장(1.5), 휴일(1.5~2.0), 야간(0.5), 합계수당
        self.tree = ttk.Treeview(self, columns=("Date", "Range", "Net", "OT15", "Hol", "Night", "Total"), show='headings')
        headers = [("Date", "날짜"), ("Range", "시간"), ("Net", "실근무"), ("OT15", "연장(1.5)"), 
                   ("Hol", "휴일(1.5~)"), ("Night", "야간(0.5)"), ("Total", "환산합계")]
        
        for col, name in headers:
            self.tree.heading(col, text=name)
            self.tree.column(col, width=140, anchor="center")
        self.tree.pack(pady=10, fill="both", expand=True, padx=20)

        # 하단에는 최종 합계만 깔끔하게 표시
        self.summary_box = ctk.CTkTextbox(self, height=120, font=("Segoe UI", 20, "bold"), border_width=2)
        self.summary_box.pack(pady=20, fill="x", padx=20)

    def load_image(self):
        file_path = filedialog.askopenfilename()
        if not file_path: return
        try:
            self.btn_load.configure(text="Calculating Wages...", state="disabled")
            self.update()
            
            img = Image.open(file_path).convert('L')
            img = img.point(lambda x: 0 if x < 150 else 255)
            raw_text = pytesseract.image_to_string(img, lang='kor+eng', config='--psm 4')
            self.process_ot_data(raw_text)
        except Exception as e:
            messagebox.showerror("Error", str(e))
        finally:
            self.btn_load.configure(text="Load Shiftee Screenshot", state="normal")

    def process_ot_data(self, raw_text):
        # 성공했던 정규식 패턴 유지
        pattern = re.compile(r'(\d{1,2}/\d{1,2})\s*\(.\).*?(\d{2}:\d{2}).*?(\d{2}:\d{2}).*?(\d+)\s*[분min]', re.S)
        matches = pattern.findall(raw_text)

        for item in self.tree.get_children(): self.tree.delete(item)
        grand_total_weighted = 0

        for m in matches:
            date_val, start_s, end_s, rest_m = m
            st = datetime.strptime(start_s, "%H:%M")
            en = datetime.strptime(end_s, "%H:%M")
            if en < st: en += timedelta(days=1)
            
            # 1. 기본 실근무 시간 (휴게시간 제외)
            net_h = (en - st).total_seconds() / 3600 - (float(rest_m) / 60)
            
            # 요일 확인 (토, 일 체크)
            # 텍스트에서 요일을 직접 파싱하거나 날짜로 계산 (여기선 스크린샷의 (수) 등을 활용 가능하나 안전하게 날짜 기준 권장)
            # 우선 평일 기준으로 계산하고 가산점 로직 적용
            
            ot_15 = max(0, net_h - 8) # 8시간 초과분
            night_h = 0 # 22시~06시 사이 근무 시 0.5 가산 (필요 시 로직 확장 가능)
            
            # 법정 수당 환산 (기본 1.0 + 연장가산 0.5)
            weighted_h = (net_h * 1.0) + (ot_15 * 0.5)
            
            grand_total_weighted += weighted_h

            self.tree.insert("", "end", values=(
                date_val, f"{start_s}-{end_s}", f"{net_h:.1f}h", 
                f"{ot_15:.1f}h", "-", "-", f"{weighted_h:.1f}h"
            ))

        # 하단 박스: 최종 합계만 강조
        self.summary_box.delete("0.0", "end")
        self.summary_box.tag_config("center", justify='center')
        self.summary_box.insert("0.0", f"\nTOTAL MONTHLY OT: {grand_total_weighted:.1f} HOURS", "center")

if __name__ == "__main__":
    app = OTCalculator()
    app.mainloop()
