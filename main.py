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
        self.title("CSV Chart Viewer - OT 상세 분석기 (Producer: KI.Shin)")
        self.geometry("1200x900")
        ctk.set_appearance_mode("light")
        
        # 엔진 경로 설정
        engine_root = resource_path("Tesseract-OCR")
        pytesseract.pytesseract.tesseract_cmd = os.path.join(engine_root, "tesseract.exe")
        os.environ["TESSDATA_PREFIX"] = os.path.join(engine_root, "tessdata")

        self.setup_ui()

    def setup_ui(self):
        self.btn_load = ctk.CTkButton(self, text="근무표 스크린샷 로드 (PNG, JPG)", 
                                      command=self.load_image, font=("Segoe UI", 15, "bold"),
                                      width=400, height=60)
        self.btn_load.pack(pady=20)

        # 표 구성 (상세 내역)
        self.tree = ttk.Treeview(self, columns=("Date", "Time", "Net", "Base", "OT", "Hol", "Weighted"), show='headings')
        headers = [("Date", "날짜(요일)"), ("Time", "시간 범위"), ("Net", "순근무"), ("Base", "기본(8h)"), ("OT", "연장"), ("Hol", "휴일"), ("Weighted", "수당가산")]
        widths = [130, 150, 100, 100, 100, 100, 120]
        
        for i, (col, name) in enumerate(headers):
            self.tree.heading(col, text=name)
            self.tree.column(col, width=widths[i], anchor="center")
        self.tree.pack(pady=10, fill="both", expand=True, padx=20)

        self.summary_box = ctk.CTkTextbox(self, height=200, font=("Consolas", 13))
        self.summary_box.pack(pady=20, fill="x", padx=20)

    def load_image(self):
        file_path = filedialog.askopenfilename()
        if not file_path: return
        try:
            self.btn_load.configure(text="데이터 정밀 분석 중...", state="disabled")
            self.update()
            
            # kor+eng 모드로 텍스트 추출
            text = pytesseract.image_to_string(Image.open(file_path), lang='kor+eng')
            self.process_detailed_ot(text)
        except Exception as e:
            messagebox.showerror("Error", f"인식 실패: {str(e)}")
        finally:
            self.btn_load.configure(text="근무표 스크린샷 로드 (PNG, JPG)", state="normal")

    def process_detailed_ot(self, text):
        # 날짜, 요일, 시작, 종료, 휴게시간 추출 (스크린샷 구조 대응)
        pattern = re.compile(r'(\d{1,2}/\d{1,2})\s*\((.)\)\s*(\d{2}:\d{2})\s*-\s*(\d{2}:\d{2}).*?(\d+)분', re.S)
        matches = pattern.findall(text)

        for item in self.tree.get_children(): self.tree.delete(item)

        totals = {"net": 0, "ot": 0, "hol": 0, "weighted": 0}

        for m in matches:
            date, day, start_s, end_s, rest_m = m
            st = datetime.strptime(start_s, "%H:%M")
            en = datetime.strptime(end_s, "%H:%M")
            if en < st: en += timedelta(days=1)
            
            # 1. 시간 계산
            total_work = (en - st).total_seconds() / 3600
            net_work = total_work - (float(rest_m) / 60)
            is_holiday = day in ["토", "일"]
            
            # 2. OT 세분화 로직
            base_8 = min(8, net_work) if not is_holiday else 0
            ot_hours = max(0, net_work - 8) if not is_holiday else 0
            holiday_hours = net_work if is_holiday else 0
            
            # 3. 가산 수당 계산 (연장 1.5, 휴일 1.5, 휴일연장 2.0 등)
            weighted_ot = (ot_hours * 1.5) + (holiday_hours * 1.5)
            if is_holiday and net_work > 8: # 휴일 연장 추가 가산
                weighted_ot += (net_work - 8) * 0.5

            # 합산
            totals["net"] += net_work
            totals["ot"] += ot_hours
            totals["hol"] += holiday_hours
            totals["weighted"] += weighted_ot

            self.tree.insert("", "end", values=(
                f"{date}({day})", f"{start_s}-{end_s}", f"{net_work:.2f}h", 
                f"{base_8:.1f}h", f"{ot_hours:.1f}h", f"{holiday_hours:.1f}h", f"{weighted_ot:.2f}h"
            ))

        summary = f""" [ OT 및 근무 수당 정밀 분석 리포트 ]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  ▶ 총 근무 시간 (휴게 제외) : {totals['net']:.2f} 시간
  ▶ 평일 연장 근로 시간 (OT) : {totals['ot']:.1f} 시간
  ▶ 휴일 근로 시간 (Holiday) : {totals['hol']:.1f} 시간
  -----------------------------------------------------------
  ▶ [수당 합산] 가산 적용 OT  : {totals['weighted']:.2f} 시간
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
* 평일 8시간 초과분은 1.5배, 휴일 근로는 전체 1.5배가 적용되었습니다."""
        
        self.summary_box.delete("0.0", "end")
        self.summary_box.insert("0.0", summary)

if __name__ == "__main__":
    app = OTCalculator()
    app.mainloop()
