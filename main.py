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
        
        # 1. 창 크기 최적화 (평균 PC 해상도 대응: 1000x700)
        self.title("CSV Chart Viewer - Producer: KI.Shin")
        self.geometry("1000x700")
        ctk.set_appearance_mode("light")
        
        # 엔진 설정
        try:
            engine_root = resource_path("Tesseract-OCR")
            pytesseract.pytesseract.tesseract_cmd = os.path.join(engine_root, "tesseract.exe")
            os.environ["TESSDATA_PREFIX"] = os.path.join(engine_root, "tessdata")
        except:
            pass

        self.setup_ui()

    def setup_ui(self):
        # 상단 버튼 영역
        self.btn_load = ctk.CTkButton(self, text="Load shiftee screenshot", 
                                      command=self.load_image, font=("Segoe UI", 16, "bold"),
                                      width=300, height=50)
        self.btn_load.pack(pady=15)

        # 2. 가독성을 위한 글꼴 크기 설정 (표 영역)
        self.style = ttk.Style()
        self.style.theme_use("default")
        self.style.configure("Treeview.Heading", font=("Segoe UI", 12, "bold")) # 제목 크게
        self.style.configure("Treeview", font=("Segoe UI", 11), rowheight=30)  # 내용 크게

        # 표 구성
        self.tree = ttk.Treeview(self, columns=("Date", "Range", "Net", "OT", "Total"), show='headings')
        headers = [("Date", "Date(Day)"), ("Range", "Time Range"), ("Net", "Work"), ("OT", "OT(1.5x)"), ("Total", "Weighted")]
        for col, name in headers:
            self.tree.heading(col, text=name)
            self.tree.column(col, width=150, anchor="center")
        self.tree.pack(pady=10, fill="both", expand=True, padx=20)

        # 요약창 글씨 크기 조정
        self.summary_box = ctk.CTkTextbox(self, height=150, font=("Consolas", 14, "bold"))
        self.summary_box.pack(pady=15, fill="x", padx=20)

    def load_image(self):
        file_path = filedialog.askopenfilename()
        if not file_path: return
        try:
            self.btn_load.configure(text="Analyzing Data...", state="disabled")
            self.update()
            
            # OCR 인식률 향상을 위한 설정
            img = Image.open(file_path)
            # 3. 인식률 개선: --psm 6 (단일 텍스트 블록 가정) 모드 사용
            custom_config = r'--oem 3 --psm 6'
            text = pytesseract.image_to_string(img, lang='kor+eng', config=custom_config)
            
            self.process_ot_data(text)
        except Exception as e:
            messagebox.showerror("Error", f"Failed: {str(e)}")
        finally:
            self.btn_load.configure(text="Load shiftee screenshot", state="normal")

    def process_ot_data(self, raw_text):
        # 4. 시프티 스크린샷 전용 유연한 정규식
        # 날짜(요일), 시작시간, 종료시간, 휴게분 순서로 매칭
        pattern = re.compile(r'(\d{1,2}/\d{1,2})\s*\(.\)\s*(\d{2}:\d{2})\s*-\s*(\d{2}:\d{2}).*?(\d+)분', re.S)
        matches = pattern.findall(raw_text)

        for item in self.tree.get_children(): self.tree.delete(item)

        total_weighted = 0
        if not matches:
            self.summary_box.insert("0.0", "Error: Could not read data. Please check image quality.")
            return

        for m in matches:
            date, start_s, end_s, rest_m = m
            st = datetime.strptime(start_s, "%H:%M")
            en = datetime.strptime(end_s, "%H:%M")
            if en < st: en += timedelta(days=1)
            
            net_h = (en - st).total_seconds() / 3600 - (float(rest_m) / 60)
            ot_h = max(0, net_h - 8)
            weighted_h = (ot_h * 1.5) # 기본 연장 수당만 계산 (휴일 로직 필요시 추가)
            
            total_weighted += weighted_h

            self.tree.insert("", "end", values=(date, f"{start_s}-{end_s}", f"{net_h:.1f}h", f"{ot_h:.1f}h", f"{weighted_h:.1f}h"))

        self.summary_box.delete("0.0", "end")
        self.summary_box.insert("0.0", f" [ MONTHLY OT SUMMARY ]\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n Total Weighted OT: {total_weighted:.1f} hours\n (Calculated based on 8h base work)")

if __name__ == "__main__":
    app = OTCalculator()
    app.mainloop()
