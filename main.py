import os
import sys
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk
from PIL import Image, ImageOps, ImageFilter
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
        self.title("CSV Chart Viewer - Producer: KI.Shin")
        self.geometry("1000(cite: 750") # 가시성 좋은 창 크기
        ctk.set_appearance_mode("light")
        
        try:
            engine_root = resource_path("Tesseract-OCR")
            pytesseract.pytesseract.tesseract_cmd = os.path.join(engine_root, "tesseract.exe")
            os.environ["TESSDATA_PREFIX"] = os.path.join(engine_root, "tessdata")
        except: pass

        self.setup_ui()

    def setup_ui(self):
        self.btn_load = ctk.CTkButton(self, text="Load shiftee screenshot", 
                                      command=self.load_image, font=("Segoe UI", 18, "bold"),
                                      width=350, height=55)
        self.btn_load.pack(pady=20)

        # 표 글씨 크기 대폭 확대 (가독성 해결)
        style = ttk.Style()
        style.configure("Treeview.Heading", font=("Segoe UI", 13, "bold"))
        style.configure("Treeview", font=("Segoe UI", 12), rowheight=35) 

        self.tree = ttk.Treeview(self, columns=("Date", "Range", "Net", "OT", "Total"), show='headings')
        headers = [("Date", "날짜"), ("Range", "시간"), ("Net", "실근무"), ("OT", "연장(OT)"), ("Total", "가산결과")]
        for col, name in headers:
            self.tree.heading(col, text=name)
            self.tree.column(col, width=160, anchor="center")
        self.tree.pack(pady=10, fill="both", expand=True, padx=20)

        self.summary_box = ctk.CTkTextbox(self, height=160, font=("Consolas", 15, "bold"), border_width=2)
        self.summary_box.pack(pady=20, fill="x", padx=20)

    def load_image(self):
        file_path = filedialog.askopenfilename()
        if not file_path: return
        try:
            self.btn_load.configure(text="Deep Scanning Data...", state="disabled")
            self.update()
            
            # [인식률 개선 핵심 1] 이미지 전처리
            img = Image.open(file_path).convert('L') # 흑백 변환
            img = img.point(lambda x: 0 if x < 150 else 255) # 대비 강화 (글자 선명하게)
            
            # [인식률 개선 핵심 2] PSM 4 (가변 길이 컬럼의 텍스트 뭉치 읽기) 사용
            custom_config = r'--oem 3 --psm 4'
            raw_text = pytesseract.image_to_string(img, lang='kor+eng', config=custom_config)
            
            self.process_ot_data(raw_text)
        except Exception as e:
            messagebox.showerror("Error", f"Failed: {str(e)}")
        finally:
            self.btn_load.configure(text="Load shiftee screenshot", state="normal")

    def process_ot_data(self, raw_text):
        # [인식률 개선 핵심 3] 정규식 파격 완화 
        # (중간에 어떤 글자가 섞여있어도 날짜 -> 시간 -> 시간 -> 분 순서만 맞으면 추출)
        pattern = re.compile(r'(\d{1,2}/\d{1,2}).*?(\d{2}:\d{2}).*?(\d{2}:\d{2}).*?(\d+)분', re.S)
        matches = pattern.findall(raw_text)

        for item in self.tree.get_children(): self.tree.delete(item)
        self.summary_box.delete("0.0", "end")

        if not matches:
            self.summary_box.insert("0.0", "⚠️ 데이터를 찾지 못했습니다.\n1. 스크린샷에 '날짜/시간/휴게시간'이 모두 포함되어 있는지 확인하세요.\n2. 이미지 화질이 너무 낮으면 인식이 안 될 수 있습니다.")
            return

        total_weighted = 0
        for m in matches:
            date, start_s, end_s, rest_m = m
            st = datetime.strptime(start_s, "%H:%M")
            en = datetime.strptime(end_s, "%H:%M")
            if en < st: en += timedelta(days=1)
            
            net_h = (en - st).total_seconds() / 3600 - (float(rest_m) / 60)
            ot_h = max(0, net_h - 8)
            weighted_h = (ot_h * 1.5)
            total_weighted += weighted_h

            self.tree.insert("", "end", values=(date, f"{start_s}-{end_s}", f"{net_h:.1f}h", f"{ot_h:.1f}h", f"{weighted_h:.1f}h"))

        self.summary_box.insert("0.0", f" [ MONTHLY OT SUMMARY - KI.Shin ]\n" + "━" * 35 + f"\n ▶ Total Weighted OT: {total_weighted:.1f} hours\n ▶ Data Detected: {len(matches)} rows")

if __name__ == "__main__":
    app = OTCalculator()
    app.mainloop()
