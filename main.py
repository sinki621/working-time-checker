import os
import sys
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk
from PIL import Image, ImageOps, ImageFilter
import pytesseract
from datetime import datetime, timedelta

# EXE 빌드 후 내부 리소스(엔진, 데이터) 경로를 동적으로 찾는 함수
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class OTCalculator(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # 1. 창 크기 및 테마 최적화 (평균 해상도 대응)
        self.title("CSV Chart Viewer - OT Calculator (Producer: KI.Shin)")
        self.geometry("1000x750")
        ctk.set_appearance_mode("light")
        
        # Tesseract 엔진 경로 설정
        try:
            engine_root = resource_path("Tesseract-OCR")
            pytesseract.pytesseract.tesseract_cmd = os.path.join(engine_root, "tesseract.exe")
            os.environ["TESSDATA_PREFIX"] = os.path.join(engine_root, "tessdata")
        except:
            pass

        self.setup_ui()

    def setup_ui(self):
        # 상단 로드 버튼 (가독성 위해 폰트 확대)
        self.btn_load = ctk.CTkButton(self, text="Load Shiftee Screenshot", 
                                      command=self.load_image, font=("Segoe UI", 18, "bold"),
                                      width=350, height=55)
        self.btn_load.pack(pady=20)

        # 2. 결과 표(Treeview) 디자인 및 폰트 크기 강화
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview.Heading", font=("Segoe UI", 13, "bold"))
        style.configure("Treeview", font=("Segoe UI", 12), rowheight=35) 

        self.tree = ttk.Treeview(self, columns=("Date", "Range", "Net", "OT", "Total"), show='headings')
        headers = [("Date", "Date"), ("Range", "Time Range"), ("Net", "Work(h)"), ("OT", "OT(h)"), ("Total", "Weighted")]
        for col, name in headers:
            self.tree.heading(col, text=name)
            self.tree.column(col, width=160, anchor="center")
        self.tree.pack(pady=10, fill="both", expand=True, padx=20)

        # 3. 분석 결과 및 디버깅 로그 출력창
        self.summary_box = ctk.CTkTextbox(self, height=180, font=("Consolas", 14, "bold"), border_width=2)
        self.summary_box.pack(pady=20, fill="x", padx=20)

    def load_image(self):
        file_path = filedialog.askopenfilename(filetypes=[("Image Files", "*.png;*.jpg;*.jpeg")])
        if not file_path: return
        
        try:
            self.btn_load.configure(text="Scanning Data...", state="disabled")
            self.update()
            
            # [인식률 개선 핵심 1] 이미지 전처리: 흑백 대비 극대화
            img = Image.open(file_path).convert('L')
            img = img.point(lambda x: 0 if x < 150 else 255) # 글자를 아주 선명하게 만듦
            
            # [인식률 개선 핵심 2] PSM 4 모드: 흩어진 표 형태의 텍스트 읽기에 최적화
            raw_text = pytesseract.image_to_string(img, lang='kor+eng', config='--psm 4')
            
            # 엔진이 실제로 읽은 텍스트를 출력 (문제 진단용)
            self.show_debug_data(raw_text)
            
            self.process_ot_data(raw_text)
        except Exception as e:
            messagebox.showerror("Error", f"Failed: {str(e)}")
        finally:
            self.btn_load.configure(text="Load Shiftee Screenshot", state="normal")

    def show_debug_data(self, text):
        self.summary_box.delete("0.0", "end")
        # 처음 200자만 요약창에 미리 보여주어 글자가 깨지는지 확인
        debug_msg = f"--- [DEBUG: Engine Read Samples] ---\n{text[:200]}\n" + "-"*35 + "\n"
        self.summary_box.insert("0.0", debug_msg)

    def process_ot_data(self, raw_text):
        # [인식률 개선 핵심 3] 초유연 정규식
        # 날짜(숫자/숫자) -> 아무글자 -> 시작시간(:) -> 아무글자 -> 종료시간(:) -> 아무글자 -> 휴게시간(분)
        # 이 구조는 컬럼 사이에 어떤 노이즈가 있어도 숫자만 순서대로 있으면 잡아냅니다.
        pattern = re.compile(r'(\d{1,2}/\d{1,2}).*?(\d{2}:\d{2}).*?(\d{2}:\d{2}).*?(\d+)\s*[분min]', re.S)
        matches = pattern.findall(raw_text)

        for item in self.tree.get_children(): self.tree.delete(item)

        if not matches:
            self.summary_box.insert("end", "⚠️ No Data Found. Check debug log above.\n(Try cropping the image to focus on Date/Time/Break.)")
            return

        total_weighted = 0
        for m in matches:
            date_val, start_s, end_s, rest_m = m
            try:
                st = datetime.strptime(start_s, "%H:%M")
                en = datetime.strptime(end_s, "%H:%M")
                if en < st: en += timedelta(days=1)
                
                net_h = (en - st).total_seconds() / 3600 - (float(rest_m) / 60)
                ot_h = max(0, net_h - 8)
                weighted_h = (ot_h * 1.5)
                total_weighted += weighted_h

                self.tree.insert("", "end", values=(date_val, f"{start_s}-{end_s}", f"{net_h:.1f}h", f"{ot_h:.1f}h", f"{weighted_h:.1f}h"))
            except: continue

        report = f"\n [ ANALYSIS RESULT ]\n" + "━" * 35 + f"\n ▶ Data Rows Detected: {len(matches)}\n ▶ Total Weighted OT: {total_weighted:.1f} hours"
        self.summary_box.insert("end", report)

if __name__ == "__main__":
    app = OTCalculator()
    app.mainloop()
