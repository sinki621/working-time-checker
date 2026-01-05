import os
import sys
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
try:
    import customtkinter as ctk
except ImportError:
    print("Error: 'customtkinter' module not found. Please run 'pip install customtkinter'")
    sys.exit(1)
from PIL import Image, ImageEnhance
import pytesseract
from datetime import datetime, timedelta
import holidays
import ctypes

# 윈도우 고해상도(4K) DPI 인식 설정
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    try: ctypes.windll.user32.SetProcessDPIAware()
    except: pass

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class OTCalculator(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # 해상도 배율 감지 (4K 가독성 핵심)
        self.scaling = self.winfo_fpixels('1i') / 96.0
        self.title("CSV Chart Viewer - OT Calculator (Producer: KI.Shin)")
        
        # 배율에 따른 창 크기 설정
        self.geometry(f"{int(1150 * self.scaling)}x{int(850 * self.scaling)}")
        ctk.set_appearance_mode("light")
        
        self.kr_holidays = holidays.KR()
        
        # Tesseract 엔진 설정
        try:
            engine_root = resource_path("Tesseract-OCR")
            pytesseract.pytesseract.tesseract_cmd = os.path.join(engine_root, "tesseract.exe")
            os.environ["TESSDATA_PREFIX"] = os.path.join(engine_root, "tessdata")
        except:
            pass

        self.setup_ui()

    def setup_ui(self):
        # 버튼 폰트 및 크기 조절
        btn_font = int(18 * self.scaling)
        self.btn_load = ctk.CTkButton(self, text="Load Screenshot (한/영 자동인식)", 
                                      command=self.load_image, font=("Segoe UI", btn_font, "bold"),
                                      width=int(400 * self.scaling), height=int(60 * self.scaling))
        self.btn_load.pack(pady=20)

        # 표(Treeview) 스타일 및 폰트 조절
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
            self.tree.column(col, width=int(160 * self.scaling), anchor="center")
        self.tree.pack(pady=10, fill="both", expand=True, padx=20)

        # 요약 박스 폰트 조절
        s_font = int(24 * self.scaling)
        self.summary_box = ctk.CTkTextbox(self, height=int(130 * self.scaling), 
                                          font=("Segoe UI", s_font, "bold"), border_width=2)
        self.summary_box.pack(pady=20, fill="x", padx=20)

    def load_image(self):
        file_path = filedialog.askopenfilename()
        if not file_path: return
        try:
            self.btn_load.configure(text="Processing...", state="disabled")
            self.update()
            
            # OCR 인식률 향상 처리
            img = Image.open(file_path).convert('L')
            img = ImageEnhance.Contrast(img).enhance(2.2)
            img = img.point(lambda x: 0 if x < 155 else 255)
            
            # 한/영 통합 엔진 가동
            raw_text = pytesseract.image_to_string(img, lang='kor+eng', config='--psm 4')
            self.process_ot_data(raw_text)
        except Exception as e:
            messagebox.showerror("Error", f"인식 실패: {str(e)}")
        finally:
            self.btn_load.configure(text="Load Screenshot (한/영 자동인식)", state="normal")

    def process_ot_data(self, raw_text):
        # 한글(분)과 영어(m/min)를 모두 잡는 정규식
        pattern = re.compile(r'(\d{1,2}/\d{1,2}).*?(\d{2}:\d{2})\s*-\s*(\d{2}:\d{2}).*?(\d+)\s*(?:분|m|min)', re.S | re.I)
        matches = pattern.findall(raw_text)

        for item in self.tree.get_children(): self.tree.delete(item)
        total_w = 0
        current_year = datetime.now().year

        if not matches:
            self.summary_box.delete("0.0", "end")
            self.summary_box.insert("0.0", "⚠️ 데이터를 찾을 수 없습니다. (Check Image Quality)")
            return

        for m in matches:
            d_val, s_s, e_s, r_m = m
            try:
                dt_obj = datetime.strptime(f"{current_year}/{d_val}", "%Y/%m/%d")
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
        self.summary_box.tag_config("center", justify='center')
        self.summary_box.insert("0.0", f"\nTOTAL WEIGHTED OT: {total_w:.1f} HOURS", "center")

if __name__ == "__main__":
    app = OTCalculator()
    app.mainloop()
