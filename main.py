import os
import sys
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk
from PIL import Image, ImageEnhance, ImageGrab, ImageFilter, ImageTk
import pytesseract
from datetime import datetime, timedelta
import ctypes
import io

# =============================================================================
# 1. 환경 설정 및 라이브러리 예외 처리
# =============================================================================

# PyInstaller 빌드 환경에서 라이브러리 경로 인식 강화
if getattr(sys, 'frozen', False):
    os.environ['PATH'] = sys._MEIPASS + os.pathsep + os.environ.get('PATH', '')

# 윈도우 DPI 인식 설정 (고해상도 대응)
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except:
    try: ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except:
        try: ctypes.windll.user32.SetProcessDPIAware()
        except: pass

# holidays 라이브러리 안전 로드
try:
    import holidays
    try:
        kr_holidays = holidays.country_holidays('KR')
    except:
        kr_holidays = holidays.KR()
except ImportError:
    kr_holidays = {}

def resource_path(relative_path):
    """PyInstaller 리소스 경로 변환"""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# =============================================================================
# 2. 메인 애플리케이션 클래스
# =============================================================================
class OTCalculator(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # UI 배율 고정
        ctk.set_widget_scaling(1.0)
        ctk.set_window_scaling(1.0)
        
        self.title("CSV Chart Viewer - OT Calculator (Producer: KI.Shin)")
        self.geometry("1400x850")
        self.minsize(1200, 700)
        ctk.set_appearance_mode("light")
        
        self.kr_holidays = kr_holidays
        self.setup_tesseract()
        self.setup_ui()
        
        self.bind('<Control-v>', self.paste_from_clipboard)
        self.bind('<Control-V>', self.paste_from_clipboard)

    def setup_tesseract(self):
        """Tesseract OCR 경로 설정"""
        try:
            engine_root = resource_path("Tesseract-OCR")
            tesseract_exe = os.path.join(engine_root, "tesseract.exe")
            tessdata_dir = os.path.join(engine_root, "tessdata")
            
            if os.path.exists(tesseract_exe):
                pytesseract.pytesseract.tesseract_cmd = tesseract_exe
                os.environ["TESSDATA_PREFIX"] = tessdata_dir
            else:
                pytesseract.pytesseract.tesseract_cmd = "tesseract"
        except Exception as e:
            print(f"Tesseract Setup Warning: {e}")

    def setup_ui(self):
        # 상단 컨트롤 영역
        top_frame = ctk.CTkFrame(self, fg_color="transparent")
        top_frame.pack(pady=15)
        
        ctk.CTkLabel(top_frame, text="Year:", font=("Segoe UI", 14, "bold")).pack(side="left", padx=(0, 10))
        
        self.year_var = ctk.StringVar(value="2025")
        self.year_dropdown = ctk.CTkComboBox(top_frame, values=["2024", "2025", "2026"], variable=self.year_var, width=100, state="readonly")
        self.year_dropdown.pack(side="left", padx=(0, 20))
        
        self.btn_load = ctk.CTkButton(top_frame, text="📁 Load File", command=self.load_image, font=("Segoe UI", 16, "bold"), width=180, height=50)
        self.btn_load.pack(side="left", padx=(0, 10))
        
        self.btn_paste = ctk.CTkButton(top_frame, text="📋 Paste (Ctrl+V)", command=lambda: self.paste_from_clipboard(None), font=("Segoe UI", 16, "bold"), width=200, height=50, fg_color="#2ecc71")
        self.btn_paste.pack(side="left", padx=(0, 10))

        # 데이터 테이블 영역
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview.Heading", font=("Segoe UI", 11, "bold"))
        style.configure("Treeview", font=("Segoe UI", 10), rowheight=28)

        tree_frame = ctk.CTkFrame(self, fg_color="white")
        tree_frame.pack(pady=10, fill="both", expand=True, padx=20)
        
        self.tree = ttk.Treeview(tree_frame, columns=("Date", "Range", "Rest", "Net", "Type", "Total"), show='headings')
        cols = [("Date", "날짜", 120), ("Range", "근무시간", 180), ("Rest", "휴게", 80), ("Net", "실근무", 90), ("Type", "유형", 120), ("Total", "환산합계", 100)]
        for cid, head, w in cols:
            self.tree.heading(cid, text=head)
            self.tree.column(cid, width=w, anchor="center")
        self.tree.pack(side="left", fill="both", expand=True)

        # 하단 요약 영역
        self.summary_box = ctk.CTkTextbox(self, height=100, font=("Segoe UI", 18, "bold"))
        self.summary_box.pack(pady=15, fill="x", padx=20)

    def load_image(self):
        f_path = filedialog.askopenfilename(filetypes=[("Images", "*.png *.jpg *.jpeg *.bmp")])
        if f_path:
            self.process_image(Image.open(f_path))

    def paste_from_clipboard(self, event):
        img = ImageGrab.grabclipboard()
        if isinstance(img, Image.Image):
            self.process_image(img)
        elif isinstance(img, list) and len(img) > 0:
            self.process_image(Image.open(img[0]))

    def process_image(self, img):
        try:
            img_gray = img.convert('L')
            img_enh = ImageEnhance.Contrast(img_gray).enhance(2.5)
            img_bin = img_enh.point(lambda x: 0 if x < 145 else 255)
            
            # OCR 수행 (한글+영어)
            text = pytesseract.image_to_string(img_bin, lang='kor+eng', config='--psm 6')
            self.parse_data(text)
        except Exception as e:
            messagebox.showerror("OCR Error", str(e))

    def parse_data(self, text):
        # 데이터 파싱 로직 (날짜 시간 휴게 매칭)
        pattern = re.compile(r'(\d{1,2}/\d{1,2}).*?(\d{2}:\d{2})\s*-\s*(\d{2}:\d{2}).*?(\d+)\s*(?:분|m|min)', re.S | re.I)
        matches = pattern.findall(text)
        
        for item in self.tree.get_children(): self.tree.delete(item)
        total_h = 0
        year = int(self.year_var.get())

        for d_v, s_t, e_t, r_m in matches:
            try:
                dt = datetime.strptime(f"{year}/{d_v}", "%Y/%m/%d")
                is_h = dt.weekday() >= 5 or dt in self.kr_holidays
                st_obj = datetime.strptime(s_t, "%H:%M")
                en_obj = datetime.strptime(e_t, "%H:%M")
                if en_obj < st_obj: en_obj += timedelta(days=1)
                
                net = (en_obj - st_obj).total_seconds() / 3600 - (int(r_m)/60)
                weighted = net * 1.5 if is_h else net + (max(0, net-8)*0.5)
                total_h += weighted
                
                self.tree.insert("", "end", values=(d_v, f"{s_t}-{e_t}", f"{r_m}m", f"{net:.1f}h", "Holiday" if is_h else "Weekday", f"{weighted:.1f}h"))
            except: continue
            
        self.summary_box.delete("0.0", "end")
        self.summary_box.insert("0.0", f"계산 결과: 총 환산 OT 합계 = {total_h:.1f} 시간")

if __name__ == "__main__":
    app = OTCalculator()
    app.mainloop()
