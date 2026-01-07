import os
import sys
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk
from PIL import Image, ImageEnhance, ImageGrab, ImageTk
import pytesseract
from datetime import datetime, timedelta
import ctypes

# =============================================================================
# 1. 환경 설정 및 라이브러리 예외 처리
# =============================================================================

if getattr(sys, 'frozen', False):
    os.environ['PATH'] = sys._MEIPASS + os.pathsep + os.environ.get('PATH', '')

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except:
    try: ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except: pass

try:
    import holidays
    kr_holidays = holidays.KR()
except:
    kr_holidays = {}

def resource_path(relative_path):
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
        
        self.title("CSV Chart Viewer - OT Calculator (Producer: KI.Shin)")
        self.geometry("1500x900")
        ctk.set_appearance_mode("light")
        
        self.setup_tesseract()
        self.setup_ui()
        
        self.bind('<Control-v>', lambda e: self.paste_from_clipboard())
        self.bind('<Control-V>', lambda e: self.paste_from_clipboard())

    def setup_tesseract(self):
        engine_root = resource_path("Tesseract-OCR")
        tesseract_exe = os.path.join(engine_root, "tesseract.exe")
        if os.path.exists(tesseract_exe):
            pytesseract.pytesseract.tesseract_cmd = tesseract_exe
            os.environ["TESSDATA_PREFIX"] = os.path.join(engine_root, "tessdata")

    def setup_ui(self):
        # 상단 컨트롤 레이아웃
        top_frame = ctk.CTkFrame(self, fg_color="transparent")
        top_frame.pack(pady=15, fill="x", padx=20)
        
        left_ctrl = ctk.CTkFrame(top_frame, fg_color="transparent")
        left_ctrl.pack(side="left")
        
        ctk.CTkLabel(left_ctrl, text="Year:", font=("Segoe UI", 14, "bold")).pack(side="left", padx=5)
        self.year_var = ctk.StringVar(value=str(datetime.now().year))
        self.year_dropdown = ctk.CTkComboBox(left_ctrl, values=["2024", "2025", "2026"], variable=self.year_var, width=90)
        self.year_dropdown.pack(side="left", padx=5)
        
        self.btn_load = ctk.CTkButton(top_frame, text="📁 Load File", command=self.load_image, width=140)
        self.btn_load.pack(side="left", padx=10)
        
        self.btn_paste = ctk.CTkButton(top_frame, text="📋 Paste (Ctrl+V)", command=self.paste_from_clipboard, fg_color="#2ecc71", width=160)
        self.btn_paste.pack(side="left", padx=10)
        
        self.btn_sample = ctk.CTkButton(top_frame, text="💡 Sample", command=self.show_sample, fg_color="#3498db", width=120)
        self.btn_sample.pack(side="right", padx=10)

        # 테이블 구성
        style = ttk.Style()
        style.configure("Treeview", rowheight=30, font=("Segoe UI", 10))
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))
        
        self.tree_frame = ctk.CTkFrame(self)
        self.tree_frame.pack(pady=10, fill="both", expand=True, padx=20)
        
        self.tree = ttk.Treeview(self.tree_frame, columns=("Date", "Range", "Break", "NetDiff", "Details", "Weighted"), show='headings')
        
        headers = [("Date", "날짜 (요일)", 150), ("Range", "근무시간", 180), ("Break", "휴게", 80), 
                   ("NetDiff", "실근무 (+/-)", 120), ("Details", "배율 상세", 250), ("Weighted", "환산합계", 100)]
        
        for cid, txt, w in headers:
            self.tree.heading(cid, text=txt)
            self.tree.column(cid, width=w, anchor="center")
        
        self.tree.pack(side="left", fill="both", expand=True)
        
        # 요약 정보창
        self.summary_box = ctk.CTkTextbox(self, height=120, font=("Segoe UI", 16))
        self.summary_box.pack(pady=15, fill="x", padx=20)

    def show_sample(self):
        """Root 폴더의 sample.png를 원본 해상도로 표시"""
        sample_path = resource_path("sample.png")
        if not os.path.exists(sample_path):
            messagebox.showwarning("Warning", "sample.png 파일을 찾을 수 없습니다.")
            return
        
        top = ctk.CTkToplevel(self)
        top.title("Sample Image")
        img = Image.open(sample_path)
        img_tk = ImageTk.PhotoImage(img)
        
        label = tk.Label(top, image=img_tk)
        label.image = img_tk 
        label.pack()

    def load_image(self):
        f = filedialog.askopenfilename()
        if f: self.process_image(Image.open(f))

    def paste_from_clipboard(self):
        img = ImageGrab.grabclipboard()
        if isinstance(img, Image.Image): self.process_image(img)

    def process_image(self, img):
        try:
            # 전처리
            gray = img.convert('L')
            enhancer = ImageEnhance.Contrast(gray).enhance(2.0)
            
            # 1. 언어 감지용 사전 스캔
            temp_text = pytesseract.image_to_string(enhancer, lang='kor+eng', config='--psm 3')
            lang = 'kor' if any(k in temp_text for k in ['날짜', '근무', '휴게']) else 'eng'
            
            # 2. 본 OCR 수행
            full_text = pytesseract.image_to_string(enhancer, lang=f'{lang}+eng', config='--psm 6')
            self.parse_and_calculate(full_text)
        except Exception as e:
            messagebox.showerror("Error", f"인식 실패: {e}")

    def parse_and_calculate(self, text):
        # 정규식: 날짜, 시간범위, 휴게시간(분/시간) 추출 강화
        # 예: 12/31, 06:50-03:40, 120m 또는 2h
        pattern = re.compile(r'(\d{1,2}/\d{1,2}).*?(\d{2}:\d{2})\s*-\s*(\d{2}:\d{2}).*?(\d+)\s*(?:분|m|min|h)', re.S | re.I)
        matches = pattern.findall(text)
        
        for item in self.tree.get_children(): self.tree.delete(item)
        
        data_list = []
        year = int(self.year_var.get())
        has_standby = False

        for d_v, s_t, e_t, r_v in matches:
            try:
                dt = datetime.strptime(f"{year}/{d_v}", "%Y/%m/%d")
                weekday_str = ["월", "화", "수", "목", "금", "토", "일"][dt.weekday()]
                is_holiday = dt.weekday() >= 5 or dt.strftime('%Y-%m-%d') in kr_holidays
                
                # 시간 계산
                fmt = "%H:%M"
                start = datetime.strptime(s_t, fmt)
                end = datetime.strptime(e_t, fmt)
                if end < start: end += timedelta(days=1)
                
                raw_duration = (end - start).total_seconds() / 3600
                
                # 휴게시간 인식 (m단위 가정, h단위일 경우 조건 추가 가능)
                break_h = int(r_v) / 60 if int(r_v) > 5 else int(r_v)
                net_h = raw_duration - break_h
                
                data_list.append({
                    'dt': dt, 'date_str': f"{d_v} ({weekday_str})", 'range': f"{s_t}-{e_t}",
                    'break': f"{r_v}m", 'net': net_h, 'is_h': is_holiday
                })
                if is_holiday: has_standby = True
            except: continue

        # 주차별 그룹화 및 계산 (월요일 기준)
        data_list.sort(key=lambda x: x['dt'])
        total_ot = 0
        total_weighted = 0
        
        for item in data_list:
            # 1. 일일 편차 (+/- 8시간 기준)
            diff = item['net'] - 8
            diff_str = f"{item['net']:.1f}h ({'+' if diff>=0 else ''}{diff:.1f})"
            
            # 2. 상세 배율 및 가중치 계산
            details = []
            day_weighted = 0
            
            if not item['is_h']: # 평일
                day_weighted += min(8, item['net']) # 기본 8시간
                if item['net'] > 8:
                    ot8 = item['net'] - 8
                    day_weighted += (ot8 * 1.5)
                    details.append(f"연장 {ot8:.1f}x1.5")
            else: # 휴일/주말
                if item['net'] <= 8:
                    day_weighted += (item['net'] * 1.5)
                    details.append(f"휴일 {item['net']:.1f}x1.5")
                else:
                    day_weighted += (8 * 1.5) + ((item['net'] - 8) * 2.0)
                    details.append(f"휴일 8x1.5 + 연장 {(item['net']-8):.1f}x2.0")
            
            total_weighted += day_weighted
            total_ot += max(0, item['net'] - 8) if not item['is_h'] else item['net']
            
            self.tree.insert("", "end", values=(
                item['date_str'], item['range'], item['break'], 
                diff_str, ", ".join(details), f"{day_weighted:.1f}h"
            ))

        # 요약 업데이트
        self.summary_box.delete("0.0", "end")
        summary_text = f"▶ 총 연장/휴일 근무 시간: {total_ot:.1f} 시간\n"
        summary_text += f"▶ 최종 환산 OT 합계 (가중치 적용): {total_weighted:.1f} 시간\n"
        if has_standby:
            summary_text += "\n⚠️ [공휴일/주말 근무 감지] 해당 일자 Stand-by 근무여부를 반드시 확인하세요."
        
        self.summary_box.insert("0.0", summary_text)

if __name__ == "__main__":
    app = OTCalculator()
    app.mainloop()
