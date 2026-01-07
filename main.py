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
# 1. 환경 설정 및 리소스 경로 처리
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
        self.geometry("1600x950")
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
        # 상단 바
        top_bar = ctk.CTkFrame(self, fg_color="transparent")
        top_bar.pack(pady=15, fill="x", padx=20)
        
        # 컨트롤 영역
        ctrl_frame = ctk.CTkFrame(top_bar, fg_color="transparent")
        ctrl_frame.pack(side="left")
        
        ctk.CTkLabel(ctrl_frame, text="Year:", font=("Segoe UI", 14, "bold")).pack(side="left", padx=5)
        self.year_var = ctk.StringVar(value=str(datetime.now().year))
        self.year_dropdown = ctk.CTkComboBox(ctrl_frame, values=["2024", "2025", "2026"], variable=self.year_var, width=90)
        self.year_dropdown.pack(side="left", padx=5)
        
        self.btn_load = ctk.CTkButton(top_bar, text="📁 Load File", command=self.load_image, width=140)
        self.btn_load.pack(side="left", padx=10)
        
        self.btn_paste = ctk.CTkButton(top_bar, text="📋 Paste (Ctrl+V)", command=self.paste_from_clipboard, fg_color="#2ecc71", width=160)
        self.btn_paste.pack(side="left", padx=10)
        
        self.btn_sample = ctk.CTkButton(top_bar, text="💡 Sample", command=self.show_sample, fg_color="#3498db", width=120)
        self.btn_sample.pack(side="right", padx=10)

        # 테이블 스타일 및 구성
        style = ttk.Style()
        style.configure("Treeview", rowheight=30, font=("Segoe UI", 10))
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))
        
        self.tree_frame = ctk.CTkFrame(self)
        self.tree_frame.pack(pady=10, fill="both", expand=True, padx=20)
        
        # 상세 항목 (1.5x, 2.0x, 2.5x) 추가
        self.tree = ttk.Treeview(self.tree_frame, 
                                columns=("Date", "Range", "Break", "NetDiff", "1.5x", "2.0x", "2.5x", "Weighted"), 
                                show='headings')
        
        cols = [
            ("Date", "날짜(요일)", 130), ("Range", "근무시간", 160), ("Break", "휴게", 70), 
            ("NetDiff", "실근무 (+/-)", 110), ("1.5x", "연장/휴일(1.5)", 100), 
            ("2.0x", "휴일연장(2.0)", 100), ("2.5x", "야간/기타(2.5)", 100), ("Weighted", "환산합계", 100)
        ]
        
        for cid, txt, w in cols:
            self.tree.heading(cid, text=txt)
            self.tree.column(cid, width=w, anchor="center")
        
        self.tree.pack(side="left", fill="both", expand=True)
        
        # 하단 요약 박스
        self.summary_box = ctk.CTkTextbox(self, height=150, font=("Segoe UI", 15))
        self.summary_box.pack(pady=15, fill="x", padx=20)

    def show_sample(self):
        sample_path = resource_path("sample.png")
        if not os.path.exists(sample_path):
            messagebox.showwarning("Notice", "sample.png가 루트 폴더에 없습니다.")
            return
        
        top = ctk.CTkToplevel(self)
        top.title("Sample View")
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
            # OCR 전처리
            enhancer = ImageEnhance.Contrast(img.convert('L')).enhance(2.0)
            
            # 언어 자동 감지 (한글 키워드 유무로 판단)
            test_scan = pytesseract.image_to_string(enhancer, lang='kor+eng', config='--psm 3')
            target_lang = 'kor' if any(x in test_scan for x in ['날짜', '근무', '휴게', '시간']) else 'eng'
            
            full_text = pytesseract.image_to_string(enhancer, lang=f'{target_lang}+eng', config='--psm 6')
            self.calculate_data(full_text)
        except Exception as e:
            messagebox.showerror("Error", f"이미지 분석 실패: {e}")

    def calculate_data(self, text):
        # 3자리 휴게시간(\d{1,3}) 및 유연한 단위 인식 정규식
        pattern = re.compile(r'(\d{1,2}/\d{1,2}).*?(\d{2}:\d{2})\s*-\s*(\d{2}:\d{2}).*?(\d{1,3})\s*(?:분|m|min|h|8|B)', re.S | re.I)
        matches = pattern.findall(text)
        
        for item in self.tree.get_children(): self.tree.delete(item)
        
        year = int(self.year_var.get())
        records = []
        
        for d_v, s_t, e_t, r_v in matches:
            try:
                dt = datetime.strptime(f"{year}/{d_v}", "%Y/%m/%d")
                is_h = dt.weekday() >= 5 or dt.strftime('%Y-%m-%d') in kr_holidays
                
                # 시간 계산
                fmt = "%H:%M"
                start, end = datetime.strptime(s_t, fmt), datetime.strptime(e_t, fmt)
                if end < start: end += timedelta(days=1)
                
                # 휴게시간 처리 (3자리수 대응)
                r_val = int(r_v)
                break_h = r_val / 60 if r_val > 5 else r_val # 5보다 크면 분 단위로 간주
                
                net_h = ((end - start).total_seconds() / 3600) - break_h
                records.append({'dt': dt, 'net': net_h, 'is_h': is_h, 'range': f"{s_t}-{e_t}", 'break': f"{r_v}m"})
            except: continue

        records.sort(key=lambda x: x['dt'])
        
        total_weighted_sum = 0
        total_net_sum = 0
        has_holiday_work = False
        
        # 주간 단위 보상 계산을 위한 그룹핑 (ISO 주차 기준)
        for r in records:
            weekday_name = ["월", "화", "수", "목", "금", "토", "일"][r['dt'].weekday()]
            date_str = f"{r['dt'].strftime('%m/%d')} ({weekday_name})"
            
            diff = r['net'] - 8
            diff_display = f"{r['net']:.1f} ({'+' if diff>=0 else ''}{diff:.1f})"
            
            # 배율별 계산 (상세 표기)
            m15, m20, m25 = 0, 0, 0
            
            if not r['is_h']: # 평일
                if r['net'] > 8: m15 = r['net'] - 8
            else: # 휴일/주말
                has_holiday_work = True
                m15 = min(8, r['net'])
                if r['net'] > 8: m20 = r['net'] - 8
            
            weighted_day = (r['net'] if not r['is_h'] else 0) + (m15 * 1.5) + (m20 * 2.0) + (m25 * 2.5)
            
            total_weighted_sum += weighted_day
            total_net_sum += r['net']
            
            self.tree.insert("", "end", values=(
                date_str, r['range'], r['break'], diff_display,
                f"{m15:.1f}" if m15>0 else "-", 
                f"{m20:.1f}" if m20>0 else "-", 
                f"{m25:.1f}" if m25>0 else "-", 
                f"{weighted_day:.1f}h"
            ))

        # 요약 업데이트
        self.summary_box.delete("0.0", "end")
        summary = f"■ 총 실근무 합계: {total_net_sum:.1f} 시간\n"
        summary += f"■ 총 환산 OT 합계 (가중치 적용): {total_weighted_sum:.1f} 시간\n"
        summary += f"■ 주 40시간 대비 정산: {total_net_sum - (len(records)*8):+.1f} 시간 (일별 상쇄 반영)\n"
        
        if has_holiday_work:
            summary += "\n⚠️ [주의] 주말/공휴일 근무가 포함됨: 'Stand-by 근무여부'를 반드시 확인하세요."
            
        self.summary_box.insert("0.0", summary)

if __name__ == "__main__":
    app = OTCalculator()
    app.mainloop()
