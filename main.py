import os
import sys
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk
from PIL import Image, ImageEnhance, ImageGrab, ImageTk, ImageOps
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
        self.geometry("1600(950")
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
        top_bar = ctk.CTkFrame(self, fg_color="transparent")
        top_bar.pack(pady=15, fill="x", padx=20)
        
        ctrl_frame = ctk.CTkFrame(top_bar, fg_color="transparent")
        ctrl_frame.pack(side="left")
        
        ctk.CTkLabel(ctrl_frame, text="Year:", font=("Segoe UI", 14, "bold")).pack(side="left", padx=5)
        self.year_var = ctk.StringVar(value=str(datetime.now().year))
        self.year_dropdown = ctk.CTkComboBox(ctrl_frame, values=["2024", "2025", "2026", "2027"], variable=self.year_var, width=90)
        self.year_dropdown.pack(side="left", padx=5)
        
        self.btn_load = ctk.CTkButton(top_bar, text="📁 Load File", command=self.load_image, width=140)
        self.btn_load.pack(side="left", padx=10)
        
        self.btn_paste = ctk.CTkButton(top_bar, text="📋 Paste (Ctrl+V)", command=self.paste_from_clipboard, fg_color="#2ecc71", width=160)
        self.btn_paste.pack(side="left", padx=10)
        
        self.btn_sample = ctk.CTkButton(top_bar, text="💡 Sample", command=self.show_sample, fg_color="#3498db", width=120)
        self.btn_sample.pack(side="right", padx=10)

        style = ttk.Style()
        style.configure("Treeview", rowheight=30, font=("Segoe UI", 10))
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))
        
        self.tree_frame = ctk.CTkFrame(self)
        self.tree_frame.pack(pady=10, fill="both", expand=True, padx=20)
        
        # 열 제목 간소화: x1.5, x2.0, x2.5
        self.tree = ttk.Treeview(self.tree_frame, 
                                columns=("Date", "Range", "Break", "NetDiff", "x1.5", "x2.0", "x2.5", "Weighted"), 
                                show='headings')
        
        cols = [
            ("Date", "날짜(요일)", 130), ("Range", "근무시간", 160), ("Break", "휴게", 80), 
            ("NetDiff", "실근무 (+/-)", 110), ("x1.5", "x1.5 (h)", 100), 
            ("x2.0", "x2.0 (h)", 100), ("x2.5", "x2.5 (h)", 100), ("Weighted", "환산합계", 100)
        ]
        
        for cid, txt, w in cols:
            self.tree.heading(cid, text=txt)
            self.tree.column(cid, width=w, anchor="center")
        
        self.tree.pack(side="left", fill="both", expand=True)
        
        self.summary_box = ctk.CTkTextbox(self, height=180, font=("Segoe UI", 15))
        self.summary_box.pack(pady=15, fill="x", padx=20)

    def show_sample(self):
        sample_path = resource_path("sample.png")
        if not os.path.exists(sample_path):
            messagebox.showwarning("Notice", "sample.png가 없습니다.")
            return
        top = ctk.CTkToplevel(self)
        img = Image.open(sample_path)
        img_tk = ImageTk.PhotoImage(img)
        label = tk.Label(top, image=img_tk); label.image = img_tk; label.pack()

    def load_image(self):
        f = filedialog.askopenfilename(); 
        if f: self.process_image(Image.open(f))

    def paste_from_clipboard(self):
        img = ImageGrab.grabclipboard()
        if isinstance(img, Image.Image): self.process_image(img)

    def process_image(self, img):
        try:
            w, h = img.size
            img = img.resize((w*2, h*2), Image.Resampling.LANCZOS)
            gray = ImageOps.grayscale(img)
            enhancer = ImageEnhance.Contrast(gray).enhance(1.8)
            full_text = pytesseract.image_to_string(enhancer, lang='kor+eng', config='--psm 6')
            self.calculate_data(full_text)
        except Exception as e:
            messagebox.showerror("Error", f"이미지 분석 실패: {e}")

    def calculate_data(self, text):
        # 정규식: 날짜, 시간 범위, 휴게시간(2~3자리)
        pattern = re.compile(r'(\d{1,2}/\d{1,2}).*?(\d{2}:\d{2})\s*-\s*(\d{2}:\d{2}).*?(\d{2,3})', re.S | re.I)
        
        for item in self.tree.get_children(): self.tree.delete(item)
        
        year = int(self.year_var.get())
        records = []
        lines = text.split('\n')
        
        for line in lines:
            match = pattern.search(line)
            if not match: continue
            
            try:
                d_v, s_t, e_t, r_raw = match.groups()
                r_val = int(r_raw)
                # OCR 60m -> 602m 오인식 보정 로직 유지
                if r_val > 240 and (r_raw.endswith('2') or r_raw.endswith('7')):
                    r_val = int(r_raw[:-1])
                
                dt = datetime.strptime(f"{year}/{d_v}", "%Y/%m/%d")
                is_holiday = dt.weekday() >= 5 or dt.strftime('%Y-%m-%d') in kr_holidays
                
                fmt = "%H:%M"
                start_time = datetime.strptime(s_t, fmt)
                end_time = datetime.strptime(e_t, fmt)
                if end_time < start_time: end_time += timedelta(days=1)
                
                records.append({
                    'dt': dt, 'start': start_time, 'end': end_time, 
                    'break_min': r_val, 'is_h': is_holiday, 'range': f"{s_t}-{e_t}"
                })
            except: continue

        records.sort(key=lambda x: x['dt'])
        
        total_net_sum = 0
        total_weighted_sum = 0
        holiday_dates = []

        for r in records:
            # 1분 단위로 정밀 계산 (야간 및 8시간 초과 중첩 계산을 위함)
            current = r['start']
            total_minutes = int((r['end'] - r['start']).total_seconds() / 60) - r['break_min']
            net_h = total_minutes / 60
            
            # 배율별 시간 합산 변수
            h15, h20, h25 = 0, 0, 0
            base_h = 0
            
            # 실근무 시간을 1분씩 순회하며 해당 시점의 배율 판정
            worked_min = 0
            temp_time = r['start']
            
            # 휴게시간을 제외한 실제 근무 시점별로 가산금 계산
            # (단순화를 위해 휴게시간은 근무 초반에 소진한 것으로 가정하거나 전체 비율로 계산 가능하나, 
            # 여기서는 실무적 편의를 위해 누적 근무시간 기준으로 판정)
            for m in range(int((r['end'] - r['start']).total_seconds() / 60)):
                # 현재 시각 (시:분)
                check_time = (temp_time + timedelta(minutes=m))
                hour_now = check_time.hour
                
                # 야간 근로 시간대 판정 (22:00 ~ 06:00)
                is_night = (hour_now >= 22 or hour_now < 6)
                
                # 휴게시간 제외 로직 (전체 시간 중 휴게시간 비율만큼 제외하고 계산)
                # 정밀도를 위해 휴게시간은 계산에서 제외
                if m < r['break_min']: continue
                
                worked_min += 1
                is_over_8h = (worked_min > 480) # 8시간(480분) 초과 여부
                
                # 근로기준법 배율 판정 매트릭스
                mult = 1.0
                if not r['is_h']: # 평일
                    if is_over_8h and is_night: mult = 2.0
                    elif is_over_8h or is_night: mult = 1.5
                    else: mult = 1.0
                else: # 휴일
                    if is_over_8h and is_night: mult = 2.5
                    elif is_over_8h: mult = 2.0
                    elif is_night: mult = 2.0 # 휴일 야간은 기본 1.5 + 야간 0.5
                    else: mult = 1.5
                
                # 배율별 시간 누적
                if mult == 1.0: base_h += 1/60
                elif mult == 1.5: h15 += 1/60
                elif mult == 2.0: h20 += 1/60
                elif mult == 2.5: h25 += 1/60

            weighted_day = (base_h * 1.0) + (h15 * 1.5) + (h20 * 2.0) + (h25 * 2.5)
            
            weekday_name = ["월", "화", "수", "목", "금", "토", "일"][r['dt'].weekday()]
            date_str = f"{r['dt'].strftime('%m/%d')} ({weekday_name})"
            if r['is_h']: holiday_dates.append(date_str)
            
            diff = net_h - 8
            diff_display = f"{net_h:.1f} ({'+' if diff>=0 else ''}{diff:.1f})"
            
            self.tree.insert("", "end", values=(
                date_str, r['range'], f"{r['break_min']}m", diff_display,
                f"{h15:.1f}" if h15>0 else "-", f"{h20:.1f}" if h20>0 else "-", 
                f"{h25:.1f}" if h25>0 else "-", f"{weighted_day:.1f}h"
            ))
            
            total_net_sum += net_h
            total_weighted_sum += weighted_day

        # 요약 및 Stand-by 알림 (날짜 포함)
        self.summary_box.delete("0.0", "end")
        summary = f"■ 총 실근무 합계: {total_net_sum:.1f} 시간\n"
        summary += f"■ 총 환산 OT 합계 (법정 가산 반영): {total_weighted_sum:.1f} 시간\n"
        summary += f"■ 정산: {total_net_sum - (len(records)*8):+.1f} 시간 (일별 상쇄 반영)\n"
        
        if holiday_dates:
            summary += f"\n⚠️ [Stand-by 근무여부 확인 필요]\n대상 일자: {', '.join(holiday_dates)}"
        
        self.summary_box.insert("0.0", summary)

if __name__ == "__main__":
    app = OTCalculator()
    app.mainloop()
