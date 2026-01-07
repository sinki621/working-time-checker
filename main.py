import os
import sys
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk
from PIL import Image, ImageGrab, ImageTk, ImageOps
import pytesseract
from datetime import datetime, timedelta
import ctypes
from tkinter.font import Font

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
    kr_holidays = holidays.country_holidays('KR')
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
        
        self.title("OT calculator (Producer: KI.Shin)")
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

        # 동적 레이아웃 설정 (폰트 크기 기반)
        tree_font = Font(family="Segoe UI", size=11)
        row_h = int(tree_font.metrics('linespace') * 2.5)

        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview", rowheight=row_h, font=tree_font, background="#ffffff", fieldbackground="#ffffff")
        style.configure("Treeview.Heading", font=("Segoe UI", 11, "bold"))
        
        self.tree_frame = ctk.CTkFrame(self)
        self.tree_frame.pack(pady=10, fill="both", expand=True, padx=20)
        
        scrollbar = ttk.Scrollbar(self.tree_frame)
        scrollbar.pack(side="right", fill="y")

        self.tree = ttk.Treeview(self.tree_frame, 
                                columns=("Date", "Range", "NetTime", "Break", "x1.5", "x2.0", "x2.5", "Weighted"), 
                                show='headings', yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.tree.yview)

        cols = [
            ("Date", "날짜(요일)", 130), ("Range", "근무범위", 180), 
            ("NetTime", "실근무(총시간)", 130), ("Break", "휴게(역산)", 100), 
            ("x1.5", "x1.5", 90), ("x2.0", "x2.0", 90), ("x2.5", "x2.5", 90), ("Weighted", "환산합계", 100)
        ]
        for cid, txt, w in cols:
            self.tree.heading(cid, text=txt)
            self.tree.column(cid, width=w, anchor="center", stretch=True)
        self.tree.pack(side="left", fill="both", expand=True)
        
        self.summary_box = ctk.CTkTextbox(self, height=260, font=("Segoe UI", 15))
        self.summary_box.pack(pady=15, fill="x", padx=20)

    def show_sample(self):
        sample_path = resource_path("sample.png")
        if not os.path.exists(sample_path):
            messagebox.showinfo("Notice", "sample.png 파일이 없습니다.")
            return
        top = ctk.CTkToplevel(self)
        img = Image.open(sample_path)
        img_tk = ImageTk.PhotoImage(img)
        label = tk.Label(top, image=img_tk); label.image = img_tk; label.pack()

    def load_image(self):
        f = filedialog.askopenfilename()
        if f: self.process_image(Image.open(f))

    def paste_from_clipboard(self):
        img = ImageGrab.grabclipboard()
        if isinstance(img, Image.Image): self.process_image(img)

    def process_image(self, img):
        try:
            img = ImageOps.grayscale(img)
            img = ImageOps.expand(img, border=50, fill='white')
            # 텍스트 구조 파악을 위해 psm 6 사용
            custom_config = r'--oem 1 --psm 6'
            full_text = pytesseract.image_to_string(img, lang='kor+eng', config=custom_config)
            self.calculate_data(full_text)
        except Exception as e:
            messagebox.showerror("Error", f"이미지 분석 실패: {e}")

    def calculate_data(self, text):
        # 정규표현식 재설계
        # 1. 날짜 및 시간 범위 (예: 12/31 (수) 06:50 - 03:40)
        line_pattern = re.compile(r'(\d{1,2}/\d{1,2}).*?(\d{2}:\d{2})\s*-\s*(\d{2}:\d{2})')
        # 2. 총 시간 (예: 18시간 50분)
        total_time_pattern = re.compile(r'(\d{1,2})\s*시간(?:\s*(\d{1,2})\s*분)?')
        
        for item in self.tree.get_children(): self.tree.delete(item)
        
        year = int(self.year_var.get())
        records = []
        
        for line in text.split('\n'):
            match = line_pattern.search(line)
            if not match: continue
            
            try:
                d_v, s_t, e_t = match.groups()
                # 출근/퇴근 객체 생성
                st_obj = datetime.strptime(s_t, "%H:%M")
                et_obj = datetime.strptime(e_t, "%H:%M")
                if et_obj < st_obj: et_obj += timedelta(days=1)
                
                # 출퇴근 총 소요 분(Minute)
                range_min = int((et_obj - st_obj).total_seconds() / 60)
                
                # '총 시간' 값 추출 (이 줄에서 가장 마지막에 나타나는 'X시간 Y분' 패턴 사용)
                time_found = list(total_time_pattern.finditer(line))
                if time_found:
                    last_m = time_found[-1]
                    h_val = int(last_m.group(1))
                    m_val = int(last_m.group(2)) if last_m.group(2) else 0
                    actual_worked_min = (h_val * 60) + m_val
                else:
                    actual_worked_min = range_min - 60 # 기본값
                
                # [핵심 로직] 휴게시간은 무조건 (범위 - 실제인식된총시간)으로 정의
                calculated_break = range_min - actual_worked_min
                if calculated_break < 0: calculated_break = 0

                dt = datetime.strptime(f"{year}/{d_v}", "%Y/%m/%d")
                is_h = dt.weekday() >= 5 or dt.strftime('%Y-%m-%d') in kr_holidays
                
                records.append({
                    'dt': dt, 'st': st_obj, 'et': et_obj, 
                    'brk': calculated_break, 'net_min': actual_worked_min, 
                    'is_h': is_h, 'range': f"{s_t}-{e_t}"
                })
            except: continue

        # 날짜순 정렬
        records.sort(key=lambda x: x['dt'])
        
        total_real_h, sum15, sum20, sum25, total_minus = 0, 0, 0, 0, 0
        holiday_list = []

        for r in records:
            h10, h15, h20, h25 = 0, 0, 0, 0
            # 1분 단위 가중치 계산 루프
            # 근무 시작 시점부터 '이미지에서 역산된 휴게시간'을 뺀 나머지만 루프를 돌림
            # 이렇게 하면 최종 net_h는 무조건 r['net_min']과 일치하게 됨
            dur = int((r['et'] - r['st']).total_seconds() / 60)
            
            worked_min_counter = 0
            for m in range(dur):
                # 출근 직후부터 휴게시간만큼은 계산 제외
                if m < r['brk']: continue
                
                check_time = r['st'] + timedelta(minutes=m)
                is_night = (check_time.hour >= 22 or check_time.hour < 6)
                worked_min_counter += 1
                is_over8 = (worked_min_counter > 480) # 8시간(480분) 초과
                
                weight = 1.0
                if not r['is_h']: # 평일
                    if is_over8 and is_night: weight = 2.0
                    elif is_over8 or is_night: weight = 1.5
                else: # 휴일
                    if is_over8 and is_night: weight = 2.5
                    elif is_over8 or is_night: weight = 2.0
                
                # 가중치별 시간 누적
                if weight == 1.0: h10 += 1/60
                elif weight == 1.5: h15 += 1/60
                elif weight == 2.0: h20 += 1/60
                elif weight == 2.5: h25 += 1/60

            # 이번 행의 최종 실근무 시간 (소수점 오차 방지를 위해 인식된 분값 사용)
            row_net_h = r['net_min'] / 60
            total_real_h += row_net_h
            sum15 += h15; sum20 += h20; sum25 += h25
            
            # 부족분 계산 (평일 8시간 미달 시)
            if not r['is_h'] and row_net_h < 8:
                total_minus += (8 - row_net_h)
            
            day_weighted_sum = (h10 * 1.0) + (h15 * 1.5) + (h20 * 2.0) + (h25 * 2.5)
            w_name = ["월", "화", "수", "목", "금", "토", "일"][r['dt'].weekday()]
            d_str = f"{r['dt'].strftime('%m/%d')} ({w_name})"
            if r['is_h']: holiday_list.append(d_str)
            
            self.tree.insert("", "end", values=(
                d_str, r['range'], 
                f"{int(r['net_min']//60)}h {int(r['net_min']%60)}m", 
                f"{r['brk']}m",
                f"{h15:.1f}", f"{h20:.1f}", f"{h25:.1f}", f"{day_weighted_sum:.1f}h"
            ))

        # 최종 합산 및 유연근무 상쇄
        adj_x15 = max(0, sum15 - total_minus)
        final_ot_total = (adj_x15 * 1.5) + (sum20 * 2.0) + (sum25 * 2.5)

        self.summary_box.delete("0.0", "end")
        msg = f"1. 총 실근무 합계: {total_real_h:.1f} 시간\n"
        msg += "-"*60 + "\n2. 배율별 OT 합계 (유연근무 상쇄 적용):\n"
        msg += f"   - [x1.5]: {adj_x15:.1f} h (부족분 {total_minus:.1f}h 차감됨)\n"
        msg += f"   - [x2.0]: {sum20:.1f} h\n   - [x2.5]: {sum25:.1f} h\n"
        msg += "-"*60 + "\n3. 최종 환산 OT 합계 (가중치 결과): {0:.1f} 시간\n".format(final_ot_total)
        if holiday_list: 
            msg += "\n⚠️ [Stand-by 근무여부 확인 필요]\n대상 일자: " + ", ".join(holiday_list)
        self.summary_box.insert("0.0", msg)

if __name__ == "__main__":
    app = OTCalculator()
    app.mainloop()
