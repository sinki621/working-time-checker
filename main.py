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

        # ---------------------------------------------------------
        # [동적 간격 설정] 폰트 크기에 따른 행 높이 자동 계산
        # ---------------------------------------------------------
        base_font_size = 11
        tree_font = Font(family="Segoe UI", size=base_font_size)
        # 폰트 높이의 약 2.5배를 행 높이로 설정 (겹침 방지)
        calculated_row_height = int(tree_font.metrics('linespace') * 2.5)

        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview", 
                        rowheight=calculated_row_height, 
                        font=tree_font,
                        background="#ffffff",
                        fieldbackground="#ffffff")
        style.configure("Treeview.Heading", font=("Segoe UI", 11, "bold"))
        
        self.tree_frame = ctk.CTkFrame(self)
        self.tree_frame.pack(pady=10, fill="both", expand=True, padx=20)
        
        # 스크롤바 추가 (동적 조절 시 필수)
        scrollbar = ttk.Scrollbar(self.tree_frame)
        scrollbar.pack(side="right", fill="y")

        self.tree = ttk.Treeview(self.tree_frame, 
                                columns=("Date", "Range", "Break", "NetDiff", "x1.5", "x2.0", "x2.5", "Weighted"), 
                                show='headings',
                                yscrollcommand=scrollbar.set)
        
        scrollbar.config(command=self.tree.yview)

        # ---------------------------------------------------------
        # [열 너비 자동 조절] stretch 옵션을 True로 설정하여 창 크기에 비례하게 함
        # ---------------------------------------------------------
        cols = [
            ("Date", "날짜(요일)", 1), ("Range", "근무시간", 2), ("Break", "휴게(역산)", 1), 
            ("NetDiff", "실근무 (+/-)", 1), ("x1.5", "x1.5", 1), 
            ("x2.0", "x2.0", 1), ("x2.5", "x2.5", 1), ("Weighted", "환산합계", 1)
        ]
        
        for cid, txt, weight in cols:
            self.tree.heading(cid, text=txt)
            # weight 값을 주어 창이 커질 때 해당 비율만큼 열 너비가 늘어남
            self.tree.column(cid, anchor="center", stretch=True, width=100)
        
        self.tree.pack(side="left", fill="both", expand=True)
        
        self.summary_box = ctk.CTkTextbox(self, height=260, font=("Segoe UI", 15))
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
        f = filedialog.askopenfilename()
        if f: self.process_image(Image.open(f))

    def paste_from_clipboard(self):
        img = ImageGrab.grabclipboard()
        if isinstance(img, Image.Image): self.process_image(img)

    def process_image(self, img):
        try:
            img = ImageOps.grayscale(img)
            img = ImageOps.expand(img, border=50, fill='white')
            
            # kor+eng 모델 사용
            custom_config = r'--oem 1 --psm 6'
            full_text = pytesseract.image_to_string(img, lang='kor+eng', config=custom_config)
            self.calculate_data(full_text)
        except Exception as e:
            messagebox.showerror("Error", f"이미지 분석 실패: {e}")

    def calculate_data(self, text):
        line_pattern = re.compile(r'(\d{1,2}/\d{1,2}).*?(\d{2}:\d{2})\s*-\s*(\d{2}:\d{2})')
        total_time_pattern = re.compile(r'(\d{1,2})시간\s*(\d{1,2})?분?')
        
        for item in self.tree.get_children(): self.tree.delete(item)
        
        year = int(self.year_var.get())
        records = []
        
        for line in text.split('\n'):
            match = line_pattern.search(line)
            if not match: continue
            
            try:
                d_v, s_t, e_t = match.groups()
                fmt = "%H:%M"
                st, et = datetime.strptime(s_t, fmt), datetime.strptime(e_t, fmt)
                if et < st: et += timedelta(days=1)
                range_minutes = int((et - st).total_seconds() / 60)
                
                after_text = line[match.end():]
                total_match = total_time_pattern.search(after_text)
                
                if total_match:
                    h_val = int(total_match.group(1))
                    m_val = int(total_match.group(2)) if total_match.group(2) else 0
                    actual_worked_minutes = (h_val * 60) + m_val
                    break_val = range_minutes - actual_worked_minutes
                    if break_val < 0: break_val = 60
                else:
                    break_val = 60
                
                dt = datetime.strptime(f"{year}/{d_v}", "%Y/%m/%d")
                is_h = dt.weekday() >= 5 or dt.strftime('%Y-%m-%d') in kr_holidays
                
                records.append({'dt': dt, 'st': st, 'et': et, 'brk': break_val, 'is_h': is_h, 'range': f"{s_t}-{e_t}"})
            except: continue

        records.sort(key=lambda x: x['dt'])
        total_net, sum15, sum20, sum25, total_minus = 0, 0, 0, 0, 0
        holiday_list = []

        for r in records:
            h10, h15, h20, h25, worked_min = 0, 0, 0, 0, 0
            dur = int((r['et'] - r['st']).total_seconds() / 60)
            for m in range(dur):
                if m < r['brk']: continue
                check = r['st'] + timedelta(minutes=m)
                is_n = (check.hour >= 22 or check.hour < 6)
                worked_min += 1
                ov8 = (worked_min > 480)
                
                m_val = 1.0
                if not r['is_h']:
                    if ov8 and is_n: m_val = 2.0
                    elif ov8 or is_n: m_val = 1.5
                else:
                    if ov8 and is_n: m_val = 2.5
                    elif ov8 or is_n: m_val = 2.0 
                
                if m_val == 1.0: h10 += 1/60
                elif m_val == 1.5: h15 += 1/60
                elif m_val == 2.0: h20 += 1/60
                elif m_val == 2.5: h25 += 1/60

            net = h10 + h15 + h20 + h25
            total_net += net
            sum15 += h15; sum20 += h20; sum25 += h25
            if not r['is_h'] and net < 8: total_minus += (8 - net)
            
            day_w = (h10 * 1.0) + (h15 * 1.5) + (h20 * 2.0) + (h25 * 2.5)
            w_name = ["월", "화", "수", "목", "금", "토", "일"][r['dt'].weekday()]
            d_str = f"{r['dt'].strftime('%m/%d')} ({w_name})"
            if r['is_h']: holiday_list.append(d_str)
            
            diff = net - 8
            self.tree.insert("", "end", values=(d_str, r['range'], f"{r['brk']}m", f"{net:.1f} ({'+' if diff>=0 else ''}{diff:.1f})",
                                                f"{h15:.1f}", f"{h20:.1f}", f"{h25:.1f}", f"{day_w:.1f}h"))

        adj_x15 = max(0, sum15 - total_minus)
        final_ot = (adj_x15 * 1.5) + (sum20 * 2.0) + (sum25 * 2.5)

        self.summary_box.delete("0.0", "end")
        msg = f"1. 총 실근무 합계: {total_net:.1f} 시간\n"
        msg += "-"*60 + "\n2. 배율별 OT 합계 (유연근무 상쇄 적용):\n"
        msg += f"   - [x1.5]: {adj_x15:.1f} h (부족분 {total_minus:.1f}h 차감됨)\n"
        msg += f"   - [x2.0]: {sum20:.1f} h\n   - [x2.5]: {sum25:.1f} h\n"
        msg += "-"*60 + "\n3. 최종 환산 OT 합계 (가중치 결과): {0:.1f} 시간\n".format(final_ot)
        if holiday_list: msg += "\n⚠️ [Stand-by 근무여부 확인 필요]\n대상 일자: " + ", ".join(holiday_list)
        self.summary_box.insert("0.0", msg)

if __name__ == "__main__":
    app = OTCalculator()
    app.mainloop()
