import os
import sys
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
import customtkinter as ctk
from PIL import Image, ImageGrab, ImageTk, ImageOps
import pytesseract
from datetime import datetime, timedelta
import ctypes

# =============================================================================
# 1. 환경 설정
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
        
        self.raw_records = []
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

        self.tree_frame = ctk.CTkFrame(self)
        self.tree_frame.pack(pady=10, fill="both", expand=True, padx=20)
        
        self.tree = ttk.Treeview(self.tree_frame, 
                                columns=("Date", "Range", "Break", "NetDiff", "x1.5", "x2.0", "x2.5", "Weighted"), 
                                show='headings')
        
        cols = [
            ("Date", "날짜(요일)", 130), ("Range", "근무시간", 160), ("Break", "휴게(분)", 80), 
            ("NetDiff", "실근무 (+/-)", 110), ("x1.5", "x1.5", 100), 
            ("x2.0", "x2.0", 100), ("x2.5", "x2.5", 100), ("Weighted", "환산합계", 100)
        ]
        
        for cid, txt, w in cols:
            self.tree.heading(cid, text=txt)
            self.tree.column(cid, width=w, anchor="center")
        
        self.tree.pack(side="left", fill="both", expand=True)
        self.tree.bind("<Double-1>", self.on_double_click)
        
        self.summary_box = ctk.CTkTextbox(self, height=260, font=("Segoe UI", 15))
        self.summary_box.pack(pady=15, fill="x", padx=20)

    def on_double_click(self, event):
        item_id = self.tree.identify_row(event.y)
        column = self.tree.identify_column(event.x)
        if item_id and column == "#3":
            old_val = self.tree.item(item_id, "values")[2].replace("m", "")
            new_val = simpledialog.askinteger("Edit", "정확한 휴게시간(분):", initialvalue=int(old_val))
            if new_val is not None:
                idx = self.tree.index(item_id)
                self.raw_records[idx]['brk'] = new_val
                self.refresh_table()

    def show_sample(self):
        sample_path = resource_path("sample.png")
        if os.path.exists(sample_path):
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
            # [이미지 처리 완전 제거] 
            # 단, Tesseract가 글자 가장자리를 인식할 공간(Padding)은 반드시 필요합니다.
            img = ImageOps.grayscale(img)
            img = ImageOps.expand(img, border=80, fill='white')
            
            # [핵심 수정: 엔진 튜닝]
            # --oem 1: LSTM 기반의 신경망 엔진 사용 (문맥 파악 우수)
            # --psm 6: 단일 텍스트 블록으로 간주 (표 읽기에 최적)
            # -c preserve_interword_spaces=1: 글자 사이의 미세한 간격 보존
            # -c tessedit_do_invert=0: 이미지 반전 시도 안 함 (원본 품질 유지)
            custom_config = (
                r'--oem 1 --psm 6 '
                r'-c tessedit_char_whitelist=0123456789/:- '
                r'-c preserve_interword_spaces=1 '
                r'-c tessedit_do_invert=0'
            )
            
            text = pytesseract.image_to_string(img, lang='eng', config=custom_config)
            self.parse_and_fill(text)
        except Exception as e:
            messagebox.showerror("Error", f"인식 실패: {e}")

    def parse_and_fill(self, text):
        line_pattern = re.compile(r'(\d{1,2}/\d{1,2}).*?(\d{2}:\d{2})\s*-\s*(\d{2}:\d{2})')
        num_pattern = re.compile(r'\d+')
        
        self.raw_records = []
        year = int(self.year_var.get())
        
        for line in text.split('\n'):
            match = line_pattern.search(line)
            if not match: continue
            
            try:
                d_v, s_t, e_t = match.groups()
                after_text = line[match.end():]
                nums = num_pattern.findall(after_text)
                
                # [보정 없이 원본 추출]
                break_val = 60
                if nums:
                    # 추출된 모든 숫자를 결합 (예: '12', '0' -> '120')
                    break_val = int("".join(nums))
                
                dt = datetime.strptime(f"{year}/{d_v}", "%Y/%m/%d")
                st, et = datetime.strptime(s_t, "%H:%M"), datetime.strptime(e_t, "%H:%M")
                if et < st: et += timedelta(days=1)
                is_h = dt.weekday() >= 5 or dt.strftime('%Y-%m-%d') in kr_holidays
                
                self.raw_records.append({'dt': dt, 'st': st, 'et': et, 'brk': break_val, 'is_h': is_h, 'range': f"{s_t}-{e_t}"})
            except: continue
        
        self.raw_records.sort(key=lambda x: x['dt'])
        self.refresh_table()

    def refresh_table(self):
        for item in self.tree.get_children(): self.tree.delete(item)
        t_net, s15, s20, s25, t_minus = 0, 0, 0, 0, 0
        h_list = []

        for r in self.raw_records:
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
            t_net += net; s15 += h15; s20 += h20; s25 += h25
            if not r['is_h'] and net < 8: t_minus += (8 - net)
            
            day_w = (h10 * 1.0) + (h15 * 1.5) + (h20 * 2.0) + (h25 * 2.5)
            d_str = f"{r['dt'].strftime('%m/%d')} ({(['월','화','수','목','금','토','일'][r['dt'].weekday()])})"
            if r['is_h']: h_list.append(d_str)
            
            diff = net - 8
            self.tree.insert("", "end", values=(d_str, r['range'], f"{r['brk']}m", f"{net:.1f} ({'+' if diff>=0 else ''}{diff:.1f})",
                                                f"{h15:.1f}", f"{h20:.1f}", f"{h25:.1f}", f"{day_w:.1f}h"))

        adj15 = max(0, s15 - t_minus)
        total = (adj15 * 1.5) + (s20 * 2.0) + (s25 * 2.5)
        self.summary_box.delete("0.0", "end")
        res = f"1. 실근무 합계: {t_net:.1f} h\n" + "-"*50
        res += f"\n2. 가중치별 OT:\n   x1.5: {adj15:.1f} h (부족분 {t_minus:.1f}h 차감)\n   x2.0: {s20:.1f} h\n   x2.5: {s25:.1f} h\n"
        res += "-"*50 + f"\n3. 최종 환산 OT: {total:.1f} 시간"
        if h_list: res += "\n\n⚠️ 휴일/특근 대상: " + ", ".join(h_list)
        self.summary_box.insert("0.0", res)

if __name__ == "__main__":
    app = OTCalculator()
    app.mainloop()
