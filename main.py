import os
import sys
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk
from PIL import Image, ImageGrab, ImageTk, ImageOps
from rapidocr_onnxruntime import RapidOCR
import numpy as np
from datetime import datetime, timedelta
import ctypes
from tkinter.font import Font

# =============================================================================
# 환경 설정 및 리소스 경로 처리
# =============================================================================
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
    """ PyInstaller 임시 폴더에서 파일을 찾는 함수 """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# =============================================================================
# 메인 애플리케이션 클래스
# =============================================================================
class OTCalculator(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.title("OT calculator (Producer: KI.Shin)")
        self.geometry("1600x950")
        ctk.set_appearance_mode("light")
        
        # RapidOCR 초기화 (config.yaml 경로 문제 방지)
        try:
            self.engine = RapidOCR()
        except Exception as e:
            messagebox.showerror("OCR Error", f"RapidOCR 초기화 실패: {e}")

        self.setup_ui()
        
        # 단축키 바인딩
        self.bind('<Control-v>', lambda e: self.paste_from_clipboard())
        self.bind('<Control-V>', lambda e: self.paste_from_clipboard())

    def setup_ui(self):
        # 상단바
        top_bar = ctk.CTkFrame(self, fg_color="transparent")
        top_bar.pack(pady=15, fill="x", padx=20)
        
        self.year_var = ctk.StringVar(value=str(datetime.now().year))
        ctk.CTkLabel(top_bar, text="Year:", font=("Segoe UI", 14, "bold")).pack(side="left", padx=5)
        self.year_dropdown = ctk.CTkComboBox(top_bar, values=["2024", "2025", "2026", "2027"], variable=self.year_var, width=90)
        self.year_dropdown.pack(side="left", padx=5)
        
        self.btn_load = ctk.CTkButton(top_bar, text="📁 Load File", command=self.load_image, width=140)
        self.btn_load.pack(side="left", padx=10)
        
        self.btn_paste = ctk.CTkButton(top_bar, text="📋 Paste (Ctrl+V)", command=self.paste_from_clipboard, fg_color="#2ecc71", width=160)
        self.btn_paste.pack(side="left", padx=10)
        
        # 오류가 났던 부분: 존재하지 않는 함수를 호출하는 버튼을 제거하거나 함수를 정의해야 함
        # 여기서는 도움말 텍스트로 대체합니다.
        ctk.CTkLabel(top_bar, text="* 실근무 칸 더블클릭 수정 가능", font=("Segoe UI", 12, "italic"), text_color="gray").pack(side="right", padx=20)

        # 트리뷰 설정
        tree_font = Font(family="Segoe UI", size=11)
        row_h = int(tree_font.metrics('linespace') * 2.5)
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview", rowheight=row_h, font=tree_font, background="#ffffff", fieldbackground="#ffffff")
        style.configure("Treeview.Heading", font=("Segoe UI", 11, "bold"))
        
        self.tree_frame = ctk.CTkFrame(self)
        self.tree_frame.pack(pady=10, fill="both", expand=True, padx=20)
        
        self.tree = ttk.Treeview(self.tree_frame, columns=("Date", "Range", "NetTime", "Break", "x1.5", "x2.0", "x2.5", "Weighted"), show='headings')
        cols = [("Date", "날짜(요일)", 130), ("Range", "근무범위", 180), ("NetTime", "실근무(총시간)", 150), 
                ("Break", "휴게(역산)", 100), ("x1.5", "x1.5", 90), ("x2.0", "x2.0", 90), ("x2.5", "x2.5", 90), ("Weighted", "환산합계", 100)]
        for cid, txt, w in cols:
            self.tree.heading(cid, text=txt)
            self.tree.column(cid, width=w, anchor="center")
        self.tree.pack(side="left", fill="both", expand=True)
        self.tree.bind("<Double-1>", self.on_double_click)

        self.summary_box = ctk.CTkTextbox(self, height=200, font=("Segoe UI", 15))
        self.summary_box.pack(pady=15, fill="x", padx=20)

    def on_double_click(self, event):
        item = self.tree.identify_row(event.y)
        column = self.tree.identify_column(event.x)
        if column == "#3":
            x, y, w, h = self.tree.bbox(item, column)
            entry = tk.Entry(self.tree)
            entry.insert(0, self.tree.item(item, 'values')[2])
            entry.place(x=x, y=y, width=w, height=h)
            entry.focus_set()
            def save(e):
                self.tree.set(item, column=column, value=entry.get()); self.recalculate_from_table(); entry.destroy()
            entry.bind("<Return>", save); entry.bind("<FocusOut>", lambda e: entry.destroy())

    def load_image(self):
        f = filedialog.askopenfilename()
        if f: self.process_image(Image.open(f))

    def paste_from_clipboard(self):
        img = ImageGrab.grabclipboard()
        if isinstance(img, Image.Image): self.process_image(img)

    def process_image(self, img):
        try:
            img_np = np.array(img.convert('RGB'))
            result, _ = self.engine(img_np)
            if result:
                texts = [line[1] for line in result]
                self.calculate_data(texts)
        except Exception as e:
            messagebox.showerror("Error", f"분석 오류: {e}")

    def calculate_data(self, results):
        for item in self.tree.get_children(): self.tree.delete(item)
        year = int(self.year_var.get())
        for i, text in enumerate(results):
            text = text.replace(" ", "")
            date_match = re.search(r'(\d{1,2}/\d{1,2})', text)
            if date_match:
                d_val = date_match.group(1)
                f_range, f_net = "", 0
                for j in range(i+1, min(i+10, len(results))):
                    t = results[j].replace(" ", "")
                    if ":" in t and "-" in t: f_range = t
                    if "시간" in t:
                        h = re.search(r'(\d+)시간', t); m = re.search(r'(\d+)분', t)
                        f_net = (int(h.group(1)) if h else 0)*60 + (int(m.group(1)) if m else 0)
                    if j < len(results)-1 and re.search(r'\d{1,2}/\d{1,2}', results[j+1]): break
                if f_range and f_net > 0:
                    try:
                        times = re.findall(r'\d{2}:\d{2}', f_range)
                        st_s, et_s = times[0], times[1]
                        st = datetime.strptime(st_s, "%H:%M"); et = datetime.strptime(et_s, "%H:%M")
                        if et < st: et += timedelta(days=1)
                        range_min = int((et-st).total_seconds()/60); brk = range_min - f_net
                        dt = datetime.strptime(f"{year}/{d_val}", "%Y/%m/%d")
                        is_h = dt.weekday() >= 5 or dt.strftime('%Y-%m-%d') in kr_holidays
                        self.insert_row(dt, st_s, et_s, f_net, brk, is_h)
                    except: pass
        self.recalculate_from_table()

    def insert_row(self, dt, s_t, e_t, net_min, brk, is_h):
        w_name = ["월", "화", "수", "목", "금", "토", "일"][dt.weekday()]
        d_str = f"{dt.strftime('%m/%d')} ({w_name})"
        self.tree.insert("", "end", values=(d_str, f"{s_t}-{e_t}", f"{int(net_min//60)}h {int(net_min%60)}m", f"{int(brk)}m", "", "", "", ""))

    def recalculate_from_table(self):
        total_net, sum15, sum20, sum25, total_minus = 0, 0, 0, 0, 0
        year = int(self.year_var.get())
        for item in self.tree.get_children():
            v = self.tree.item(item, 'values')
            dt = datetime.strptime(f"{year}/{v[0].split(' ')[0]}", "%Y/%m/%d")
            st_s, et_s = v[1].split('-')
            st = datetime.strptime(st_s, "%H:%M"); et = datetime.strptime(et_s, "%H:%M")
            if et < st: et += timedelta(days=1)
            h_m = re.search(r'(\d+)h', v[2]); m_m = re.search(r'(\d+)m', v[2])
            net_min = (int(h_m.group(1))*60 if h_m else 0) + (int(m_m.group(1)) if m_m else 0)
            range_min = int((et-st).total_seconds()/60); brk = range_min - net_min
            h10, h15, h20, h25, w_cnt = 0, 0, 0, 0, 0
            is_h = dt.weekday() >= 5 or dt.strftime('%Y-%m-%d') in kr_holidays
            for m in range(range_min):
                if m < brk: continue
                c = st + timedelta(minutes=m); is_n = (c.hour >= 22 or c.hour < 6)
                w_cnt += 1; ov8 = (w_cnt > 480); w = 1.0
                if not is_h:
                    if ov8 and is_n: w = 2.0
                    elif ov8 or is_n: w = 1.5
                else:
                    if ov8 and is_n: w = 2.5
                    elif ov8 or is_n: w = 2.0
                if w == 1.0: h10 += 1/60
                elif w == 1.5: h15 += 1/60
                elif w == 2.0: h20 += 1/60
                elif w == 2.5: h25 += 1/60
            row_net = net_min/60; total_net += row_net; sum15 += h15; sum20 += h20; sum25 += h25
            if not is_h and row_net < 8: total_minus += (8 - row_net)
            w_sum = (h10*1 + h15*1.5 + h20*2 + h25*2.5)
            self.tree.item(item, values=(v[0], v[1], f"{int(net_min//60)}h {int(net_min%60)}m", f"{int(brk)}m", f"{h15:.1f}", f"{h20:.1f}", f"{h25:.1f}", f"{w_sum:.1f}h"))
        adj_x15 = max(0, sum15 - total_minus); f_ot = (adj_x15 * 1.5) + (sum20 * 2.0) + (sum25 * 2.5)
        self.summary_box.delete("0.0", "end")
        self.summary_box.insert("0.0", f"1. 총 실근무 합계: {total_net:.1f}h\n2. OT 합계: x1.5({adj_x15:.1f}h), x2.0({sum20:.1f}h), x2.5({sum25:.1f}h)\n3. 최종 환산 합계: {f_ot:.1f} 시간")

if __name__ == "__main__":
    app = OTCalculator()
    app.mainloop()
