import os
import sys
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk
from PIL import Image, ImageGrab, ImageTk, ImageOps
import easyocr
import numpy as np
from datetime import datetime, timedelta
import ctypes
from tkinter.font import Font

# =============================================================================
# 1. 환경 설정 및 리소스 경로 처리
# =============================================================================
try:
    # 윈도우 고해상도(DPI) 인식 설정
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
    """ .exe 내부 리소스 또는 외부 파일 경로를 가져오는 함수 """
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
        
        # [중요] .exe 내부에 포함된 models 폴더를 사용하도록 설정
        model_path = resource_path("models")
        try:
            # GPU가 없는 환경이 많으므로 gpu=False 설정
            self.reader = easyocr.Reader(['ko', 'en'], gpu=False, model_storage_directory=model_path)
        except Exception as e:
            messagebox.showerror("OCR Error", f"모델을 로드할 수 없습니다: {e}")

        self.setup_ui()
        
        # 붙여넣기 단축키 바인딩
        self.bind('<Control-v>', lambda e: self.paste_from_clipboard())
        self.bind('<Control-V>', lambda e: self.paste_from_clipboard())

    def setup_ui(self):
        # 상단 컨트롤바
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
        
        ctk.CTkLabel(top_bar, text="* 인식 오류 시 '실근무' 열을 더블클릭하여 수정 가능", 
                     font=("Segoe UI", 12, "italic"), text_color="gray").pack(side="right", padx=20)

        # 트리뷰 (표) 설정
        tree_font = Font(family="Segoe UI", size=11)
        row_h = int(tree_font.metrics('linespace') * 2.5)

        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview", rowheight=row_h, font=tree_font, background="#ffffff", fieldbackground="#ffffff")
        style.configure("Treeview.Heading", font=("Segoe UI", 11, "bold"))
        
        self.tree_frame = ctk.CTkFrame(self)
        self.tree_frame.pack(pady=10, fill="both", expand=True, padx=20)
        
        self.tree = ttk.Treeview(self.tree_frame, 
                                columns=("Date", "Range", "NetTime", "Break", "x1.5", "x2.0", "x2.5", "Weighted"), 
                                show='headings')

        cols = [
            ("Date", "날짜(요일)", 130), ("Range", "근무범위", 180), 
            ("NetTime", "실근무(총시간)", 150), ("Break", "휴게(역산)", 100), 
            ("x1.5", "x1.5", 90), ("x2.0", "x2.0", 90), ("x2.5", "x2.5", 90), ("Weighted", "환산합계", 100)
        ]
        for cid, txt, w in cols:
            self.tree.heading(cid, text=txt)
            self.tree.column(cid, width=w, anchor="center")
        
        self.tree.pack(side="left", fill="both", expand=True)
        self.tree.bind("<Double-1>", self.on_double_click)

        # 요약 정보창
        self.summary_box = ctk.CTkTextbox(self, height=200, font=("Segoe UI", 15))
        self.summary_box.pack(pady=15, fill="x", padx=20)

    def on_double_click(self, event):
        """ 실근무 시간 수동 수정 기능 """
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell": return
        column = self.tree.identify_column(event.x)
        if column != "#3": return
        
        item = self.tree.identify_row(event.y)
        x, y, w, h = self.tree.bbox(item, column)
        
        entry = tk.Entry(self.tree)
        entry.insert(0, self.tree.item(item, 'values')[2])
        entry.place(x=x, y=y, width=w, height=h)
        entry.focus_set()
        
        def save_edit(event):
            self.tree.set(item, column=column, value=entry.get())
            self.recalculate_from_table()
            entry.destroy()
        
        entry.bind("<Return>", save_edit)
        entry.bind("<FocusOut>", lambda e: entry.destroy())

    def load_image(self):
        f = filedialog.askopenfilename()
        if f: self.process_image(Image.open(f))

    def paste_from_clipboard(self):
        img = ImageGrab.grabclipboard()
        if isinstance(img, Image.Image): self.process_image(img)

    def process_image(self, img):
        try:
            # EasyOCR용 Numpy 변환 및 전처리
            img_np = np.array(img.convert('RGB'))
            results = self.reader.readtext(img_np, detail=0)
            self.calculate_with_logic(results)
        except Exception as e:
            messagebox.showerror("Error", f"이미지 분석 실패: {e}")

    def calculate_with_logic(self, results):
        for item in self.tree.get_children(): self.tree.delete(item)
        year = int(self.year_var.get())
        
        i = 0
        while i < len(results):
            text = results[i].replace(" ", "")
            # 날짜 패턴 찾기 (12/31 등)
            date_match = re.search(r'(\d{1,2}/\d{1,2})', text)
            if date_match:
                date_val = date_match.group(1)
                found_range, found_net = "", 0
                
                # 날짜 이후 텍스트에서 범위와 총 시간을 탐색
                for j in range(i+1, min(i+10, len(results))):
                    t = results[j].replace(" ", "")
                    # 1. 근무 범위 (06:50-03:40)
                    if ":" in t and "-" in t:
                        found_range = t
                    # 2. 총 시간 (18시간50분)
                    if "시간" in t:
                        h = re.search(r'(\d+)시간', t)
                        m = re.search(r'(\d+)분', t)
                        found_net = (int(h.group(1)) if h else 0) * 60 + (int(m.group(1)) if m else 0)
                    
                    # 다음 날짜를 만나면 탐색 중단
                    if j < len(results)-1 and re.search(r'\d{1,2}/\d{1,2}', results[j+1]):
                        break
                
                if found_range and found_net > 0:
                    try:
                        times = re.findall(r'\d{2}:\d{2}', found_range)
                        st_s, et_s = times[0], times[1]
                        st = datetime.strptime(st_s, "%H:%M")
                        et = datetime.strptime(et_s, "%H:%M")
                        if et < st: et += timedelta(days=1)
                        
                        range_min = int((et - st).total_seconds() / 60)
                        brk = range_min - found_net # 휴게시간 역산
                        
                        dt = datetime.strptime(f"{year}/{date_val}", "%Y/%m/%d")
                        is_h = dt.weekday() >= 5 or dt.strftime('%Y-%m-%d') in kr_holidays
                        
                        self.insert_row(dt, st_s, et_s, found_net, brk, is_h)
                        i = j
                    except: pass
            i += 1
        self.recalculate_from_table()

    def insert_row(self, dt, s_t, e_t, net_min, brk, is_h):
        w_name = ["월", "화", "수", "목", "금", "토", "일"][dt.weekday()]
        d_str = f"{dt.strftime('%m/%d')} ({w_name})"
        net_str = f"{int(net_min//60)}h {int(net_min%60)}m"
        self.tree.insert("", "end", values=(d_str, f"{s_t}-{e_t}", net_str, f"{int(brk)}m", "", "", "", ""))

    def recalculate_from_table(self):
        total_net, sum15, sum20, sum25, total_minus = 0, 0, 0, 0, 0
        year = int(self.year_var.get())
        
        for item in self.tree.get_children():
            v = list(self.tree.item(item, 'values'))
            dt = datetime.strptime(f"{year}/{v[0].split(' ')[0]}", "%Y/%m/%d")
            st_s, et_s = v[1].split('-')
            st = datetime.strptime(st_s, "%H:%M"); et = datetime.strptime(et_s, "%H:%M")
            if et < st: et += timedelta(days=1)
            
            h_m = re.search(r'(\d+)h', v[2]); m_m = re.search(r'(\d+)m', v[2])
            net_min = (int(h_m.group(1)) if h_m else 0)*60 + (int(m_m.group(1)) if m_m else 0)
            range_min = int((et-st).total_seconds()/60)
            brk = range_min - net_min
            
            h10, h15, h20, h25, w_cnt = 0, 0, 0, 0, 0
            is_h = dt.weekday() >= 5 or dt.strftime('%Y-%m-%d') in kr_holidays
            
            # 가중치 계산 (1분 단위 정밀 루프)
            for m in range(range_min):
                if m < brk: continue
                curr = st + timedelta(minutes=m)
                is_n = (curr.hour >= 22 or curr.hour < 6)
                w_cnt += 1; ov8 = (w_cnt > 480)
                
                w = 1.0
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

            row_net_h = net_min/60
            total_net += row_net_h
            sum15 += h15; sum20 += h20; sum25 += h25
            if not is_h and row_net_h < 8: total_minus += (8 - row_net_h)
            
            w_sum = (h10*1 + h15*1.5 + h20*2 + h25*2.5)
            self.tree.item(item, values=(v[0], v[1], f"{net_min//60}h {net_min%60}m", f"{int(brk)}m", 
                                         f"{h15:.1f}", f"{h20:.1f}", f"{h25:.1f}", f"{w_sum:.1f}h"))

        adj_x15 = max(0, sum15 - total_minus)
        final_ot = (adj_x15 * 1.5) + (sum20 * 2.0) + (sum25 * 2.5)
        
        self.summary_box.delete("0.0", "end")
        msg = f"1. 총 실근무 합계: {total_net:.1f} 시간\n"
        msg += f"2. OT 배율별 합계 (유연 상쇄 적용):\n"
        msg += f"   - x1.5: {adj_x15:.1f}h (부족분 {total_minus:.1f}h 차감됨)\n"
        msg += f"   - x2.0: {sum20:.1f}h / x2.5: {sum25:.1f}h\n"
        msg += "-"*50 + f"\n3. 최종 환산 OT 합계: {final_ot:.1f} 시간"
        self.summary_box.insert("0.0", msg)

if __name__ == "__main__":
    app = OTCalculator()
    app.mainloop()
