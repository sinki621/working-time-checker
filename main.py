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

# 환경 설정
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
        # 상단 바
        top_bar = ctk.CTkFrame(self, fg_color="transparent")
        top_bar.pack(pady=15, fill="x", padx=20)
        
        self.year_var = ctk.StringVar(value=str(datetime.now().year))
        ctk.CTkLabel(top_bar, text="Year:", font=("Segoe UI", 14, "bold")).pack(side="left", padx=5)
        ctk.CTkComboBox(top_bar, values=["2024", "2025", "2026", "2027"], variable=self.year_var, width=90).pack(side="left", padx=5)
        
        ctk.CTkButton(top_bar, text="📁 Load File", command=self.load_image, width=140).pack(side="left", padx=10)
        ctk.CTkButton(top_bar, text="📋 Paste (Ctrl+V)", command=self.paste_from_clipboard, fg_color="#2ecc71", width=160).pack(side="left", padx=10)
        
        # 도움말 안내 추가
        ctk.CTkLabel(top_bar, text="* 인식 오류 시 '실근무' 열을 더블클릭하여 수정 가능", font=("Segoe UI", 12, "italic"), text_color="gray").pack(side="right", padx=20)

        # Treeview 설정 (동적 높이)
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
        
        # 수동 수정 기능 바인딩
        self.tree.bind("<Double-1>", self.on_double_click)

        self.summary_box = ctk.CTkTextbox(self, height=200, font=("Segoe UI", 15))
        self.summary_box.pack(pady=15, fill="x", padx=20)

    def on_double_click(self, event):
        """표의 값을 더블클릭하여 수동 수정 (OCR 오류 대비)"""
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell": return
        
        column = self.tree.identify_column(event.x)
        if column != "#3": return # 실근무 열만 수정 가능하도록 제한
        
        item = self.tree.identify_row(event.y)
        x, y, w, h = self.tree.bbox(item, column)
        
        entry = tk.Entry(self.tree)
        entry.insert(0, self.tree.item(item, 'values')[2])
        entry.place(x=x, y=y, width=w, height=h)
        entry.focus_set()
        
        def save_edit(event):
            val = entry.get()
            # 간단한 유효성 검사 (Xh Ym 형식)
            if 'h' in val:
                old_values = list(self.tree.item(item, 'values'))
                old_values[2] = val
                self.tree.item(item, values=old_values)
                self.recalculate_from_table() # 수정된 값으로 전체 다시 계산
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
            img = ImageOps.grayscale(img)
            img = ImageOps.expand(img, border=40, fill='white')
            full_text = pytesseract.image_to_string(img, lang='kor+eng', config='--oem 1 --psm 6')
            self.calculate_data(full_text)
        except Exception as e:
            messagebox.showerror("Error", f"분석 오류: {e}")

    def calculate_data(self, text):
        line_pattern = re.compile(r'(\d{1,2}/\d{1,2}).*?(\d{2}:\d{2})\s*-\s*(\d{2}:\d{2})')
        # '시간' 단어를 포함한 숫자만 정밀하게 타격
        total_time_pattern = re.compile(r'(\d{1,2})\s*시간(?:\s*(\d{1,2})\s*분)?')
        
        for item in self.tree.get_children(): self.tree.delete(item)
        
        year = int(self.year_var.get())
        
        for line in text.split('\n'):
            match = line_pattern.search(line)
            if not match: continue
            
            try:
                d_v, s_t, e_t = match.groups()
                st = datetime.strptime(s_t, "%H:%M")
                et = datetime.strptime(e_t, "%H:%M")
                if et < st: et += timedelta(days=1)
                range_min = int((et - st).total_seconds() / 60)
                
                # [강력한 인식 로직] 줄 전체에서 '시간' 패턴을 모두 찾아 가장 큰 값을 채택
                all_times = total_time_pattern.findall(line)
                if all_times:
                    # (시간, 분) 튜플 리스트에서 총 분이 가장 큰 것 선택
                    minutes_list = [(int(h)*60 + (int(m) if m else 0)) for h, m in all_times]
                    actual_worked_min = max(minutes_list)
                else:
                    actual_worked_min = range_min - 60
                
                brk = range_min - actual_worked_min
                if brk < 0: brk = 0

                dt = datetime.strptime(f"{year}/{d_v}", "%Y/%m/%d")
                is_h = dt.weekday() >= 5 or dt.strftime('%Y-%m-%d') in kr_holidays
                
                # 표에 삽입 (가중치 계산은 별도 함수에서 수행)
                self.insert_row(dt, s_t, e_t, actual_worked_min, brk, is_h)
            except: continue
        
        self.recalculate_from_table()

    def insert_row(self, dt, s_t, e_t, net_min, brk, is_h):
        w_name = ["월", "화", "수", "목", "금", "토", "일"][dt.weekday()]
        d_str = f"{dt.strftime('%m/%d')} ({w_name})"
        net_str = f"{net_min//60}h {net_min%60}m"
        self.tree.insert("", "end", values=(d_str, f"{s_t}-{e_t}", net_str, f"{brk}m", "", "", "", ""))

    def recalculate_from_table(self):
        """표의 데이터를 기반으로 가중치 및 합계 재계산 (수동 수정 대응)"""
        total_net, sum15, sum20, sum25, total_minus = 0, 0, 0, 0, 0
        holiday_list = []
        year = int(self.year_var.get())

        for item in self.tree.get_children():
            vals = list(self.tree.item(item, 'values'))
            date_str = vals[0].split(' ')[0]
            range_str = vals[1]
            net_str = vals[2]
            
            # 시간 파싱
            dt = datetime.strptime(f"{year}/{date_str}", "%Y/%m/%d")
            st_str, et_str = range_str.split('-')
            st = datetime.strptime(st_str, "%H:%M")
            et = datetime.strptime(et_str, "%H:%M")
            if et < st: et += timedelta(days=1)
            
            # 실근무 파싱 (예: 18h 50m)
            h_part = int(re.search(r'(\d+)h', net_str).group(1))
            m_part = int(re.search(r'(\d+)m', net_str).group(1)) if 'm' in net_str else 0
            net_min = h_part * 60 + m_part
            
            # 휴게시간 역산
            range_min = int((et - st).total_seconds() / 60)
            brk = range_min - net_min
            
            # 가중치 계산 루프
            h10, h15, h20, h25 = 0, 0, 0, 0
            is_h = dt.weekday() >= 5 or dt.strftime('%Y-%m-%d') in kr_holidays
            if is_h: holiday_list.append(vals[0])

            worked_cnt = 0
            for m in range(range_min):
                if m < brk: continue # 휴게 제외
                
                check = st + timedelta(minutes=m)
                is_n = (check.hour >= 22 or check.hour < 6)
                worked_cnt += 1
                ov8 = (worked_cnt > 480)
                
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

            row_net = net_min / 60
            total_net += row_net
            sum15 += h15; sum20 += h20; sum25 += h25
            if not is_h and row_net < 8: total_minus += (8 - row_net)
            
            weighted_total = (h10*1 + h15*1.5 + h20*2 + h25*2.5)
            
            # 업데이트
            self.tree.item(item, values=(vals[0], vals[1], net_str, f"{brk}m", f"{h15:.1f}", f"{h20:.1f}", f"{h25:.1f}", f"{weighted_total:.1f}h"))

        # 최종 요약
        adj_x15 = max(0, sum15 - total_minus)
        final_ot = (adj_x15 * 1.5) + (sum20 * 2.0) + (sum25 * 2.5)
        
        self.summary_box.delete("0.0", "end")
        msg = f"1. 총 실근무 합계: {total_net:.1f} 시간\n"
        msg += f"2. OT 합계 (유연 상쇄): x1.5({adj_x15:.1f}h), x2.0({sum20:.1f}h), x2.5({sum25:.1f}h)\n"
        msg += f"3. 최종 환산 OT: {final_ot:.1f} 시간"
        self.summary_box.insert("0.0", msg)

if __name__ == "__main__":
    OTCalculator().mainloop()
