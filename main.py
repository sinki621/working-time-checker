import os
import sys
import re
import cv2
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk
from PIL import Image, ImageGrab, ImageOps
from rapidocr_onnxruntime import RapidOCR
import numpy as np
from datetime import datetime, timedelta
import ctypes

# DPI 설정
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except:
    pass

try:
    import holidays
    kr_holidays = holidays.country_holidays('KR')
except:
    kr_holidays = {}

class OTCalculator(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.title("OT calculator (Producer: KI.Shin)")
        self.geometry("1600x950")
        ctk.set_appearance_mode("light")
        
        try:
            self.engine = RapidOCR()
        except Exception as e:
            messagebox.showerror("OCR Error", f"엔진 초기화 실패: {e}")

        self.setup_ui()
        self.bind('<Control-v>', lambda e: self.paste_from_clipboard())
        self.bind('<Control-V>', lambda e: self.paste_from_clipboard())

    def setup_ui(self):
        top_bar = ctk.CTkFrame(self, fg_color="transparent")
        top_bar.pack(pady=15, fill="x", padx=20)
        
        self.year_var = ctk.StringVar(value="2024")
        ctk.CTkLabel(top_bar, text="Year:", font=("Segoe UI", 14, "bold")).pack(side="left", padx=5)
        ctk.CTkComboBox(top_bar, values=["2024", "2025", "2026", "2027"], variable=self.year_var, width=90).pack(side="left", padx=5)
        
        ctk.CTkButton(top_bar, text="📁 Load File", command=self.load_image, width=140).pack(side="left", padx=10)
        ctk.CTkButton(top_bar, text="📋 Paste (Ctrl+V)", command=self.paste_from_clipboard, fg_color="#2ecc71", width=160).pack(side="left", padx=10)
        
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview", rowheight=35, font=("Segoe UI", 11))
        
        self.tree_frame = ctk.CTkFrame(self)
        self.tree_frame.pack(pady=10, fill="both", expand=True, padx=20)
        
        self.tree = ttk.Treeview(self.tree_frame, columns=("Date", "Range", "NetTime", "Break", "x1.5", "x2.0", "x2.5", "Weighted"), show='headings')
        cols = [("Date", "날짜", 130), ("Range", "시간범위", 180), ("NetTime", "실근무(총시간)", 150), ("Break", "휴게", 100), ("x1.5", "x1.5", 80), ("x2.0", "x2.0", 80), ("x2.5", "x2.5", 80), ("Weighted", "환산합계", 100)]
        for cid, txt, w in cols:
            self.tree.heading(cid, text=txt)
            self.tree.column(cid, width=w, anchor="center")
        self.tree.pack(side="left", fill="both", expand=True)

        self.summary_box = ctk.CTkTextbox(self, height=180, font=("Segoe UI", 15))
        self.summary_box.pack(pady=15, fill="x", padx=20)

    # [핵심] 이미지 전처리 함수 추가
    def preprocess_image(self, pil_img):
        # 1. PIL 이미지를 OpenCV 포맷으로 변환
        img = cv2.cvtColor(np.array(pil_img), cv2.COLOR_RGB2BGR)
        
        # 2. 그레이스케일 변환
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        
        # 3. 대비(Contrast) 강화
        # CLAHE (Contrast Limited Adaptive Histogram Equalization) 적용
        clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8,8))
        contrast = clahe.apply(gray)
        
        # 4. 이진화 (Binarization) - 글자를 더 뚜렷하게
        # 배경은 흰색, 글자는 검은색으로 최적화
        _, thresh = cv2.threshold(contrast, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        
        return thresh

    def load_image(self):
        f = filedialog.askopenfilename()
        if f: self.process_image(Image.open(f))

    def paste_from_clipboard(self):
        img = ImageGrab.grabclipboard()
        if isinstance(img, Image.Image): self.process_image(img)

    def process_image(self, img):
        try:
            # ✅ OCR 실행 전 이미지 전처리 적용
            processed_img = self.preprocess_image(img)
            
            # 전처리된 이미지로 OCR 엔진 실행
            result, _ = self.engine(processed_img)
            if not result: return

            result.sort(key=lambda x: x[0][0][1])
            lines_data = []
            if result:
                last_y = result[0][0][0][1]
                current_line_elements = []
                for res in result:
                    if abs(res[0][0][1] - last_y) < 25:
                        current_line_elements.append(res)
                    else:
                        current_line_elements.sort(key=lambda x: x[0][0][0])
                        lines_data.append([el[1] for el in current_line_elements])
                        current_line_elements = [res]
                        last_y = res[0][0][1]
                current_line_elements.sort(key=lambda x: x[0][0][0])
                lines_data.append([el[1] for el in current_line_elements])

            self.parse_rows(lines_data)
        except Exception as e:
            messagebox.showerror("Error", f"분석 오류: {e}")

    def parse_rows(self, lines_data):
        for item in self.tree.get_children(): self.tree.delete(item)
        year = int(self.year_var.get())
        
        for elements in lines_data:
            line_full = "".join(elements).replace(" ", "")
            date_m = re.search(r'(\d{1,2}/\d{1,2})', line_full)
            if not date_m: continue
            
            times = re.findall(r'\d{2}:\d{2}', line_full)
            if len(times) < 2: continue
            
            f_net = 0
            for elem in reversed(elements):
                elem_c = elem.replace(" ", "")
                # 시간/분 인식 패턴 강화
                h_m = re.search(r'(\d+)(?:시간|시|h|H)', elem_c)
                m_m = re.search(r'(\d+)(?:분|m|M)', elem_c)
                
                if h_m or m_m:
                    f_net = (int(h_m.group(1)) if h_m else 0) * 60 + (int(m_m.group(1)) if m_m else 0)
                    if f_net > 0: break
            
            if f_net == 0: # 텍스트 인식 실패 시 마지막 숫자 패턴 사용
                all_nums = re.findall(r'\d+', line_full)
                if len(all_nums) >= 8:
                    try: f_net = int(all_nums[-2]) * 60 + int(all_nums[-1])
                    except: pass

            if f_net > 0:
                try:
                    st_s, et_s = times[0], times[1]
                    st = datetime.strptime(st_s, "%H:%M")
                    et = datetime.strptime(et_s, "%H:%M")
                    if et < st: et += timedelta(days=1)
                    range_min = int((et-st).total_seconds()/60)
                    brk = range_min - f_net
                    dt = datetime.strptime(f"{year}/{date_m.group(1)}", "%Y/%m/%d")
                    self.insert_row(dt, st_s, et_s, f_net, brk)
                except: pass
        self.recalculate_from_table()

    def insert_row(self, dt, s_t, e_t, net_min, brk):
        w_name = ["월", "화", "수", "목", "금", "토", "일"][dt.weekday()]
        d_str = f"{dt.strftime('%m/%d')} ({w_name})"
        self.tree.insert("", "end", values=(d_str, f"{s_t}-{e_t}", f"{int(net_min//60)}h {int(net_min%60)}m", f"{max(0, int(brk))}m", "", "", "", ""))

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
            self.tree.item(item, values=(v[0], v[1], f"{int(net_min//60)}h {int(net_min%60)}m", f"{max(0, int(brk))}m", f"{h15:.1f}", f"{h20:.1f}", f"{h25:.1f}", f"{w_sum:.1f}h"))
        adj_x15 = max(0, sum15 - total_minus); f_ot = (adj_x15 * 1.5) + (sum20 * 2.0) + (sum25 * 2.5)
        self.summary_box.delete("0.0", "end")
        self.summary_box.insert("0.0", f"1. 총 실근무: {total_net:.1f}h\n2. OT: x1.5({adj_x15:.1f}h), x2.0({sum20:.1f}h), x2.5({sum25:.1f}h)\n3. 합계: {f_ot:.1f}h")

if __name__ == "__main__":
    OTCalculator().mainloop()
