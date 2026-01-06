import os
import sys
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk
from PIL import Image, ImageEnhance, ImageGrab
import pytesseract
from datetime import datetime, timedelta
import ctypes
import io

# =============================================================================
# 1. 윈도우 DPI 인식 강제 설정 (4K, 고해상도 모니터 대응)
# =============================================================================
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)  # PROCESS_PER_MONITOR_DPI_AWARE
except:
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)  # PROCESS_SYSTEM_DPI_AWARE
    except:
        try:
            ctypes.windll.user32.SetProcessDPIAware()  # 레거시 방식
        except:
            pass  # DPI 설정 실패해도 프로그램은 실행

# holidays 라이브러리 버전 호환성 처리
try:
    import holidays
    # 최신 버전 시도
    try:
        kr_holidays = holidays.country_holidays('KR')
    except:
        # 구버전 방식
        kr_holidays = holidays.KR()
except ImportError:
    # holidays 라이브러리가 없는 경우 빈 딕셔너리 사용
    kr_holidays = {}

# =============================================================================
# 2. PyInstaller 번들 리소스 경로 처리
# =============================================================================
def resource_path(relative_path):
    """PyInstaller로 빌드된 실행 파일에서 리소스 경로를 가져옴"""
    try:
        # PyInstaller가 생성한 임시 폴더 경로
        base_path = sys._MEIPASS
    except Exception:
        # 개발 환경에서는 현재 디렉토리 사용
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# =============================================================================
# 3. 메인 애플리케이션 클래스
# =============================================================================
class OTCalculator(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # CustomTkinter 스케일링 고정 (모든 환경에서 동일한 크기)
        ctk.set_widget_scaling(1.0)
        ctk.set_window_scaling(1.0)
        
        # 창 설정
        self.title("CSV Chart Viewer - OT Calculator (Producer: KI.Shin)")
        self.geometry("1400x850")
        self.minsize(1200, 700)
        
        # 테마 설정
        ctk.set_appearance_mode("light")
        
        # 한국 공휴일 데이터 로드 (전역에서 이미 초기화됨)
        self.kr_holidays = kr_holidays
        
        # Tesseract OCR 엔진 경로 설정
        self.setup_tesseract()
        
        # UI 구성
        self.setup_ui()
        
        # 클립보드 붙여넣기 단축키 바인딩 (Ctrl+V)
        self.bind('<Control-v>', self.paste_from_clipboard)
        self.bind('<Control-V>', self.paste_from_clipboard)

    def setup_tesseract(self):
        """Tesseract OCR 엔진 경로 설정"""
        try:
            # PyInstaller로 번들된 Tesseract 경로
            engine_root = resource_path("Tesseract-OCR")
            tesseract_exe = os.path.join(engine_root, "tesseract.exe")
            tessdata_dir = os.path.join(engine_root, "tessdata")
            
            # 경로 존재 확인
            if os.path.exists(tesseract_exe):
                pytesseract.pytesseract.tesseract_cmd = tesseract_exe
                os.environ["TESSDATA_PREFIX"] = tessdata_dir
                print(f"✓ Tesseract found at: {tesseract_exe}")
            else:
                print(f"⚠ Tesseract not found at: {tesseract_exe}")
                # 시스템에 설치된 Tesseract 사용 시도
                pytesseract.pytesseract.tesseract_cmd = "tesseract"
        except Exception as e:
            print(f"⚠ Tesseract setup warning: {e}")
            # 기본 경로 사용

    def setup_ui(self):
        """UI 구성 요소 생성"""
        
        # 고정 폰트 크기 정의
        BTN_FONT_SIZE = 16
        HEADER_FONT_SIZE = 11
        BODY_FONT_SIZE = 10
        ROW_HEIGHT = 28
        
        # =====================================================================
        # 상단: 연도 선택 + 파일 로드 버튼
        # =====================================================================
        top_frame = ctk.CTkFrame(self, fg_color="transparent")
        top_frame.pack(pady=15)
        
        # 연도 선택 레이블
        year_label = ctk.CTkLabel(
            top_frame,
            text="Year:",
            font=("Segoe UI", 14, "bold")
        )
        year_label.pack(side="left", padx=(0, 10))
        
        # 연도 선택 드롭다운
        self.year_var = ctk.StringVar(value="2025")
        self.year_dropdown = ctk.CTkComboBox(
            top_frame,
            values=["2024", "2025", "2026"],
            variable=self.year_var,
            font=("Segoe UI", 14),
            width=100,
            height=50,
            state="readonly"
        )
        self.year_dropdown.pack(side="left", padx=(0, 20))
        
        # 파일 로드 버튼
        self.btn_load = ctk.CTkButton(
            top_frame, 
            text="📁 Load File", 
            command=self.load_image, 
            font=("Segoe UI", BTN_FONT_SIZE, "bold"),
            width=180, 
            height=50,
            corner_radius=8
        )
        self.btn_load.pack(side="left", padx=(0, 10))
        
        # 클립보드 붙여넣기 버튼
        self.btn_paste = ctk.CTkButton(
            top_frame, 
            text="📋 Paste (Ctrl+V)", 
            command=lambda: self.paste_from_clipboard(None), 
            font=("Segoe UI", BTN_FONT_SIZE, "bold"),
            width=200, 
            height=50,
            corner_radius=8,
            fg_color="#2ecc71",
            hover_color="#27ae60"
        )
        self.btn_paste.pack(side="left", padx=(0, 10))
        
        # Sample 버튼
        self.btn_sample = ctk.CTkButton(
            top_frame, 
            text="📄 Sample", 
            command=self.show_sample, 
            font=("Segoe UI", BTN_FONT_SIZE, "bold"),
            width=150, 
            height=50,
            corner_radius=8,
            fg_color="#9b59b6",
            hover_color="#8e44ad"
        )
        self.btn_sample.pack(side="left")

        # =====================================================================
        # 중앙: 데이터 테이블 (Treeview)
        # =====================================================================
        
        # Treeview 스타일 설정
        style = ttk.Style()
        style.theme_use("clam")
        
        # 헤더 스타일
        style.configure(
            "Treeview.Heading", 
            font=("Segoe UI", HEADER_FONT_SIZE, "bold"),
            background="#E0E0E0",
            foreground="black",
            relief="flat"
        )
        
        # 본문 스타일
        style.configure(
            "Treeview", 
            font=("Segoe UI", BODY_FONT_SIZE),
            rowheight=ROW_HEIGHT,
            background="white",
            foreground="black",
            fieldbackground="white",
            borderwidth=1
        )
        
        # 선택된 행 스타일
        style.map('Treeview', 
                  background=[('selected', '#0078D7')],
                  foreground=[('selected', 'white')])

        # Treeview 프레임 (스크롤바 포함)
        tree_frame = ctk.CTkFrame(self, fg_color="white")
        tree_frame.pack(pady=10, fill="both", expand=True, padx=20)
        
        # 수직 스크롤바
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical")
        
        # Treeview 생성
        self.tree = ttk.Treeview(
            tree_frame,
            columns=("Date", "Range", "Rest", "Net", "Diff", "Type", "x1.5", "x2.0", "x2.5", "Total"),
            show='headings',
            yscrollcommand=scrollbar.set
        )
        
        scrollbar.config(command=self.tree.yview)
        scrollbar.pack(side="right", fill="y")
        
        # 컬럼 정의 (이름, 헤더 텍스트, 너비)
        columns = [
            ("Date", "날짜", 120),
            ("Range", "근무시간", 140),
            ("Rest", "휴게", 80),
            ("Net", "실근무", 90),
            ("Diff", "기준차이", 100),
            ("Type", "근무유형", 140),
            ("x1.5", "OT×1.5", 90),
            ("x2.0", "OT×2.0", 90),
            ("x2.5", "OT×2.5", 90),
            ("Total", "환산합계", 100)
        ]
        
        for col_id, header_text, width in columns:
            self.tree.heading(col_id, text=header_text)
            self.tree.column(col_id, width=width, anchor="center", minwidth=50)
        
        self.tree.pack(side="left", fill="both", expand=True)

        # =====================================================================
        # 하단: 합계 테이블
        # =====================================================================
        summary_frame = ctk.CTkFrame(self, fg_color="white", border_width=2)
        summary_frame.pack(pady=15, fill="x", padx=20)
        
        # 합계 테이블 생성
        self.summary_tree = ttk.Treeview(
            summary_frame,
            columns=("Label", "Net", "OT", "x1.5", "x2.0", "x2.5", "Total"),
            show='headings',
            height=3
        )
        
        # 합계 테이블 스타일
        style.configure("Summary.Treeview", rowheight=35)
        self.summary_tree.configure(style="Summary.Treeview")
        
        # 합계 컬럼 정의
        summary_columns = [
            ("Label", "구분", 140),
            ("Net", "실근무", 120),
            ("OT", "순수OT", 120),
            ("x1.5", "OT×1.5", 120),
            ("x2.0", "OT×2.0", 120),
            ("x2.5", "OT×2.5", 120),
            ("Total", "환산합계", 120)
        ]
        
        for col_id, header_text, width in summary_columns:
            self.summary_tree.heading(col_id, text=header_text)
            self.summary_tree.column(col_id, width=width, anchor="center")
        
        self.summary_tree.pack(fill="x", padx=10, pady=10)

    def load_image(self):
        """스크린샷 파일 선택 및 로드"""
        file_path = filedialog.askopenfilename(
            title="Select Shiftee Screenshot",
            filetypes=[
                ("Image files", "*.png *.jpg *.jpeg *.bmp *.gif"),
                ("All files", "*.*")
            ]
        )
        
        if not file_path:
            return
        
        try:
            # 버튼 상태 변경 (처리 중)
            self.btn_load.configure(text="Analyzing...", state="disabled")
            self.btn_paste.configure(state="disabled")
            self.update()
            
            # 이미지 로드 및 처리
            img = Image.open(file_path)
            self.process_image(img)
            
        except Exception as e:
            messagebox.showerror(
                "Error", 
                f"Failed to process image:\n\n{str(e)}\n\nPlease check:\n"
                "1. Tesseract OCR is properly installed\n"
                "2. Korean language data (kor.traineddata) exists\n"
                "3. Image file is not corrupted"
            )
        finally:
            # 버튼 상태 복원
            self.btn_load.configure(text="📁 Load File", state="normal")
            self.btn_paste.configure(state="normal")

    def paste_from_clipboard(self, event):
        """클립보드에서 이미지 붙여넣기"""
        try:
            # 버튼 상태 변경 (처리 중)
            self.btn_load.configure(state="disabled")
            self.btn_paste.configure(text="Analyzing...", state="disabled")
            self.update()
            
            # 클립보드에서 이미지 가져오기
            img = ImageGrab.grabclipboard()
            
            if img is None:
                messagebox.showwarning(
                    "No Image in Clipboard",
                    "클립보드에 이미지가 없습니다.\n\n"
                    "스크린샷을 복사한 후 다시 시도해주세요.\n"
                    "(Win + Shift + S 또는 Print Screen)"
                )
                return
            
            # PIL Image 객체인지 확인
            if not isinstance(img, Image.Image):
                # 파일 경로 리스트인 경우 (윈도우에서 파일 복사)
                if isinstance(img, list) and len(img) > 0:
                    img = Image.open(img[0])
                else:
                    messagebox.showwarning(
                        "Invalid Clipboard Content",
                        "클립보드 내용을 이미지로 변환할 수 없습니다."
                    )
                    return
            
            # 이미지 처리
            self.process_image(img)
            
        except Exception as e:
            messagebox.showerror(
                "Clipboard Error",
                f"클립보드에서 이미지를 가져오는데 실패했습니다:\n\n{str(e)}"
            )
        finally:
            # 버튼 상태 복원
            self.btn_load.configure(state="normal")
            self.btn_paste.configure(text="📋 Paste (Ctrl+V)", state="normal")

    def process_image(self, img):
        """이미지 전처리 및 OCR 실행"""
        try:
            # 이미지 전처리 (OCR 정확도 향상)
            img = img.convert('L')  # 흑백 변환
            img = ImageEnhance.Contrast(img).enhance(2.2)  # 대비 더 강화
            img = img.point(lambda x: 0 if x < 155 else 255)  # 이진화
            
            # OCR 실행 (한국어 + 영어)
            raw_text = pytesseract.image_to_string(
                img, 
                lang='kor+eng',
                config='--psm 6'  # PSM 6: 단일 텍스트 블록 (더 정확한 인식)
            )
            
            # 추출된 텍스트 처리
            self.process_ot_data(raw_text)
            
        except Exception as e:
            raise Exception(f"Image processing failed: {str(e)}")

    def show_sample(self):
        """예제 이미지 표시"""
        try:
            # sample.png 파일 경로 찾기
            sample_path = resource_path("sample.png")
            
            # 파일이 없으면 현재 디렉토리에서도 시도
            if not os.path.exists(sample_path):
                sample_path = "sample.png"
            
            if not os.path.exists(sample_path):
                messagebox.showwarning(
                    "Sample Not Found",
                    "sample.png 파일을 찾을 수 없습니다.\n\n"
                    "루트 폴더에 sample.png 파일이 있는지 확인해주세요."
                )
                return
            
            # 새 창 생성
            sample_window = tk.Toplevel(self)
            sample_window.title("스크린샷 예제")
            sample_window.geometry("1000x700")
            
            # 창을 화면 중앙에 배치
            sample_window.update_idletasks()
            x = (sample_window.winfo_screenwidth() // 2) - (1000 // 2)
            y = (sample_window.winfo_screenheight() // 2) - (700 // 2)
            sample_window.geometry(f"1000x700+{x}+{y}")
            
            # 안내 텍스트
            info_label = tk.Label(
                sample_window,
                text="📸 예제와 같이 스크린샷을 찍으세요",
                font=("Segoe UI", 18, "bold"),
                fg="#2c3e50",
                bg="white",
                pady=15
            )
            info_label.pack(fill="x")
            
            # 추가 설명
            detail_label = tk.Label(
                sample_window,
                text="• 날짜, 근무시간, 휴게시간이 모두 보이도록 캡처하세요\n"
                     "• 여러 날의 데이터를 한번에 캡처할 수 있습니다\n"
                     "• Win + Shift + S 로 화면 일부를 캡처한 후 Ctrl+V로 붙여넣기",
                font=("Segoe UI", 11),
                fg="#34495e",
                bg="white",
                justify="left",
                pady=10
            )
            detail_label.pack(fill="x")
            
            # 이미지 표시를 위한 프레임
            img_frame = tk.Frame(sample_window, bg="white")
            img_frame.pack(fill="both", expand=True, padx=20, pady=10)
            
            # 스크롤바 추가
            canvas = tk.Canvas(img_frame, bg="white")
            scrollbar_y = tk.Scrollbar(img_frame, orient="vertical", command=canvas.yview)
            scrollbar_x = tk.Scrollbar(img_frame, orient="horizontal", command=canvas.xview)
            
            # 이미지를 담을 프레임
            scrollable_frame = tk.Frame(canvas, bg="white")
            scrollable_frame.bind(
                "<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
            )
            
            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
            
            # 이미지 로드 및 표시
            sample_img = Image.open(sample_path)
            
            # 이미지 크기 조정 (너무 크면 축소)
            max_width = 950
            max_height = 550
            img_width, img_height = sample_img.size
            
            if img_width > max_width or img_height > max_height:
                ratio = min(max_width / img_width, max_height / img_height)
                new_width = int(img_width * ratio)
                new_height = int(img_height * ratio)
                sample_img = sample_img.resize((new_width, new_height), Image.Resampling.LANCZOS)
            
            # PIL Image를 Tkinter에서 표시 가능한 형태로 변환
            # RGB 모드로 변환 (RGBA나 다른 모드일 경우 대비)
            if sample_img.mode != 'RGB':
                sample_img = sample_img.convert('RGB')
            
            # PhotoImage 생성 (PIL의 ImageTk 사용)
            from PIL import ImageTk
            photo = ImageTk.PhotoImage(sample_img)
            
            img_label = tk.Label(scrollable_frame, image=photo, bg="white")
            img_label.image = photo  # 참조 유지 (가비지 컬렉션 방지)
            img_label.pack()
            
            # 스크롤바 배치
            scrollbar_y.pack(side="right", fill="y")
            scrollbar_x.pack(side="bottom", fill="x")
            canvas.pack(side="left", fill="both", expand=True)
            
            # 닫기 버튼
            close_btn = tk.Button(
                sample_window,
                text="닫기",
                font=("Segoe UI", 12, "bold"),
                bg="#3498db",
                fg="white",
                padx=30,
                pady=10,
                command=sample_window.destroy,
                relief="flat",
                cursor="hand2"
            )
            close_btn.pack(pady=15)
            
        except Exception as e:
            messagebox.showerror(
                "Error",
                f"예제 이미지를 표시할 수 없습니다:\n\n{str(e)}"
            )

    def process_ot_data(self, raw_text):
        """OCR로 추출한 텍스트를 파싱하여 초과근무 시간 계산"""
        
        # 정규식 패턴 개선: 휴게시간 인식 정확도 향상
        # 60분, 60 분, 60m, 60 m, 60min 등 다양한 형식 지원
        pattern = re.compile(
            r'(\d{1,2}/\d{1,2}).*?'  # 날짜 (예: 12/25)
            r'(\d{2}:\d{2})\s*[-~]\s*(\d{2}:\d{2}).*?'  # 시간 범위
            r'(\d+)\s*(?:분|m|min|M|MIN)',  # 휴게시간 (다양한 형식)
            re.S | re.I
        )
        matches = pattern.findall(raw_text)

        # 기존 테이블 데이터 삭제
        for item in self.tree.get_children():
            self.tree.delete(item)
        for item in self.summary_tree.get_children():
            self.summary_tree.delete(item)
        
        # 변수 초기화
        selected_year = int(self.year_var.get())
        processed_count = 0
        
        # 일별 데이터 저장
        daily_data = []
        
        # 각 매칭된 데이터 처리
        for match in matches:
            date_val, start_time, end_time, rest_minutes = match
            
            try:
                # 날짜 객체 생성 (선택된 연도 사용)
                month, day = map(int, date_val.split('/'))
                date_obj = datetime(selected_year, month, day)
                
                # 공휴일 및 주말 판단
                is_weekend = date_obj.weekday() >= 5  # 토요일(5), 일요일(6)
                is_public_holiday = date_obj in self.kr_holidays
                is_holiday = is_weekend or is_public_holiday
                
                holiday_name = self.kr_holidays.get(date_obj) if is_public_holiday else ""
                day_name = ["월", "화", "수", "목", "금", "토", "일"][date_obj.weekday()]

                # 시간 계산
                start = datetime.strptime(start_time, "%H:%M")
                end = datetime.strptime(end_time, "%H:%M")
                
                # 종료 시간이 시작 시간보다 이른 경우 (다음날 새벽)
                if end < start:
                    end += timedelta(days=1)
                
                # 실근무 시간 = 총 근무시간 - 휴게시간
                total_hours = (end - start).total_seconds() / 3600
                rest_hours = float(rest_minutes) / 60
                net_hours = total_hours - rest_hours
                
                # 기준시간 대비 차이 (8시간 기준)
                diff_hours = net_hours - 8.0
                
                # 일별 데이터 저장
                daily_data.append({
                    'date': date_obj,
                    'date_val': date_val,
                    'day_name': day_name,
                    'start_time': start_time,
                    'end_time': end_time,
                    'rest_minutes': rest_minutes,
                    'net_hours': net_hours,
                    'diff_hours': diff_hours,
                    'is_holiday': is_holiday,
                    'holiday_name': holiday_name
                })
                
                processed_count += 1
                
            except Exception as e:
                print(f"⚠ Failed to process row: {match} - {e}")
                continue
        
        # 데이터가 없는 경우
        if processed_count == 0:
            messagebox.showwarning(
                "No Data Found",
                "근무 데이터를 찾을 수 없습니다.\n\n"
                "확인사항:\n"
                "1. 날짜 형식: MM/DD\n"
                "2. 시간 형식: HH:MM-HH:MM\n"
                "3. 휴게시간: 숫자+분 (예: 60분)\n"
                "4. 이미지가 선명한지 확인"
            )
            return
        
        # 날짜순 정렬
        daily_data.sort(key=lambda x: x['date'])
        
        # 유연근무제 계산: 누적 차이 시간 추적
        cumulative_diff = 0
        
        # 합계 변수
        total_net = 0
        total_ot_15 = 0  # 1.5배 OT
        total_ot_20 = 0  # 2.0배 OT
        total_ot_25 = 0  # 2.5배 OT
        
        # 각 일자별 계산 및 표시
        for data in daily_data:
            net_hours = data['net_hours']
            diff_hours = data['diff_hours']
            is_holiday = data['is_holiday']
            
            # 누적 차이 업데이트
            cumulative_diff += diff_hours
            
            # 순수 OT 계산 (누적 기준)
            if cumulative_diff > 0:
                pure_ot = cumulative_diff
            else:
                pure_ot = 0
            
            # 배율별 OT 계산
            ot_15 = 0  # 평일 8시간 초과
            ot_20 = 0  # 휴일 8시간 초과
            ot_25 = 0  # 사용 안 함
            
            if is_holiday:
                # 휴일: 전체 근무시간에 1.5배 (8시간까지) + 2.0배 (초과분)
                type_str = f"휴일({data['holiday_name'] if data['holiday_name'] else data['day_name']})"
                if net_hours > 0:
                    if net_hours <= 8:
                        ot_15 = net_hours * 0.5  # 실제로는 1.5배이므로 0.5 추가
                    else:
                        ot_15 = 8 * 0.5
                        ot_20 = net_hours - 8
            else:
                # 평일: 8시간 초과분만 1.5배
                type_str = "평일"
                if diff_hours > 0:
                    ot_15 = diff_hours * 0.5
            
            # 환산 합계
            weighted_total = net_hours + ot_15 + ot_20 + ot_25
            
            # 합계 누적
            total_net += net_hours
            total_ot_15 += ot_15
            total_ot_20 += ot_20
            total_ot_25 += ot_25
            
            # 기준차이 표시 (+ 또는 -)
            if abs(diff_hours) < 0.1:
                diff_str = "-"
            elif diff_hours > 0:
                diff_str = f"+{diff_hours:.1f}h"
            else:
                diff_str = f"{diff_hours:.1f}h"
            
            # 테이블에 행 추가
            self.tree.insert("", "end", values=(
                f"{data['date_val']}({data['day_name']})",
                f"{data['start_time']}-{data['end_time']}",
                f"{data['rest_minutes']}분",
                f"{net_hours:.1f}h",
                diff_str,
                type_str,
                f"{ot_15:.1f}h" if ot_15 > 0 else "-",
                f"{ot_20:.1f}h" if ot_20 > 0 else "-",
                f"{ot_25:.1f}h" if ot_25 > 0 else "-",
                f"{weighted_total:.1f}h"
            ))
        
        # 순수 OT 계산 (40시간 기준 주간 또는 전체 누적)
        pure_ot_total = max(0, cumulative_diff)
        
        # 최종 환산 합계
        final_weighted = total_net + total_ot_15 + total_ot_20 + total_ot_25
        
        # 합계 테이블 업데이트
        self.summary_tree.insert("", "end", values=(
            "합계",
            f"{total_net:.1f}h",
            f"{pure_ot_total:.1f}h",
            f"{total_ot_15:.1f}h",
            f"{total_ot_20:.1f}h",
            f"{total_ot_25:.1f}h" if total_ot_25 > 0 else "-",
            f"{final_weighted:.1f}h"
        ), tags=('total',))
        
        # 합계 행 스타일 (굵게)
        self.summary_tree.tag_configure('total', font=("Segoe UI", 11, "bold"))

# =============================================================================
# 4. 프로그램 진입점
# =============================================================================
if __name__ == "__main__":
    app = OTCalculator()
    app.mainloop()
