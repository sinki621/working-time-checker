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
            # 이미지 전처리
            img_gray = img.convert('L')  # 흑백 변환
            
            # 선명도 향상
            from PIL import ImageFilter
            img_sharp = img_gray.filter(ImageFilter.SHARPEN)
            
            # 대비 강화
            img_contrast = ImageEnhance.Contrast(img_sharp).enhance(2.8)
            
            # 이진화
            img_binary = img_contrast.point(lambda x: 0 if x < 145 else 255)
            
            # 1단계: 전체 텍스트 읽기 (언어 감지 및 구조 파악)
            full_text = pytesseract.image_to_string(
                img_binary, 
                lang='kor+eng',
                config='--psm 6 --oem 3'
            )
            
            print(f"=== 1단계: 전체 텍스트 읽기 ===\n{full_text[:500]}...\n")
            
            # 2단계: 숫자만 정확히 추출
            digit_text = pytesseract.image_to_string(
                img_binary, 
                lang='eng',
                config='--psm 6 --oem 3 -c tessedit_char_whitelist=0123456789/:- '
            )
            
            print(f"=== 2단계: 숫자 추출 ===\n{digit_text[:500]}...\n")
            
            # 두 결과를 함께 처리
            self.process_ot_data(full_text, digit_text)
            
        except Exception as e:
            import traceback
            traceback.print_exc()
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
            max_width = 1000
            max_height = 700
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

    def process_ot_data(self, full_text, digit_text):
        """OCR로 추출한 텍스트를 파싱하여 초과근무 시간 계산
        
        Args:
            full_text: 전체 텍스트 (한글+영어)
            digit_text: 숫자만 추출한 텍스트
        """
        
        # 전체 텍스트에서 휴게시간 관련 정보 추출 (한글/영어 구분)
        is_korean = '분' in full_text or '시간' in full_text
        
        print(f"=== 언어 감지: {'한글' if is_korean else '영문'} ===\n")
        
        # 숫자 텍스트를 라인별로 분리
        digit_lines = [line.strip() for line in digit_text.strip().split('\n') if line.strip()]
        full_lines = [line.strip() for line in full_text.strip().split('\n') if line.strip()]
        
        # 각 라인 매칭
        matches = []
        
        for i, digit_line in enumerate(digit_lines):
            # 해당 라인의 전체 텍스트 찾기 (휴게시간 단위 확인용)
            full_line = full_lines[i] if i < len(full_lines) else ""
            
            # 1. 날짜 추출: MM/DD
            date_match = re.search(r'(\d{1,2}/\d{1,2})', digit_line)
            if not date_match:
                continue
            date_val = date_match.group(1)
            
            # 2. 시간 추출: HH:MM-HH:MM 또는 HH:MM - HH:MM
            time_match = re.search(r'(\d{2}:\d{2})\s*-?\s*(\d{2}:\d{2})', digit_line)
            if not time_match:
                continue
            start_time = time_match.group(1)
            end_time = time_match.group(2)
            
            # 3. 휴게시간 추출 (시간 이후의 숫자)
            rest_minutes = None
            
            # 시간 패턴 이후의 텍스트
            after_time = digit_line[time_match.end():]
            
            # 모든 2~3자리 숫자 찾기
            rest_candidates = re.findall(r'\b(\d{2,3})\b', after_time)
            
            # 전체 텍스트에서 "분", "본", "min" 등의 단위 찾기
            if is_korean:
                # 한글: "N분" 또는 "N본" 패턴
                rest_pattern = re.search(r'(\d{2,3})\s*[분본]', full_line)
            else:
                # 영문: "Nmin" 또는 "N min" 패턴
                rest_pattern = re.search(r'(\d{2,3})\s*min', full_line, re.I)
            
            if rest_pattern:
                rest_minutes = rest_pattern.group(1)
            elif rest_candidates:
                # 패턴 매칭 실패 시 첫 번째 숫자 사용
                rest_minutes = rest_candidates[0]
            else:
                rest_minutes = "60"  # 기본값
            
            # 휴게시간 검증 (15~180분)
            try:
                rest_int = int(rest_minutes)
                # 일반적인 휴게시간: 15, 30, 45, 60, 90, 120분
                if rest_int < 15 or rest_int > 180:
                    rest_minutes = "60"
                # 30분 단위가 아니면 반올림
                elif rest_int % 15 != 0:
                    rest_minutes = str(round(rest_int / 15) * 15)
            except:
                rest_minutes = "60"
            
            matches.append((date_val, start_time, end_time, rest_minutes))
            print(f"✓ 추출: {date_val} {start_time}-{end_time} 휴게 {rest_minutes}분")
        
        print(f"\n=== 총 {len(matches)}개 데이터 추출 ===\n")

        # 기존 테이블 데이터 삭제
        for item in self.tree.get_children():
            self.tree.delete(item)
        for item in self.summary_tree.get_children():
            self.summary_tree.delete(item)
        
        if len(matches) == 0:
            messagebox.showwarning(
                "No Data Found",
                "근무 데이터를 찾을 수 없습니다.\n\n"
                "확인사항:\n"
                "• 날짜가 MM/DD 형식인지 확인\n"
                "• 시간이 HH:MM-HH:MM 형식인지 확인\n"
                "• 휴게시간이 표시되어 있는지 확인\n"
                "• 이미지가 선명한지 확인\n\n"
                "콘솔 창에서 OCR 결과를 확인하세요."
            )
            return
        
        # 변수 초기화
        selected_year = int(self.year_var.get())
        processed_count = 0
        
        # 일별 데이터 저장
        daily_data = []
        
        #
