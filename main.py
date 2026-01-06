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
        self.geometry("1100x800")
        self.minsize(900, 600)
        
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
        SUMMARY_FONT_SIZE = 20
        
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
        self.btn_paste.pack(side="left")

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
            columns=("Date", "Range", "Rest", "Net", "Type", "Total"),
            show='headings',
            yscrollcommand=scrollbar.set
        )
        
        scrollbar.config(command=self.tree.yview)
        scrollbar.pack(side="right", fill="y")
        
        # 컬럼 정의 (이름, 헤더 텍스트, 너비)
        columns = [
            ("Date", "날짜", 140),
            ("Range", "근무시간", 160),
            ("Rest", "휴게", 90),
            ("Net", "실근무", 110),
            ("Type", "근무유형", 200),
            ("Total", "환산합계", 120)
        ]
        
        for col_id, header_text, width in columns:
            self.tree.heading(col_id, text=header_text)
            self.tree.column(col_id, width=width, anchor="center", minwidth=50)
        
        self.tree.pack(side="left", fill="both", expand=True)

        # =====================================================================
        # 하단: 합계 표시 박스
        # =====================================================================
        self.summary_box = ctk.CTkTextbox(
            self, 
            height=100,
            font=("Segoe UI", SUMMARY_FONT_SIZE, "bold"),
            border_width=2,
            fg_color="white",
            corner_radius=8
        )
        self.summary_box.pack(pady=15, fill="x", padx=20)

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
            img = ImageEnhance.Contrast(img).enhance(2.0)  # 대비 증가
            img = img.point(lambda x: 0 if x < 160 else 255)  # 이진화
            
            # OCR 실행 (한국어 + 영어)
            raw_text = pytesseract.image_to_string(
                img, 
                lang='kor+eng',
                config='--psm 4'  # PSM 4: 단일 컬럼 텍스트 가정
            )
            
            # 추출된 텍스트 처리
            self.process_ot_data(raw_text)
            
        except Exception as e:
            raise Exception(f"Image processing failed: {str(e)}")

    def process_ot_data(self, raw_text):
        """OCR로 추출한 텍스트를 파싱하여 초과근무 시간 계산"""
        
        # 정규식 패턴: 날짜, 시간, 휴게시간 추출
        # 예: 12/25 09:00-18:00 60분 또는 12/25 09:00-18:00 60m
        pattern = re.compile(
            r'(\d{1,2}/\d{1,2}).*?'  # 날짜 (예: 12/25)
            r'(\d{2}:\d{2})\s*-\s*(\d{2}:\d{2}).*?'  # 시간 범위 (예: 09:00-18:00)
            r'(\d+)\s*(?:분|m|min)',  # 휴게시간 (예: 60분 또는 60m)
            re.S | re.I
        )
        matches = pattern.findall(raw_text)

        # 기존 테이블 데이터 삭제
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # 변수 초기화
        grand_total_weighted = 0
        grand_total_actual = 0  # 실제 OT 시간 합계
        selected_year = int(self.year_var.get())  # 선택된 연도 가져오기
        processed_count = 0

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
                
                # 실제 OT 시간 계산 (8시간 초과분)
                actual_ot = max(0, net_hours - 8)
                grand_total_actual += actual_ot
                
                # 환산 시간 계산 (법정 가중치 적용)
                if is_holiday:
                    # 휴일 근무: 8시간까지 1.5배, 초과분 2.0배
                    type_str = f"휴일({holiday_name if holiday_name else day_name})"
                    if net_hours <= 8:
                        weighted_hours = net_hours * 1.5
                    else:
                        weighted_hours = (8 * 1.5) + ((net_hours - 8) * 2.0)
                else:
                    # 평일 근무: 8시간까지 1.0배, 초과분 1.5배
                    type_str = "평일"
                    weighted_hours = net_hours + (max(0, net_hours - 8) * 0.5)
                
                grand_total_weighted += weighted_hours
                processed_count += 1

                # 테이블에 행 추가
                self.tree.insert("", "end", values=(
                    f"{date_val}({day_name})",
                    f"{start_time}-{end_time}",
                    f"{rest_minutes}분",
                    f"{net_hours:.1f}h",
                    type_str,
                    f"{weighted_hours:.1f}h"
                ))
                
            except Exception as e:
                # 개별 행 처리 실패 시 스킵 (전체 처리는 계속)
                print(f"⚠ Failed to process row: {match} - {e}")
                continue

        # 결과 요약 표시
        self.summary_box.delete("0.0", "end")
        self.summary_box.tag_config("center", justify='center')
        
        if processed_count > 0:
            summary_text = (
                f"\nACTUAL OT: {grand_total_actual:.1f} HOURS  |  "
                f"WEIGHTED OT: {grand_total_weighted:.1f} HOURS\n"
                f"({processed_count} days processed)"
            )
            self.summary_box.insert("0.0", summary_text, "center")
        else:
            # 데이터가 없는 경우
            error_text = "\n⚠ No overtime data detected\n\nPlease check:\n• Screenshot quality\n• Date format (MM/DD)\n• Time format (HH:MM)"
            self.summary_box.insert("0.0", error_text, "center")
            messagebox.showwarning(
                "No Data Found",
                "Could not extract overtime data from the image.\n\n"
                "Please ensure:\n"
                "1. Screenshot shows clear date and time information\n"
                "2. Format: MM/DD HH:MM-HH:MM with rest time in minutes\n"
                "3. Image is not blurry or too dark"
            )

# =============================================================================
# 4. 프로그램 진입점
# =============================================================================
if __name__ == "__main__":
    app = OTCalculator()
    app.mainloop()
