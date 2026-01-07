import easyocr
import os

# 모델을 저장할 경로 (현재 작업 디렉토리 하위)
model_dir = './models'
if not os.path.exists(model_dir):
    os.makedirs(model_dir)

# 리더를 초기화하면 자동으로 모델이 다운로드됩니다.
# model_storage_directory를 지정하여 특정 폴더에 저장합니다.
print("EasyOCR 모델 다운로드 중...")
reader = easyocr.Reader(['ko', 'en'], gpu=False, model_storage_directory=model_dir)
print("다운로드 완료.")
