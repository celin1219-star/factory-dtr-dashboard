Factory DTR/PSS-E Dashboard — Web Deploy Package
==================================================

이 폴더 전체를 GitHub 저장소에 업로드한 뒤 Render Web Service로 배포하면
VSCode/Python 설치 없이 누구나 공개 URL로 접속할 수 있습니다.

필수 파일
- app.py
- requirements.txt
- render.yaml
- .python-version
- 온도최종.xlsx
- MVA최종_DTR재산정_120C_IEC60K (1)(1).xlsx
- 전압 분단위 최종(1).xlsx
- 보상최종_기존공장_신규30MVA_보강전후_통합그래프.xlsx

Render 설정
- Build Command: pip install -r requirements.txt
- Start Command: gunicorn app:server --bind 0.0.0.0:$PORT --workers 1 --threads 4 --timeout 120
- Health Check Path: /
- Python: .python-version = 3.13

로컬 확인
python app.py --check
python app.py
브라우저: http://127.0.0.1:8050
