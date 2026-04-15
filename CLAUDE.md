# 프로젝트 개요
GA4 데이터를 조회해서 엑셀/PPT 보고서를 자동 생성하는 도구.
웹에이전시 운영 담당자가 매월 수작업으로 하던 보고서 작업을 자동화하는 것이 목적.

# 폴더 구조
/templates     → 고객사별 엑셀/PPT 원본 템플릿 보관
/output        → 생성된 보고서 파일 저장
/config        → 고객사 설정 JSON 파일
/logs          → 실행 로그
main.py        → 실행 진입점
ga4_client.py  → GA4 데이터 조회 모듈 (완성)
excel_gen.py   → 엑셀 생성 모듈 (완성)
ppt_gen.py     → PPT 생성 모듈 (완성)

# 고객사 설정
config/clients.json에서 고객사별 설정 관리
항목: 고객사명, 언어별 GA4 속성 ID, 템플릿 파일명

# 언어 구분 방식
국/영/중문 GA4 속성이 각각 분리되어 있음 (속성 3개)
언어별 속성 ID는 config/clients.json의 ga4_property_ids.ko / .en / .cn 으로 관리
URL prefix 필터 없이 속성 ID 자체로 언어 구분

# 보안 원칙
- 고객사 Google 계정 로그인 정보는 절대 저장하지 않음
- 서비스 계정 키(.json)는 .env로 관리, git에 절대 올리지 않음
- GA4 조회 권한만 사용 (읽기 전용)
- .gitignore에 반드시 포함: .env, /config/*.json, 서비스계정키.json

# 환경변수 (.env)
GOOGLE_KEY_PATH=config/ga4-automation-479406-eab0c281dd18.json

# 주요 라이브러리
- google-analytics-data (GA4 Data API v1)
- openpyxl (엑셀 읽기/쓰기)
- python-pptx (PPT 수정)
- python-dotenv (환경변수)

# GA4 조회 항목 정의
- 사용자수     → totalUsers          (fetch_summary)
- 세션수       → sessions            (fetch_summary)
- 페이지뷰수   → screenPageViews     (fetch_summary)
- 유입채널     → sessionDefaultChannelGroup  (fetch_channel_sessions)
- 인기페이지   → pagePath TOP 10     (fetch_top_pages)
- 평균참여시간 → averageSessionDuration      (fetch_avg_engagement)

# 엑셀 템플릿 구조 (운영보고서_data_기본_V1.0_260326.xlsx)

## 시트 목록 (index 순)
0: 사용자수
1: 페이지뷰수(조회수)
2: 세션수
3: 유입경로
4~7: 기타 시트 (자동화 대상 아님)

## 공통 행 구조 (시트 0~2)
- 1행: 연도 헤더 (A1=2026년 실적, I1 또는 G1=2025년 실적)
- 2행: 지표명 헤더
- 3행: 언어 헤더 (B3=국, C3=영, D3=중, E3=합계)
- 4행~15행: 1월~12월 데이터
- 16행: 합계 (SUM 수식)

## 보고월 → 행 번호 변환
month + 3 = 행 번호  (예: 3월 → 6행, 4월 → 7행)

## 입력 대상 셀 (3월 기준 예시, 6행)
시트0 사용자수:        B6(국문), C6(영문), D6(중문)
시트1 페이지뷰수:      B6(국문), C6(영문), D6(중문)
시트2 세션수:          B6(국문), C6(영문), D6(중문)
시트3 유입경로(국문):  B2~B10 (채널별 세션수)
시트3 유입경로(영문):  D2~D10
시트3 유입경로(중문):  F2~F10

## 수식 셀 (절대 덮어쓰지 말 것)
E열(합계 SUM), F~H열(전월 대비 증감률), J열(전년 대비 증감률)
excel_gen.py의 _safe_write()가 수식 셀 감지 시 예외 발생

## 유입경로 채널 고정 순서 (행 2~10)
Referral / Organic Search / Direct / Organic Social / Unassigned
/ Paid Search / Display / Organic Video / Cross-network

# PPT 슬라이드 구조 (운영보고서_기본_V1.0_260327.pptx)

## 슬라이드 목록
1: 표지 — TextBox 2(제목), TextBox 3(날짜)
2: 월간 운영 업무 내역 — 수동 입력
3: 사용자수 — 표 6 + TextBox 1(요약문)
4: 페이지뷰수 — 표 6 + TextBox 1(요약문)
5: 유입경로 — 표 3
6: 인기 페이지 — 표 3(국문), 표 4(영문), 표 7(중문)
7: 마지막 슬라이드 — 해당 없음

## 슬라이드 3~4: 사용자수 / 페이지뷰수 표 (표 6, 15행×7열)
- 보고월 행 index: month + 1 (1월=2, 3월=4)
- 입력: (month+1, 1)=국문, (month+1, 2)=영문, (month+1, 3)=중문
- 유지: col4(소계), col5(2025실적), col6(증가율), 행14(합계)

## 슬라이드 5: 유입경로 표 (표 3, 11행×7열)
- 채널 행 1~9 (Referral / Organic Search / Direct / Organic Social / Unassigned / Paid Search / Display / Organic Video / Cross-network)
- 입력: col1(국문 세션수), col3(영문 세션수), col5(중문 세션수)
- 비율 열(col2/4/6)은 PPT에 수식 없음 → 직접 계산하여 함께 업데이트
- 합계 행(행10)도 함께 업데이트

## 슬라이드 6: 인기 페이지 표 3개 (각 12행×3열)
- 표 3(국문), 표 4(영문), 표 7(중문)
- 입력: 행 2~11, col1(페이지명), col2(페이지뷰수)
- 유지: 행 0(언어 헤더), 행 1(컬럼 헤더)

## ppt_gen.py 수정 시 주의사항
- 슬라이드 index 고정: slides[0]=표지, slides[2]=사용자수, slides[3]=페이지뷰수, slides[4]=유입경로, slides[5]=인기페이지
  → 슬라이드 순서 바뀌면 write_report() 내 index 일괄 수정 필요
- CHANNEL_ORDER 상수: 채널 추가/삭제 시 ppt_gen.py 상단 리스트만 수정하면 됨
- _fill_cover() L72: "더위버크리에이티브" 문구 하드코딩 → 추후 clients.json으로 이동 예정
- _set_cell(): 런(run)이 없는 셀은 para.text로 폴백 — 이 경우 서식(폰트·색상) 손실 가능
- _set_textbox(): 단락이 여러 개인 텍스트박스는 첫 번째 단락만 수정됨

# 코딩 규칙
- 함수마다 한국어 주석 필수
- 오류 발생 시 에러 메시지를 한국어로 출력
- 파일명 규칙: {고객사명}_{YYYYMM}_report.xlsx / .pptx
- 모든 생성 결과는 logs/run_log.txt에 기록

# 실행 방식
python main.py --client 고객사명 --start 20260201 --end 20260228

# 진행 단계

## 완료
- 1단계: 프로젝트 세팅 (폴더 구조, requirements.txt)
- 2단계: config/clients.json 설계 및 GA4 속성 ID 등록
- 3단계: ga4_client.py — fetch_summary, fetch_channel_sessions, fetch_top_pages, fetch_avg_engagement
- 4-1단계: 엑셀 템플릿 시트 구조 분석 및 셀 매핑 확인
- 4-2단계: excel_gen.py — 템플릿 복사 후 GA4 데이터 자동 입력, output 저장
- 5-1단계: PPT 템플릿 슬라이드 구조 분석 (도형명, 표 행/열 매핑)
- 5-2단계: ppt_gen.py — 표지·수치표·유입경로·인기페이지 자동 입력, output 저장

## 예정
- 6단계: main.py — 전체 흐름 통합 (GA4 조회 → 엑셀 → PPT → 로그)

# 1차 자동화 범위 (현재 버전)
- 엑셀: 사용자수/세션수/페이지뷰수/유입경로 시트 자동 입력
- PPT: 보고 기간 문구, 주요 수치 표, 인기 페이지 TOP 10
- 요약 문장: 전월 대비 증감률 기반 고정 규칙으로 자동 생성

# 아직 수동으로 처리하는 것 (2차 예정)
- 레이아웃 깨짐 여부 최종 확인
- 인기 페이지 명칭 자연스러운지 검수
- 고객사별 특이사항 코멘트 추가
