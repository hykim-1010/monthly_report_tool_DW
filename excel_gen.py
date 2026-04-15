import os
import shutil
from pathlib import Path

import openpyxl


# 유입경로 시트의 채널 고정 순서 (행 2~10에 대응)
CHANNEL_ORDER = [
    "Referral",
    "Organic Search",
    "Direct",
    "Organic Social",
    "Unassigned",
    "Paid Search",
    "Display",
    "Organic Video",
    "Cross-network",
]

# 언어별 유입경로 시트 열 위치
CHANNEL_COLS = {"ko": "B", "en": "D", "cn": "F"}


def _is_formula(cell) -> bool:
    """수식 셀 여부 확인 (덮어쓰기 방지)"""
    return isinstance(cell.value, str) and cell.value.startswith("=")


def _safe_write(ws, cell_addr: str, value) -> None:
    """수식 셀이 아닌 경우에만 값을 입력한다."""
    cell = ws[cell_addr]
    if _is_formula(cell):
        raise ValueError(f"{cell_addr} 셀은 수식 셀입니다. 덮어쓸 수 없습니다.")
    cell.value = value


def _month_to_row(month: int) -> int:
    """보고월을 시트 행 번호로 변환한다. (1월=4행, 2월=5행, ...)"""
    return month + 3


def _build_channel_map(channel_data: list[dict]) -> dict:
    """
    채널 데이터를 {채널명: 세션수} 딕셔너리로 변환한다.
    GA4에서 반환되지 않은 채널은 0으로 채운다.
    """
    mapping = {ch: 0 for ch in CHANNEL_ORDER}
    for item in channel_data:
        ch = item["channel"]
        if ch in mapping:
            mapping[ch] = item["sessions"]
    return mapping


def write_report(
    template_path: str,
    output_dir: str,
    client_name: str,
    year: int,
    month: int,
    data: dict,
) -> str:
    """
    GA4 데이터를 엑셀 템플릿에 입력하고 output 폴더에 저장한다.

    Args:
        template_path: 템플릿 엑셀 파일 경로
        output_dir:    결과 파일 저장 폴더
        client_name:   고객사명 (파일명에 사용)
        year:          보고 연도 (예: 2026)
        month:         보고월 (예: 3)
        data: {
            "ko": {"users": int, "sessions": int, "pageviews": int,
                   "channels": [{"channel": str, "sessions": int}, ...]},
            "en": {...},
            "cn": {...},
        }

    Returns:
        저장된 파일 경로
    """
    # 출력 파일 경로 설정
    output_path = os.path.join(output_dir, f"{client_name}_{year}{month:02d}_report.xlsx")
    Path(output_dir).mkdir(parents=True, exist_ok=True)

    # 템플릿 복사 후 수정
    shutil.copy2(template_path, output_path)
    wb = openpyxl.load_workbook(output_path)

    row = _month_to_row(month)

    # ── 시트1: 사용자수 ──────────────────────────────────────
    ws_users = wb["사용자수"]
    _safe_write(ws_users, f"B{row}", data["ko"]["users"])
    _safe_write(ws_users, f"C{row}", data["en"]["users"])
    _safe_write(ws_users, f"D{row}", data["cn"]["users"])

    # ── 시트2: 페이지뷰수(조회수) ────────────────────────────
    ws_pv = wb["페이지뷰수(조회수)"]
    _safe_write(ws_pv, f"B{row}", data["ko"]["pageviews"])
    _safe_write(ws_pv, f"C{row}", data["en"]["pageviews"])
    _safe_write(ws_pv, f"D{row}", data["cn"]["pageviews"])

    # ── 시트3: 세션수 ────────────────────────────────────────
    ws_sessions = wb["세션수"]
    _safe_write(ws_sessions, f"B{row}", data["ko"]["sessions"])
    _safe_write(ws_sessions, f"C{row}", data["en"]["sessions"])
    _safe_write(ws_sessions, f"D{row}", data["cn"]["sessions"])

    # ── 시트4: 유입경로 ──────────────────────────────────────
    ws_ch = wb["유입경로"]
    for lang, col in CHANNEL_COLS.items():
        ch_map = _build_channel_map(data[lang]["channels"])
        for i, channel in enumerate(CHANNEL_ORDER):
            cell_addr = f"{col}{i + 2}"  # 채널 데이터는 2행부터 시작
            _safe_write(ws_ch, cell_addr, ch_map[channel])

    wb.save(output_path)
    print(f"저장 완료: {output_path}")
    return output_path
