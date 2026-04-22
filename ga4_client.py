import os
from urllib.parse import unquote
from google.analytics.data_v1beta import BetaAnalyticsDataClient
from google.analytics.data_v1beta.types import (
    RunReportRequest,
    DateRange,
    Dimension,
    Metric,
    OrderBy,
)
from google.oauth2 import service_account


def get_ga4_client() -> BetaAnalyticsDataClient:
    # GitHub Actions: GOOGLE_SERVICE_ACCOUNT_JSON 환경변수에서 직접 읽음
    key_json = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")
    if key_json:
        import json
        info = json.loads(key_json)
        credentials = service_account.Credentials.from_service_account_info(
            info,
            scopes=["https://www.googleapis.com/auth/analytics.readonly"],
        )
        return BetaAnalyticsDataClient(credentials=credentials)

    # 로컬: .env의 GOOGLE_KEY_PATH 파일에서 읽음
    key_path = os.getenv("GOOGLE_KEY_PATH")
    if not key_path:
        raise ValueError("GOOGLE_KEY_PATH 또는 GOOGLE_SERVICE_ACCOUNT_JSON이 설정되지 않았습니다.")

    credentials = service_account.Credentials.from_service_account_file(
        key_path,
        scopes=["https://www.googleapis.com/auth/analytics.readonly"],
    )
    return BetaAnalyticsDataClient(credentials=credentials)


def fetch_summary(
    property_id: str,
    start_date: str,
    end_date: str,
) -> dict:
    """
    GA4에서 기간별 요약 지표를 조회한다.
    언어 구분은 속성 ID 자체로 처리 (ko/en/cn 속성이 분리되어 있음).

    Args:
        property_id: 언어별 GA4 속성 ID (예: "properties/123456789")
        start_date:  조회 시작일 (예: "2026-03-01")
        end_date:    조회 종료일 (예: "2026-03-31")

    Returns:
        {
            "users": int,
            "sessions": int,
            "pageviews": int,
        }
    """
    client = get_ga4_client()

    request = RunReportRequest(
        property=property_id,
        date_ranges=[DateRange(start_date=start_date, end_date=end_date)],
        metrics=[
            Metric(name="totalUsers"),
            Metric(name="sessions"),
            Metric(name="screenPageViews"),
        ],
    )

    response = client.run_report(request)
    row = response.rows[0]

    return {
        "users":     int(row.metric_values[0].value),
        "sessions":  int(row.metric_values[1].value),
        "pageviews": int(row.metric_values[2].value),
    }


def fetch_channel_sessions(
    property_id: str,
    start_date: str,
    end_date: str,
) -> list[dict]:
    """
    GA4에서 유입 채널별 세션수를 조회한다.

    Args:
        property_id: 언어별 GA4 속성 ID
        start_date:  조회 시작일 (예: "2026-03-01")
        end_date:    조회 종료일 (예: "2026-03-31")

    Returns:
        [{"channel": str, "sessions": int}, ...]  # 세션수 내림차순
    """
    client = get_ga4_client()

    request = RunReportRequest(
        property=property_id,
        date_ranges=[DateRange(start_date=start_date, end_date=end_date)],
        dimensions=[Dimension(name="sessionDefaultChannelGroup")],
        metrics=[Metric(name="sessions")],
        order_bys=[OrderBy(metric=OrderBy.MetricOrderBy(metric_name="sessions"), desc=True)],
    )

    response = client.run_report(request)

    return [
        {
            "channel":  row.dimension_values[0].value,
            "sessions": int(row.metric_values[0].value),
        }
        for row in response.rows
    ]


def fetch_top_pages(
    property_id: str,
    start_date: str,
    end_date: str,
    limit: int = 10,
) -> list[dict]:
    """
    GA4에서 페이지뷰 기준 인기 페이지를 조회한다.

    Args:
        property_id: 언어별 GA4 속성 ID
        start_date:  조회 시작일 (예: "2026-03-01")
        end_date:    조회 종료일 (예: "2026-03-31")
        limit:       반환할 페이지 수 (기본 10)

    Returns:
        [{"page": str, "pageviews": int}, ...]  # 페이지뷰 내림차순
    """
    client = get_ga4_client()

    request = RunReportRequest(
        property=property_id,
        date_ranges=[DateRange(start_date=start_date, end_date=end_date)],
        dimensions=[Dimension(name="unifiedPageScreen")],
        metrics=[Metric(name="screenPageViews")],
        order_bys=[OrderBy(metric=OrderBy.MetricOrderBy(metric_name="screenPageViews"), desc=True)],
        limit=25,
    )

    response = client.run_report(request)

    aggregated: dict[str, int] = {}
    for row in response.rows:
        page = row.dimension_values[0].value
        pageviews = int(row.metric_values[0].value)
        key = _normalize_top_page_key(page)
        aggregated[key] = aggregated.get(key, 0) + pageviews

    return [
        {"page": page, "pageviews": pageviews}
        for page, pageviews in sorted(
            aggregated.items(),
            key=lambda item: item[1],
            reverse=True,
        )[:limit]
    ]


def _normalize_top_page_key(page_title: str) -> str:
    """TOP 페이지 집계와 PPT 표시를 위한 국문 표준 페이지명으로 정규화한다."""
    cleaned = unquote(page_title or "").strip().strip("/")
    if not cleaned:
        return "메인"

    normalized = (
        cleaned.lower()
        .replace("_", " ")
        .replace("-", " ")
        .replace("/", " > ")
    )
    compact = normalized.replace(" ", "")

    if _contains_any(
        compact,
        [
            "index",
            "main",
            "home",
            "homepage",
            "메인",
            "대상웰라이프",
            "daesangwellife",
            "daesangwelllife",
            "大象wellife",
            "大象welllife",
        ],
    ):
        return "메인"

    direct_rules = [
        (["news>list", "news>view", "newslist", "newsview", "media>news", "news"], "뉴스"),
        (["nucarebalancednutrition", "newcarebalancednutrition", "balancednutrition"], "브랜드 > 뉴케어 > 균형영양식"),
        (["nucareallprotein", "newcareallprotein", "allprotein"], "브랜드 > 뉴케어 > 올프로틴"),
        (["nucarespecializednutrition", "specializednutrition"], "브랜드 > 뉴케어 > 전문영양식"),
        (["nucareglucoseplan", "glucoseplan"], "브랜드 > 뉴케어 > 당플랜"),
        (["healthsupplement", "healthfunctionalfood", "functionalfood"], "브랜드 > 건강기능식품"),
        (["gutsys"], "브랜드 > 것시스"),
        (["nucare", "newcare", "뉴케어"], "브랜드 > 뉴케어"),
        (["wellifesolution", "welllifesolution"], "고객지원 > Welllife Solution"),
        (["businesslocation", "사업장위치"], "회사소개 > 사업장 위치"),
        (["affiliate", "affiliated", "계열사"], "회사소개 > 계열사 소개"),
        (["companyoverview", "회사개요"], "회사소개 > 회사개요"),
        (["companyintroduction", "companyprofile", "회사소개"], "회사소개 > 회사소개"),
        (["rdcenter", "r&dcenter", "r&d센터"], "R&D > R&D센터 소개"),
        (["rdresult", "r&dresult"], "R&D > R&D 결과"),
        (["irdat", "irdata", "announce", "announcement", "공고"], "IR > IR 자료 > 공고"),
        (["faq"], "고객지원 > FAQ"),
    ]
    for needles, label in direct_rules:
        if _contains_any(compact, needles):
            return label

    parts = _split_page_title(cleaned)
    labels = [_translate_page_part(part) for part in parts]
    deduped = []
    for label in labels:
        if not label or (deduped and deduped[-1] == label):
            continue
        deduped.append(label)

    return " > ".join(deduped) if deduped else cleaned


def _contains_any(text: str, needles: list[str]) -> bool:
    """정규화된 문자열이 후보 중 하나를 포함하는지 확인한다."""
    return any(needle.replace(" ", "") in text for needle in needles)


def _split_page_title(page_title: str) -> list[str]:
    """페이지 제목을 메뉴 단위 후보로 분해한다."""
    if ">" in page_title:
        return [part.strip() for part in page_title.split(">") if part.strip()]
    return [page_title.strip()]


def _translate_page_part(part: str) -> str:
    """페이지 제목 일부를 국문 메뉴명으로 변환한다."""
    normalized = part.lower().replace("_", " ").replace("-", " ").strip()
    compact = normalized.replace(" ", "")
    mapping = {
        "about": "회사소개",
        "aboutus": "회사소개",
        "company": "회사소개",
        "companyoverview": "회사개요",
        "overview": "회사개요",
        "companyintroduction": "회사소개",
        "companyprofile": "회사소개",
        "brand": "브랜드",
        "nucare": "뉴케어",
        "newcare": "뉴케어",
        "balancednutrition": "균형영양식",
        "healthsupplement": "건강기능식품",
        "healthfunctionalfood": "건강기능식품",
        "news": "뉴스",
        "media": "미디어",
        "support": "고객지원",
        "faq": "FAQ",
        "ir": "IR",
        "r&d": "R&D",
        "rd": "R&D",
        "rdcenter": "R&D센터 소개",
        "location": "사업장 위치",
        "businesslocation": "사업장 위치",
    }
    return mapping.get(compact, part.strip())


def fetch_avg_engagement(
    property_id: str,
    start_date: str,
    end_date: str,
) -> float:
    """
    GA4에서 평균 참여 시간(초)을 조회한다.

    Args:
        property_id: 언어별 GA4 속성 ID
        start_date:  조회 시작일 (예: "2026-03-01")
        end_date:    조회 종료일 (예: "2026-03-31")

    Returns:
        평균 참여 시간 (초, 소수점 1자리)
    """
    client = get_ga4_client()

    request = RunReportRequest(
        property=property_id,
        date_ranges=[DateRange(start_date=start_date, end_date=end_date)],
        metrics=[
            Metric(name="userEngagementDuration"),
            Metric(name="activeUsers"),
        ],
    )

    response = client.run_report(request)
    row = response.rows[0]
    engagement_duration = float(row.metric_values[0].value)
    active_users = int(row.metric_values[1].value)
    if active_users == 0:
        return 0.0

    return round(engagement_duration / active_users, 1)
