# -*- coding: utf-8 -*-
import os
import re
import glob
import time
import json
import functools
import threading
import unicodedata
import xml.etree.ElementTree as ET
from concurrent.futures import ThreadPoolExecutor, as_completed
import pdfplumber
import pandas as pd
import requests
from difflib import SequenceMatcher  # 문자열 유사도 비교를 위해 추가

# OCR(pdf2image/pytesseract)은 제거했다. 스캔본/폰트 깨짐 파일을 Tesseract로
# 복구해도 오탈자가 섞여 정규식이 어차피 잘 안 맞고, 속도도 파일당 5~10초로
# 느려서 배치 전체를 느리게 만드는 원인이었다. 이런 파일은 텍스트 기반
# 정규식/LLM보다 Vision LLM(이미지를 통째로 읽는 방식)이 훨씬 정확하고
# 간단하므로, 여기서는 "스캔본/폰트 깨짐"임을 표시만 해두고(비고), 실제
# 복구는 2단계에서 이미지를 Vision LLM에 넘겨 처리한다(별도 단계, 자동 실행 안 함).


# =====================================================================
# Colab 임시 저장소 경로 설정
PDF_DIR = "/content/papers"
OUTPUT_XLSX = "/content/참고문헌_메타데이터.xlsx"
CROSSREF_TIMEOUT = 15
CROSSREF_MAX_RETRIES = 3
CROSSREF_HEADERS = {
    "User-Agent": "biblio-extractor/1.0 (mailto:example@example.com)"
}

# --- 2단계(LLM 재검토) 설정 ---
# 1단계(정규식/CrossRef)에서 "확인필요" 또는 저신뢰도로 표시된 행만 골라
# Gemini API로 재확인시킨다(무료 티어 사용). Colab에서 실행하기 전에
# 반드시 API 키를 설정해야 한다: os.environ["GEMINI_API_KEY"] = "AIza..."
# (키 발급: https://aistudio.google.com/apikey)
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")
GEMINI_MODEL = "gemini-3.5-flash-lite"
# [수정됨] gemini-2.0-flash에 이어 gemini-2.5-flash도 신규 사용자에게는
# "더 이상 제공되지 않음"(404)으로 확인됐다 - 모델 세대가 계속 빠르게
# 넘어가는 중이라, 코드에 특정 모델명을 박아두면 금방 또 깨질 수 있다.
# 이 글을 쓰는 시점 기준 정식 출시(GA)된 최신 경량 모델로 바꿨다. "-latest"
# 별칭(예: gemini-flash-latest)은 안정적일 것 같지만 실제로 그 별칭 자체가
# 재차 폐기된 사례가 있어(가리키는 실체 모델이 바뀌다가 통째로 사라짐)
# 버전이 명시된 모델명을 직접 쓰는 편이 더 안전하다. 이후에도 404/429가
# 나면 https://ai.google.dev/gemini-api/docs/models 에서 현재 제공되는
# 모델명을 확인해서 이 값을 바꾸면 된다.
GEMINI_API_URL = f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent"
LLM_TIMEOUT = 60
LLM_MAX_RETRIES = 3
# [수정됨] 성공한 요청 사이에 쉬는 시간. 무료 티어 분당 요청 한도(RPM)가
# 모델/시점에 따라 낮게는 분당 5회 수준까지 축소된 것으로 확인됐다
# (2025년 12월 이후 여러 차례 축소됨 - 정확한 현재 수치는
# https://ai.google.dev/gemini-api/docs/rate-limits 에서 확인). 분당 5회면
# 요청 사이 최소 12초는 띄워야 429를 안 만난다. 실패해서 재시도하는 것도
# 전부 호출로 카운트되므로, 여유 있게 잡아둔다. 429가 계속 뜨면 이 값을
# 더 늘리면 된다.
LLM_REQUEST_INTERVAL = 12
# =====================================================================

# =====================================================================
# [추가됨] GROBID 헤더 추출 설정
# =====================================================================
# 지금까지 하나씩 막아온 오탐(표 헤더 "Sex"가 교신저자로, 참고문헌 약칭
# "Bibl. Lichenol."이 교신저자로 채택되는 등)은 전부 "정규식 + 줄 위치
# 규칙"이라는 접근법 자체의 구조적 한계에서 나온다. GROBID는 논문 수십만
# 편으로 학습된 모델로 각 줄을 "제목/저자/소속/이메일"로 분류하기 때문에
# 이런 유형의 오탐이 구조적으로 훨씬 적다(Semantic Scholar, ResearchGate
# 등이 실제로 이걸 씀). 그래서 우선순위를 다음과 같이 바꾼다:
#   1) 본문에서 DOI를 직접 찾았으면 -> CrossRef 조회 (가장 정확, 기존 그대로)
#   2) DOI가 없거나 CrossRef가 실패 -> GROBID로 헤더 추출
#      -> GROBID가 준 제목으로 CrossRef 서지검색을 한 번 더 시도해서
#         DOI 없는 논문도 CrossRef의 정식 메타데이터를 최대한 끌어옴
#   3) GROBID도 실패(서버 다운/스캔본 등) -> 기존 정규식 폴백 (최후 수단)
#
# 사용 전 GROBID 서버를 띄워야 한다: `docker run -t --rm -p 8070:8070
# grobid/grobid:0.8.0` (자세한 내용: https://grobid.readthedocs.io).
# 서버가 꺼져 있어도 파이프라인이 죽지 않도록, 최초 1회만 접속을 확인하고
# 안 되면 이후 호출은 전부 조용히 건너뛰어 기존 폴백으로 넘어간다.
GROBID_URL = os.environ.get("GROBID_URL", "http://localhost:8070")
GROBID_TIMEOUT = 60  # 헤더 추출은 보통 수 초 내로 끝나지만 대기 여유를 둠
GROBID_TEI_NS = {"tei": "http://www.tei-c.org/ns/1.0"}

# --- 10,000개 이상 대량 배치를 위한 설정 ---
# 파일 하나하나가 GROBID/CrossRef 같은 네트워크 호출을 기다리는 I/O 위주
# 작업이라 스레드 병렬화 효과가 크다. 워커 수는 GROBID 서버의 CPU 코어
# 수에 맞춰 조절하면 된다(코어 수보다 과하게 높이면 오히려 서버 쪽에서
# 큐잉이 걸려 느려질 수 있음).
MAX_WORKERS = 8
# 대량 배치는 중간에 끊기는 일이 흔하므로, 처리된 결과를 파일 하나씩
# 끝나는 즉시 여기 이어서 기록한다(JSON Lines). 재실행 시 이미 끝난
# 파일명은 자동으로 건너뛰고 나머지만 이어서 처리한다.
CHECKPOINT_PATH = "/content/checkpoint.jsonl"
# =====================================================================

COVER_PAGE_MARKERS = (
    "To cite this article", "Full Terms & Conditions of access",
    "Journal homepage:", "Submit your article to this journal",
)

def is_cover_page(page_text):
    return any(marker in page_text for marker in COVER_PAGE_MARKERS)

def get_article_pages(pdf):
    if not pdf.pages:
        return 0
    first_text = pdf.pages[0].extract_text() or ""
    if is_cover_page(first_text) and len(pdf.pages) > 1:
        return 1
    return 0

def get_pdf_text(pdf, max_pages=3):
    start = get_article_pages(pdf)
    texts = []
    for p in pdf.pages[start:start + max_pages]:
        # 기본값(x_tolerance=3)으로 뽑으면 일부 PDF(폰트 인코딩 특성상)는
        # 단어 사이 공백이 통째로 사라져 "Kyung-HwaPark"처럼 붙어버린다.
        # 그러면 "이름은 2단어 이상" 검증에서 무조건 탈락하므로, 값을 낮춰
        # 공백을 보존한다(다른 정상 PDF에는 부작용 없음을 확인함).
        raw = p.extract_text(x_tolerance=1) or ""
        # NFKC 정규화: 일부 PDF(특히 옛 판형)는 전각(全角) 괄호/숫자나
        # 리가처(ligature, 예: "ﬁ"가 한 글자로 합쳐진 "fi")를 쓰는데,
        # 이런 문자는 정규식의 일반 ASCII 패턴과 안 맞아 매칭이 실패한다.
        # NFKC 정규화로 호환 등가 문자를 표준 형태로 통일하면 이런 문제를
        # 줄일 수 있다(하이픈/en-dash 등 의미가 다른 문자는 NFKC가 서로
        # 합치지 않으므로 기존 정규식에 부작용이 없음을 확인했다).
        texts.append(unicodedata.normalize("NFKC", raw))
    return "\n".join(texts), texts[0] if texts else ""

DOI_REGEX = re.compile(r"10\.\d{4,9}/[-._;()/:A-Za-z0-9]+")

def extract_doi(text):
    joined_text = text.replace("\n", "")
    matches = DOI_REGEX.findall(text)
    if not matches:
        # 슬래시 직전에서 줄바꿈이 걸리면("10.1111\n/j...") 원문에서는
        # 아예 매치가 안 되므로, 줄바꿈을 제거한 텍스트에서 다시 시도한다.
        joined_only = DOI_REGEX.findall(joined_text)
        return joined_only[0].rstrip(").,;") if joined_only else None
    doi = matches[0]

    # PDF 줄바꿈 위치가 DOI 중간에 걸리면(예: 페이지 여백에서
    # "10.1111/j" 까지만 뽑히고 ".1748-5967...x"가 다음 줄로 밀려나는 경우)
    # 잘린 DOI가 그대로 채택되어 CrossRef 조회가 실패한다. 줄바꿈을 제거한
    # 텍스트에서 doi 바로 뒤에 오는 부분을 살펴보되, 숫자·점·하이픈과
    # 맨 끝의 글자 하나(".x" 같은 관례적 접미사)만 이어붙인다. 그렇지 않으면
    # "...00181.x\nworkers (268.0..." 처럼 DOI 뒤에 이어지는 무관한 다음
    # 문장이 그대로 DOI에 붙어버리는 사고가 난다.
    #
    # 추가 가드: doi가 이미 이 저널들의 관례적인 완결 형태(".x"로 끝남)라면
    # 애초에 확장을 시도하지 않는다. 안 그러면 "...00048.x" 뒤에 표에 있는
    # 숫자(예: p-value "0.17")가 줄바꿈만 사이에 두고 있을 때 그 숫자까지
    # DOI에 붙어버리는(예: "...00048.x0.17") 사고가 난다.
    already_complete = bool(re.search(r"\.[a-z]$", doi))
    if not already_complete:
        idx = joined_text.find(doi)
        if idx != -1:
            remainder = joined_text[idx + len(doi):]
            m_extra = re.match(r"[\d.\-]+[a-z]?", remainder)
            if m_extra and m_extra.group(0):
                doi = doi + m_extra.group(0)

    doi = doi.rstrip(").,;")
    return doi

def fetch_crossref_metadata(doi):
    # CrossRef 공개 API는 순간적인 타임아웃이나 429(과다요청)/5xx가 드물지
    # 않게 발생한다. 268건을 연달아 조회하는 배치 작업에서는 이게 누적되어
    # "DOI는 찾았는데 CrossRef만 실패" 케이스가 상당수 나온다. 재시도 +
    # 지수 백오프를 넣어서 일시적 실패를 흡수한다.
    last_exc = None
    for attempt in range(CROSSREF_MAX_RETRIES):
        try:
            with _CROSSREF_SEMAPHORE:
                resp = requests.get(
                    f"https://api.crossref.org/works/{doi}",
                    headers=CROSSREF_HEADERS,
                    timeout=CROSSREF_TIMEOUT,
                )
            if resp.status_code == 200:
                return resp.json().get("message")
            if resp.status_code == 429 or resp.status_code >= 500:
                # 과다요청/서버 오류는 잠시 후 재시도할 가치가 있다.
                retry_after = resp.headers.get("Retry-After")
                wait = float(retry_after) if retry_after and retry_after.isdigit() else (2 ** attempt)
                time.sleep(wait)
                continue
            # 404 등 확정적 실패는 재시도해도 소용없으므로 바로 포기한다.
            return None
        except requests.exceptions.RequestException as e:
            last_exc = e
            time.sleep(2 ** attempt)
    return None

def parse_crossref_year(message):
    for key in ("published-print", "published-online", "published", "issued"):
        dp = message.get(key, {}).get("date-parts")
        if dp and dp[0] and dp[0][0]:
            return str(dp[0][0])
    return ""

def parse_crossref_authors(message):
    authors = message.get("author") or []
    names = []
    for a in authors:
        given = a.get("given", "").strip()
        family = a.get("family", "").strip()
        name = f"{given} {family}".strip() if (given or family) else a.get("name", "")
        if name:
            names.append(name)
    all_authors = ", ".join(names)
    first_author = names[0] if names else ""
    return all_authors, first_author

def clean_crossref_text(s):
    """CrossRef가 학명 등에 붙여 반환하는 <i>, <scp> 같은 HTML 태그와
    불필요한 줄바꿈/중복 공백을 정리한다."""
    if not s:
        return s
    s = re.sub(r"<[^>]+>", "", s)          # <i>...</i>, <scp>...</scp> 등 태그 제거
    s = re.sub(r"\s+", " ", s).strip()      # 개행/중복 공백을 단일 공백으로
    return s


def _metadata_from_crossref_message(message, doi_fallback=""):
    """CrossRef API가 돌려준 message(JSON)를 결과 행 형태로 변환한다.
    DOI로 직접 조회했을 때와(metadata_from_crossref), 제목으로 서지검색
    했을 때(crossref_search_by_title) 양쪽에서 공유하는 파싱 로직이다."""
    title_list = message.get("title") or []
    title = clean_crossref_text(title_list[0]) if title_list else ""

    journal_list = message.get("container-title") or []
    journal = clean_crossref_text(journal_list[0]) if journal_list else ""

    volume = message.get("volume", "")
    page = message.get("page", "")
    year = parse_crossref_year(message)
    authors, first_author = parse_crossref_authors(message)
    authors = clean_crossref_text(authors)
    first_author = clean_crossref_text(first_author)
    doi = message.get("DOI", "") or doi_fallback

    return {
        "저자": authors or "확인필요",
        "년도": year or "확인필요",
        "제목": title or "확인필요",
        "저널 이름": journal or "확인필요",
        "볼륨": volume or "확인필요",
        "추가 정보(페이지)": page or "확인필요",
        "1저자": first_author or "확인필요",
        "DOI": doi or "확인필요",
    }


def metadata_from_crossref(doi):
    message = fetch_crossref_metadata(doi)
    if not message:
        return None
    return _metadata_from_crossref_message(message, doi_fallback=doi)


def crossref_search_by_title(title):
    """DOI를 본문에서 못 찾은 논문에 대해, (주로 GROBID가 뽑아준) 제목으로
    CrossRef 서지검색(bibliographic search)을 시도해 정식 DOI/메타데이터를
    역으로 찾는다. 검색 결과 1위가 실제로 같은 논문인지 확인하지 않으면
    완전히 다른 논문을 잘못 매칭할 위험이 크므로, 제목 유사도가 충분히
    높을 때만(0.88 이상) 채택한다 - 오탐 위험보다 놓치는 쪽을 택함."""
    if not title or len(title) < 8 or title == "확인필요":
        return None
    if _contains_hangul(title):
        # CrossRef 레코드는 이 저널군에서도 거의 항상 영문 제목으로 등록돼
        # 있다. 한글 제목으로 검색하면 애초에 매칭될 리 없어 API 호출만
        # 낭비하므로(그리고 유사도 0.88 문턱을 절대 못 넘으므로), 호출부가
        # 영문 제목(파일명 등)으로 바꿔서 넘기게 하고 여기서는 바로 포기한다.
        return None
    try:
        with _CROSSREF_SEMAPHORE:
            resp = requests.get(
                "https://api.crossref.org/works",
                params={"query.bibliographic": title, "rows": 3},
                headers=CROSSREF_HEADERS,
                timeout=CROSSREF_TIMEOUT,
            )
    except requests.exceptions.RequestException:
        return None
    if resp.status_code != 200:
        return None

    items = (resp.json().get("message") or {}).get("items") or []
    title_norm = re.sub(r"[^a-z0-9]", "", title.lower())
    for item in items:
        cand_title = clean_crossref_text((item.get("title") or [""])[0])
        cand_norm = re.sub(r"[^a-z0-9]", "", cand_title.lower())
        if not cand_norm:
            continue
        if SequenceMatcher(None, title_norm, cand_norm).ratio() >= 0.88:
            return item
    return None

# =====================================================================
# [추가됨] GROBID 클라이언트
# =====================================================================
@functools.lru_cache(maxsize=1)
def _grobid_enabled():
    """GROBID 서버 접속 가능 여부를 최초 1회만 확인하고 캐시한다(파일마다
    확인하면 서버가 꺼져 있을 때 10,000개 전부 개별 타임아웃을 기다리게
    되어 배치 전체가 느려진다). 접속 안 되면 경고만 찍고 이후 모든 GROBID
    호출은 조용히 건너뛰어 기존 정규식 폴백으로 넘어간다."""
    try:
        resp = requests.get(f"{GROBID_URL}/api/isalive", timeout=5)
        # 순수 GROBID는 200만 주지만, Hugging Face Spaces 같은 프록시 뒤에서는
        # 짧은 응답에 대해 206(Partial Content)을 돌려주는 경우가 실측 확인됐다.
        # 서버가 죽었으면 애초에 연결 자체가 안 되거나 5xx가 오므로, 2xx 전체를
        # "살아있음"으로 본다.
        ok = 200 <= resp.status_code < 300
    except requests.exceptions.RequestException:
        ok = False
    if not ok:
        print(
            f"⚠️ GROBID 서버({GROBID_URL})에 연결할 수 없습니다 - 이번 실행은 "
            f"GROBID 없이 기존 정규식 폴백만으로 진행합니다. "
            f"(예: `docker run -t --rm -p 8070:8070 grobid/grobid:0.8.0` 로 띄운 뒤 재실행)"
        )
    return ok


def _tei_text(elem):
    return " ".join(elem.itertext()).strip() if elem is not None else ""


def parse_grobid_header_xml(xml_bytes):
    """GROBID processHeaderDocument가 돌려주는 TEI XML에서 제목/저자/이메일/
    교신저자/저널/볼륨/페이지/연도/DOI를 구조화된 dict로 뽑는다. 제목이나
    저자가 하나도 안 잡히면(헤더 인식 자체가 실패한 경우) None을 반환해서
    호출부가 기존 정규식 폴백으로 넘어가게 한다."""
    try:
        root = ET.fromstring(xml_bytes)
    except ET.ParseError:
        return None
    ns = GROBID_TEI_NS

    title = _tei_text(root.find(".//tei:analytic/tei:title", ns))

    authors = []
    corresponding = None
    for author_el in root.findall(".//tei:analytic/tei:author", ns):
        persname = author_el.find(".//tei:persName", ns)
        if persname is None:
            continue
        forename = " ".join(
            _tei_text(f) for f in persname.findall("tei:forename", ns) if _tei_text(f)
        )
        surname = _tei_text(persname.find("tei:surname", ns))
        name = f"{forename} {surname}".strip()
        if not name:
            continue
        authors.append(name)
        # GROBID는 본문에 이메일이 명시된 저자만 이메일을 잡아내는데, 대부분
        # 그 저자가 교신저자다(각주의 "*Corresponding author <E-mail:...>"
        # 패턴과 본문 이메일을 매칭시켜 찾아냄). role="corresp" 속성이 있으면
        # 그걸 우선하고, 없으면 이메일 존재 여부로 판단한다.
        has_email = author_el.find(".//tei:email", ns) is not None
        if author_el.get("role") == "corresp":
            corresponding = name
        elif has_email and corresponding is None:
            corresponding = name

    journal = _tei_text(root.find(".//tei:monogr/tei:title", ns))

    volume, page, year = "", "", ""
    for bs in root.findall(".//tei:monogr/tei:imprint/tei:biblScope", ns):
        unit = bs.get("unit", "")
        if unit == "volume":
            volume = _tei_text(bs) or bs.get("from", "")
        elif unit == "page":
            frm, to = bs.get("from", ""), bs.get("to", "")
            page = f"{frm}-{to}" if frm and to else (_tei_text(bs) or frm)
    date_el = root.find(".//tei:monogr/tei:imprint/tei:date", ns)
    if date_el is not None:
        m = re.search(r"(19|20)\d{2}", date_el.get("when", "") or _tei_text(date_el))
        if m:
            year = m.group(0)

    doi = ""
    for idno in root.findall(".//tei:idno", ns):
        if idno.get("type", "").upper() == "DOI":
            doi = _tei_text(idno)
            break

    if not title or not authors:
        return None

    return {
        "title": title, "authors": authors, "journal": journal,
        "volume": volume, "page": page, "year": year, "doi": doi,
        "corresponding": corresponding,
    }


def fetch_grobid_header(filepath):
    """GROBID의 processHeaderDocument API를 호출해 헤더를 추출한다. 서버가
    꺼져 있거나 이 파일 처리가 실패해도 예외를 던지지 않고 None을 반환해서,
    호출부가 항상 기존 정규식 폴백으로 안전하게 넘어갈 수 있게 한다."""
    if not _grobid_enabled():
        return None
    try:
        with open(filepath, "rb") as f:
            resp = requests.post(
                f"{GROBID_URL}/api/processHeaderDocument",
                files={"input": (os.path.basename(filepath), f, "application/pdf")},
                # consolidateHeader=1: GROBID 쪽에 biblio-glutton/CrossRef 연동이
                # 설정돼 있으면 헤더를 한 번 더 검증/보강해준다(설정 안 돼있어도
                # 무시되고 그냥 원래 추출 결과가 옴 - 안전한 기본값).
                data={"consolidateHeader": "1"},
                timeout=GROBID_TIMEOUT,
            )
    except requests.exceptions.RequestException:
        return None
    if not (200 <= resp.status_code < 300) or not resp.content:
        return None
    return parse_grobid_header_xml(resp.content)


def metadata_from_grobid(g):
    """fetch_grobid_header()가 반환한 dict를 결과 행 형태로 변환한다."""
    authors = ", ".join(g["authors"])
    first_author = g["authors"][0] if g["authors"] else "확인필요"
    return {
        "저자": authors or "확인필요",
        "년도": g.get("year") or "확인필요",
        "제목": g.get("title") or "확인필요",
        "저널 이름": g.get("journal") or "확인필요",
        "볼륨": g.get("volume") or "확인필요",
        "추가 정보(페이지)": g.get("page") or "확인필요",
        "1저자": first_author,
        "DOI": g.get("doi") or "확인필요",
    }
# =====================================================================

def clean_filename_to_title(filename):
    name = os.path.splitext(filename)[0]
    name = re.sub(r"^\d+-", "", name)
    name = name.replace("_", " ")
    name = re.sub(r"\s+", " ", name).strip()
    return name

def parse_wiley_entomological_filename(filename):
    stem = os.path.splitext(filename)[0]
    parts = stem.split("_-_")
    if len(parts) >= 4 and parts[0].replace("_", " ").strip() == "Entomological Research":
        year = parts[1].strip()
        author = parts[2].replace("_", " ").strip()
        title = "_-_".join(parts[3:]).replace("_", " ")
        title = re.sub(r"\s+", " ", title).strip()
        return title, author, year
    return None

# 첫 페이지 상단에 제목 바로 위, 독립된 한 줄로 인쇄되는 문서 유형 배너
# (예: "RESEARCH ARTICLE", "REVIEW", "CASE REPORT")를 인식한다. 실제로
# 한국균학회지 등 여러 논문에서 "RESEARCH ARTICLE"이 이런 식으로 인쇄된
# 것을 확인했다. 지금까지 "참고문헌 종류"는 항상 "Journal Article"로
# 고정돼 있었는데, 이 배너로 좀 더 구체적인 유형을 구분할 수 있다.
# 다만 배너가 없으면 억지로 추측하지 않고 기존 기본값을 그대로 둔다.
DOC_TYPE_LABELS = (
    "Research Article", "Review Article", "Mini Review", "Review",
    "Case Report", "Short Communication", "Communication",
    "Technical Note", "Brief Report", "Editorial", "Letter",
    "Erratum", "Correction", "Corrigendum", "Addendum",
    "Commentary", "Perspective", "Conference Paper",
)
DOC_TYPE_LINE_REGEX = re.compile(
    r"^[ \t]*(?:" + "|".join(re.escape(t) for t in DOC_TYPE_LABELS) + r")[ \t]*$",
    re.IGNORECASE | re.MULTILINE,
)

def detect_document_type(first_page_text):
    m = DOC_TYPE_LINE_REGEX.search(first_page_text[:800])
    if m:
        return m.group(0).strip().title()
    return None

def detect_journal_fallback(text):
    head = re.sub(r"\s+", " ", text[:400])
    if re.search(r"korean journal of mycology|kjmycology", head, re.IGNORECASE):
        return "The Korean Journal of Mycology (한국균학회지)"
    # "Microbiological Society of Korea" 저작권 문구는 완전 2단(2-column)
    # 레이아웃 논문(예: 2019년 이후 Springer 판형)에서 본문 중간, 실측
    # 5,000자 넘는 지점까지 밀려나 있는 경우가 있다(2단 레이아웃 때문에
    # extract_text가 두 컬럼을 뒤섞어 순서를 왜곡시킴). 배너의 "Journal of
    # Microbiology" 자체는 이미 첫 400자(head) 안에서 특이하게 매칭되는
    # 신뢰도 높은 신호이므로, corroboration 창을 넉넉히 넓혀서 이런
    # 레이아웃에서도 저널명을 놓치지 않게 한다.
    if "Journal of Microbiology" in head and "Microbiological Society of Korea" in text[:8000]:
        return "Journal of Microbiology"
    # 1995년 전후의 옛 배너는 "Jour. Microbiol. March 1995 p. 16-20"처럼
    # 축약형으로 인쇄되어 있고, "Microbiological Society of Korea" 문구가
    # 아예 없는 경우가 있다(OCR 텍스트로 실측 확인). 배너 맨 앞부분의
    # 축약 표기만으로 판별한다.
    if re.match(r"\s*Jour\.?\s*Microbiol\.?\b", text[:60], re.IGNORECASE):
        return "Journal of Microbiology"
    if re.search(r"entomological\s*research", head, re.IGNORECASE):
        return "Entomological Research"
    if re.search(r"mycobiology", head, re.IGNORECASE):
        return "Mycobiology"
    return "확인필요"

# 같은 저널인데도 이름이 여러 형태로 섞여 들어오는 경우를 하나로 통일한다.
# - 텍스트 기반 폴백(detect_journal_fallback)은 항상 "Journal of Microbiology"로
#   반환하지만, CrossRef의 container-title 필드는 옛날 표기("The Journal of
#   Microbiology")를 그대로 돌려주는 레코드가 섞여 있어 소스에 따라 값이
#   갈리는 문제가 있었다. 출처와 무관하게 최종적으로 한 번 더 정규화한다.
JOURNAL_NAME_ALIASES = {
    "the journal of microbiology": "Journal of Microbiology",
    "journal of microbiology (seoul, korea)": "Journal of Microbiology",
    # 한국균학회지: "(한국균학회지)"가 붙은 것과 안 붙은 것이 섞여 들어온다.
    # 괄호 포함된 쪽을 정식 표기로 보고 통일한다.
    "the korean journal of mycology": "The Korean Journal of Mycology (한국균학회지)",
}

def normalize_journal_name(name):
    if not name or name == "확인필요":
        return name
    key = re.sub(r"\s+", " ", name).strip().lower()
    return JOURNAL_NAME_ALIASES.get(key, name)

YEAR_REGEX = re.compile(r"\b((?:19|20)\d{2})\b")

# eISSN/pISSN 번호(예: "eISSN 1976-3794")의 앞 네 자리가 연도처럼 보여서
# 오탐되는 문제가 있었다(예: Journal of Microbiology). ISSN 번호는
# "ISSN 1234-5678"이나 "ISSN: 1234-5678" 형태로 등장하므로, 그 구간에
# 있는 4자리 숫자는 연도 후보에서 제외한다.
ISSN_SPAN_REGEX = re.compile(r"(?:[pe]?ISSN)\s*[:\uFF1A]?\s*\d{4}\s*[–\-]\s*\d{4}", re.IGNORECASE)

# 더 신뢰도 높은 위치에서 우선적으로 연도를 찾는다:
# 1) 저널 배너의 "Month Year Vol." 형식 (예: "January 2025 Vol 63 No 1") - 실제 출판 시점
# 2) "Accepted:" 뒤에 오는 날짜 - 투고일(Received)과 게재일(Accepted)의 연도가
#    다른 경우(흔함, 예: 2002년 12월 투고 / 2003년 3월 게재)가 있어서, 실제
#    출판 연도에 더 가까운 Accepted만 사용한다. Received/Revised는 쓰지 않는다
#    (텍스트 순서상 Received가 먼저 나와서, 구분 없이 찾으면 투고연도를
#    출판연도로 잘못 채택하는 사고가 났었다).
MONTH_YEAR_VOL_REGEX = re.compile(
    r"(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+"
    r"((?:19|20)\d{2})\s+Vol", re.IGNORECASE
)
ACCEPTED_YEAR_REGEX = re.compile(
    r"Accepted\s*[:\uFF1A]?\s*[^\n,]*?,\s*((?:19|20)\d{2})", re.IGNORECASE
)

# The Korean Journal of Mycology(한국균학회지)는 2단 레이아웃 오른쪽 컬럼에
# "Kor. J. Mycol. 2016 December, 44(4): 300-306" 형식의 배너를 인쇄해두는데,
# pdfplumber가 2단을 뒤섞어 뽑아내는 탓에 이 배너가 페이지 맨 앞이 아니라
# 본문 중간(실측: 1400~2000번째 글자 부근)에 등장한다. 기존 head(600자)/
# wide_head(1500자) 창을 벗어나는 경우가 있어 연도/볼륨/페이지 채움률이
# 낮았다. "Kor. J. Mycol." 리터럴 문자열이 워낙 특이해서 오탐 위험이 낮으므로,
# 창 크기를 넓게 잡고 이 배너 하나로 연도·볼륨·페이지를 한 번에 뽑는다.
KOR_J_MYCOL_REGEX = re.compile(
    r"Kor\.?\s*J\.?\s*Mycol\.?\s*((?:19|20)\d{2})\s+[A-Za-z]+,?\s*"
    r"(\d{1,3})\s*\(\s*\d+\s*\)\s*[:\uFF1A]\s*(\d+)\s*[–\-]\s*(\d+)",
    re.IGNORECASE
)

def _kor_j_mycol_match(text):
    return KOR_J_MYCOL_REGEX.search(text[:3000])

def _issn_excluded_spans(text):
    return [m.span() for m in ISSN_SPAN_REGEX.finditer(text)]

def _year_candidates_excluding_issn(text):
    excluded = _issn_excluded_spans(text)

    def is_excluded(pos):
        return any(start <= pos < end for start, end in excluded)

    candidates = []
    for m in YEAR_REGEX.finditer(text):
        if is_excluded(m.start()):
            continue
        candidates.append(m.group(1))
    return candidates

def guess_year_fallback(text):
    """연도를 (값, 신뢰도)로 반환한다. 신뢰도는 "high"(전용 배너 매칭),
    "medium"(Accepted 날짜), "low"(그 외 - 본문 앞부분의 4자리 숫자 중
    ISSN이 아닌 첫 번째 것을 그냥 채택하는 것이라 오탐 위험이 상대적으로
    높음) 중 하나이며, 못 찾으면 ("확인필요", None)을 반환한다."""
    # 우선순위 1, 2에 쓰는 구간은 좀 더 넉넉하게 잡는다. 2단 레이아웃
    # PDF는 pdfplumber가 컬럼을 뒤섞어 텍스트를 뽑아내는 경우가 있어서,
    # "Received:" 같은 표지 정보가 앞부분 600자를 넘어가기도 한다
    # (예: Journal of Microbiology).
    wide_head = text[:1500]
    head = text[:600]

    # 우선순위 0: "Kor. J. Mycol. YYYY Month, VV(N): PP-PP" 배너
    # (한국균학회지 전용, 2단 레이아웃 때문에 앞의 다른 우선순위 창보다
    # 훨씬 뒤에 등장할 수 있어 넓은 창에서 먼저 확인한다)
    m0 = _kor_j_mycol_match(text)
    if m0:
        return m0.group(1), "high"

    # 우선순위 1: "Month Year Vol." 배너 (실제 출판 시점을 가장 정확히 반영)
    m = MONTH_YEAR_VOL_REGEX.search(wide_head)
    if m:
        return m.group(1), "high"

    # 우선순위 2: Accepted 날짜 (Received는 투고연도라 다를 수 있어 제외)
    m = ACCEPTED_YEAR_REGEX.search(wide_head)
    if m:
        return m.group(1), "medium"

    # 우선순위 3: 그 외 4자리 연도 숫자 중 ISSN 번호에 속하지 않는 것
    # (본문 뒤쪽에서 엉뚱한 숫자를 주워오는 걸 막기 위해 이 단계만
    # 앞부분 600자로 범위를 좁게 유지한다). 전용 배너나 Accepted 날짜처럼
    # 명시적인 근거가 아니라 "그냥 첫 번째로 보이는 4자리 숫자"이므로
    # 신뢰도를 낮게 매긴다.
    candidates = _year_candidates_excluding_issn(head)
    if candidates:
        return candidates[0], "low"

    # 우선순위 4: 년도 숫자 중간에 공백이 낀 옛 판형 배너 복구
    # (실측: "Mycobiology 29(4): 85-89 (200 1)" - "2001"이 "200 1"로 인쇄됨.
    # 페이지 잘림(_fix_truncated_page_end)과 같은 원인으로 추정되는, 같은
    # 계열의 옛 판형 결함이다. 표준 4자리 연속 패턴이 전혀 매칭되지
    # 않으므로 우선순위 3까지 전부 실패한 뒤에만 시도하고, 괄호로 감싸인
    # 형태로만 한정해서 본문 중 무관한 숫자를 잘못 붙잡지 않게 한다.)
    m4 = re.search(r"\((?:19|20)\d[ \u00a0](?:\d{1,2})\)", wide_head)
    if m4:
        repaired = re.sub(r"[ \u00a0]", "", m4.group(0)).strip("()")
        return repaired, "low"

    return "확인필요", None

VOLUME_PAGE_REGEX = re.compile(
    r"[Vv]ol\.?\s*(\d+)[,\s]*(?:No\.?\s*\d+[,\s]*)?(?:pp\.?\s*)?(\d+)\s*[–\-]\s*(\d+)"
)
COMPACT_VOLUME_PAGE_REGEX = re.compile(r"\b(\d{1,3})\s*\(\s*\d+\s*\)\s*[:\uFF1A]\s*(\d+)\s*[–\-]\s*(\d+)")
# Wiley 계열 저널(예: Entomological Research)은 콜론 없이 "38 (2008) 250–256"
# 형식으로 본문 맨 위 배너에 볼륨(연도)페이지를 적어둔다. CrossRef에 페이지
# 정보가 아예 없는 경우가 많아서, 이 배너를 직접 읽는 폴백을 추가한다.
# 일부 특집 섹션(Special Feature)/특별호는 "S66-S70", "A2-A10"처럼
# 페이지 번호 앞에 알파벳이 붙는다(예: "Entomological Research 38 (2008) S66–S70").
# 순수 숫자만 인식하던 기존 정규식으로는 이런 형식을 놓쳤다.
NO_COLON_VOLUME_PAGE_REGEX = re.compile(r"\b(\d{1,3})\s*\(\s*(?:19|20)\d{2}\s*\)\s*([A-Za-z]?\d+)\s*[–\-]\s*([A-Za-z]?\d+)")
# The Journal of Microbiology(구 한국미생물학회지) 등 일부 저널은
# "...March 2003, p.16-21 Vol. 41, No. 1"처럼 페이지 범위가 볼륨보다
# 먼저 나온다. 그룹 순서가 반대라서(페이지, 페이지, 볼륨) 위 정규식들과는
# 별도로 처리해야 한다. 1999년 전후 논문은 페이지 범위와 "Vol." 사이에
# "Copyright (c) 1999, The Microbiological Society of Korea" 같은 줄이
# 끼어 있는 경우가 있어서(OCR 텍스트에서 실측), 그 정도 길이는 건너뛸 수
# 있도록 허용 구간을 넉넉히 잡는다(단, 무한정 멀리 있는 엉뚱한 "Vol."을
# 잘못 집지 않도록 상한을 둔다).
PAGE_THEN_VOLUME_REGEX = re.compile(r"p\.?\s*(\d+)\s*[–\-]\s*(\d+)[\s\S]{0,150}?Vol\.?\s*(\d+)", re.IGNORECASE)

# 1995년 전후의 옛 배너("Jour. Microbiol. March 1995 p. 16-20")는 아예
# 같은 줄/근처에 볼륨 정보가 없다 - 원본 지면에서 "Vol. 33, No. 1"이
# 우측 상단에 별도로 인쇄되어 있어서, 텍스트 추출(OCR 포함) 순서상
# 훨씬 뒤쪽(본문 안쪽)에서 등장한다. 그래서 페이지 범위와 볼륨을
# 각각 다른 폭의 창에서 따로 찾은 뒤 조합한다.
PAGE_ONLY_NEAR_TOP_REGEX = re.compile(r"^.{0,80}?p\.?\s*(\d+)\s*[–\-]\s*(\d+)\b", re.IGNORECASE)
STANDALONE_VOL_REGEX = re.compile(r"\bVol\.?\s*(\d{1,3})\b", re.IGNORECASE)

def _fix_truncated_page_end(text, end_pos, p1, p2):
    """일부 옛 판형 PDF(실측: Mycobiology 볼륨 28~29, 2000~2001년 판형)에서
    끝 페이지 번호 중간에 원인 불명의 공백이 끼어들어("138"이 "1 38"처럼
    추출되는 등) 정규식이 끝 페이지의 앞자리만 캡처하는 사고가 있었다
    (예: "130-138"이 "130-1"로 잘림, 10건 확인 - 그 중 절반은 다른 신뢰도
    표시조차 없이 조용히 발생).

    끝 페이지가 시작 페이지보다 작을 때만(=명백히 비정상일 때만) 작동하며,
    매칭이 끝난 바로 그 지점에서 공백 하나를 사이에 두고 숫자가 더
    이어지는지 확인해서, 이어붙였을 때 비로소 페이지 범위가 말이 되면
    그 값을 채택한다. 정상적으로 끝난 페이지 번호는 건드리지 않는다.
    다만 이건 실제 원본 PDF로 확정 검증된 원인은 아니라서(추정), 보정이
    적용된 값은 신뢰도를 낮게 매겨 사람이 한 번 더 볼 수 있게 한다."""
    try:
        p1_int, p2_int = int(p1), int(p2)
    except ValueError:
        return p2, False
    if p2_int >= p1_int:
        return p2, False
    m = re.match(r"[ \u00a0]?(\d{1,2})\b", text[end_pos:end_pos + 4])
    if not m:
        return p2, False
    extended = p2 + m.group(1)
    if int(extended) > p1_int:
        return extended, True
    return p2, False

def guess_volume_page_fallback(text):
    """(볼륨, 페이지, 신뢰도)를 반환한다. 신뢰도는 "high"(전용 배너 또는
    Vol(No):page 형태의 구조화된 패턴 매칭), "medium"(페이지가 볼륨보다
    먼저 나오는 옛 배너), "low"(페이지와 볼륨을 서로 다른 구간에서 각각
    따로 찾아 조합한 것 - 실제로 관계없는 두 숫자가 우연히 조합될 위험이
    가장 큼) 중 하나이며, 못 찾으면 ("확인필요", "확인필요", None)이다."""
    head = text[:600]

    # 우선순위 0: "Kor. J. Mycol." 배너 (연도 폴백과 동일한 근거, 위 참고)
    m0 = _kor_j_mycol_match(text)
    if m0:
        vol, p1, p2 = m0.group(2), m0.group(3), m0.group(4)
        p2_fixed, was_fixed = _fix_truncated_page_end(text, m0.end(), p1, p2)
        return vol, f"{p1}-{p2_fixed}", ("low" if was_fixed else "high")

    m = (VOLUME_PAGE_REGEX.search(head)
         or COMPACT_VOLUME_PAGE_REGEX.search(head)
         or NO_COLON_VOLUME_PAGE_REGEX.search(head))
    if m:
        vol, p1, p2 = m.group(1), m.group(2), m.group(3)
        p2_fixed, was_fixed = _fix_truncated_page_end(text, m.end(), p1, p2)
        return vol, f"{p1}-{p2_fixed}", ("low" if was_fixed else "high")
    m2 = PAGE_THEN_VOLUME_REGEX.search(head)
    if m2:
        p1, p2, vol = m2.group(1), m2.group(2), m2.group(3)
        # 이 패턴은 "p.130-138 ... Vol.29"처럼 볼륨이 페이지 뒤에 있어서,
        # 매칭이 끝나는 지점(m2.end())이 이미 볼륨 숫자까지 지나친 뒤다.
        # 페이지 끝 번호 보정은 볼륨이 시작되기 전, 페이지 범위 자체가
        # 끝나는 지점(그룹 2 직후)을 기준으로 해야 한다.
        page_end_pos = m2.end(2)
        p2_fixed, was_fixed = _fix_truncated_page_end(text, page_end_pos, p1, p2)
        return vol, f"{p1}-{p2_fixed}", ("low" if was_fixed else "medium")
    # 페이지 범위는 맨 위 배너 근처(첫 80자 이내)에서, 볼륨은 첫 페이지
    # 전체(최대 2500자)에서 각각 따로 찾는다. 페이지 범위 검색을 문서
    # 맨 앞으로 좁혀서, 본문 중간의 무관한 숫자 범위(예: "pH 4-5")를
    # 오인하는 것을 막는다. 다만 두 값을 서로 다른(넓은) 구간에서 독립적으로
    # 찾아 조합하는 방식이라, 실제로는 서로 무관한 볼륨/페이지가 우연히
    # 짝지어질 위험이 다른 경로보다 높다 - 신뢰도를 낮게 매긴다.
    m3 = PAGE_ONLY_NEAR_TOP_REGEX.search(text[:150])
    if m3:
        p1, p2 = m3.group(1), m3.group(2)
        vol_m = STANDALONE_VOL_REGEX.search(text[:2500])
        if vol_m:
            p2_fixed, _ = _fix_truncated_page_end(text, m3.end(2), p1, p2)
            return vol_m.group(1), f"{p1}-{p2_fixed}", "low"
    return "확인필요", "확인필요", None

AFFIL_HINTS = ("University", "Institute", "Department", "College", "Center",
               "Centre", "Laboratory", "Faculty", "Academy", "Hospital",
               "Corporation", "Office", "Program", "Division", "School",
               "Research", "Administration")
STOP_WORDS = ("Department", "Division", "Institute", "College", "University",
              "School", "Laboratory", "Corresponding", "*Email", "Email:",
              "Keywords", "KEYWORDS", "Key words", "ABSTRACT", "Correspondence",
              "To cite this article", "Received", "Accepted")
SUPERSCRIPT_RUN = re.compile(r"[a-z](?:,[a-z]){1,}")
DIGIT_SUPERSCRIPT = re.compile(r"[A-Za-z]\d{1,2}(?=[,\s]|$)")

# 교신저자/저자 이름 후보로 채택하면 안 되는 흔한 비-이름 단어들.
# 표(Table)의 컬럼 헤더(예: "Sex **")나 섹션 제목 단어가 마커(*, †, ✉ 등)
# 바로 앞/뒤에 우연히 붙어 이름으로 오인되는 사고가 실제로 확인되었다
# (예1: Table 1의 "Sex **" -> "Sex"가 교신저자로 잘못 채택됨.
#  예2: 다른 논문에서는 "Isolates"가 같은 방식으로 잘못 채택됨).
# 두 사고 모두 "1단어짜리 후보"였다는 공통점이 있어서, 이 블록리스트로
# _extract_name_before_marker와 _valid_name 양쪽에서 공통으로 걸러낸다.
NON_NAME_CANDIDATE_WORDS = {
    "abstract", "keywords", "introduction", "correspondence", "corresponding",
    "author", "authors", "received", "accepted", "sex", "age", "male",
    "female", "total", "number", "mean", "table", "figure", "isolates",
    "species", "group", "groups", "control", "result", "results", "method",
    "methods", "materials", "discussion", "conclusion", "note", "notes",
}

# 일부 PDF는 임베디드 폰트에 ToUnicode 맵이 없어서, pdfplumber가 글자를
# 진짜 문자 대신 "(cid:123)" 같은 내부 글리프 ID로 뽑아낸다. 이런 텍스트는
# 겉보기엔 "텍스트가 있음"이라 OCR-불가 판정(글자 수 < 150)을 피해가지만,
# 실제로는 사람이 읽을 수 없는 깨진 텍스트다. 정규식이 우연히 이 잡음
# 안에서 콤마·숫자 패턴을 "위첨자"로 오인해서, "저자" 같은 필드에
# "(cid:...)" 잔해가 그대로 새어 들어가는 사고가 있었다(예시 파일에서 확인).
CID_GARBLE_REGEX = re.compile(r"\(cid:\d+\)")
# 위 정규식은 문서 단위 폰트 깨짐 판정(is_font_garbled)용으로, 숫자가 있는
# 온전한 "(cid:19)" 형태만 센다. 반면 저자/교신저자 필드에 새어 들어온
# 잔해는 _clean_author_line의 숫자 제거 단계에서 "19" 부분이 먼저 지워져
# "(cid:)" 껍데기만 남는 경우가 실측 확인됐다(실측: "Chang Soo Lee(cid:)").
# 이런 사후 잔해까지 잡아야 새는 일이 없으므로, 필드 잔해 감지 전용으로는
# 숫자를 선택 사항(\d*)으로 둔 넓은 버전을 따로 쓴다.
CID_GARBLE_ANY_REGEX = re.compile(r"\(cid:\d*\)")
# 폰트에 매핑이 아예 없으면 "(cid:19)" 형태로 남지만, 매핑은 있는데 그
# 매핑이 가리키는 글자를 표현할 수 없을 때는 pdfplumber가 유니코드
# 대체문자(U+FFFD, 화면에 " "로 보임)를 채워 넣는다. 실측: 저자 필드가
# "Jeong su moon  min young an  mal geum park  ,  jong soo lee"처럼
# 통째로 대체문자 범벅이 됐는데도, 기존 "(cid:...)" 패턴 검사에는 안
# 걸려서 비고 표시 없이 그대로 나갈 뻔했다. 같은 부류의 문제이므로 함께
# 감지한다.
REPLACEMENT_CHAR_REGEX = re.compile("\ufffd")

# =====================================================================
# [추가됨] 영문 우선 처리 (KJM 등 한글/영문을 병기하는 저널 대응)
# =====================================================================
# 한국균학회지(KJM) 같은 저널은 논문 자체가 한글로 쓰였거나, 한글/영문
# 저자명·제목을 나란히 병기하는 경우가 흔하다. GROBID나 정규식 폴백이
# 한글 텍스트 블록을 그대로 제목/저자로 채택해버리면(실측: "가강현, 정숙주
# • 박현, Kang-Hyeon Ka, Sukju Jeong, Hyun Park"처럼 한글/영문이 뒤섞임),
# 참고문헌 표기 일관성이 깨진다. 업로드된 파일들의 파일명 자체가 거의 항상
# 이미 영문 제목으로 정리되어 있음을 확인했으므로, 한글이 섞인 결과가
# 나오면 영문 소스(파일명/CrossRef)로 대체한다.
HANGUL_REGEX = re.compile(r"[\uac00-\ud7a3]")

def _contains_hangul(text):
    return bool(HANGUL_REGEX.search(text or ""))

def _strip_hangul_author_tokens(authors_str):
    """저자 문자열에서 한글 이름 토큰만 제거하고 영문 표기만 남긴다
    (예: "가강현, Kang-Hyeon Ka" -> "Kang-Hyeon Ka"). 전부 한글이라 영문
    표기가 하나도 안 남으면 빈 문자열을 반환해서 호출부가 이상 신호로
    판단할 수 있게 한다."""
    parts = [p.strip(" ,") for p in re.split(r",|\band\b", authors_str, flags=re.IGNORECASE)]
    kept = [p for p in parts if p and not _contains_hangul(p)]
    return ", ".join(kept)
# =====================================================================

# CrossRef 공개 API의 폴라이트 풀 한도(초당 약 50건)를 넘지 않도록, 여러
# 스레드가 동시에 CrossRef에 요청을 보내는 것을 제한한다. MAX_WORKERS(8)가
# 전부 동시에 CrossRef를 두드리면(직접조회 + 제목검색 두 종류 API를 같이
# 쓰므로 순간 부하가 더 큼) 자체적으로 429를 유발해 "DOI는 있는데 CrossRef만
# 실패"하는 사례를 늘릴 수 있다는 지적이 있어 추가했다.
_CROSSREF_SEMAPHORE = threading.Semaphore(4)

def is_font_garbled(text, sample_len=1500, min_hits=5):
    sample = text[:sample_len]
    return len(CID_GARBLE_REGEX.findall(sample)) >= min_hits

# 저자 후보 검증에서, 구조적으로는 "대문자로 시작하는 단어 2~4개"라는
# 이름 패턴을 만족하지만 실제로는 소속기관/회사/국가명인 경우가 있다.
# 실측 사고: "Date Processing Plant France."가 단어 3개 다 대문자로
# 시작한다는 이유만으로 이름 검증을 통과해 저자로 잘못 채택됨. 구조적
# 정규식만으로는 사람 이름과 기관명을 구분할 수 없으므로, 흔한 기관/회사
# 관련 단어와 국가명을 포함하면 이름 후보에서 제외하는 의미 기반 필터를
# 추가한다.
_NON_NAME_WORDS = (
    "Plant", "Plants", "Company", "Corporation", "Corp", "Institute", "Institution",
    "University", "College", "Laboratory", "Laboratories", "Center", "Centre",
    "Department", "Division", "Faculty", "School", "Academy", "Hospital",
    "Office", "Program", "Administration", "Association", "Society",
    "Foundation", "Council", "Ministry", "Agency", "Bureau", "Committee",
    "Network", "Service", "Services", "Trust", "Board", "Group", "Industries",
    "Industry", "Factory", "Farm", "Farms", "Processing", "Station", "Unit",
    "Branch", "Enterprise", "Enterprises", "Cooperative", "Union", "Federation",
    "Ltd", "Inc", "LLC", "PLC", "GmbH",
    # 소속 표기 맨 끝에 국가명이 붙는 경우("...Plant, France."처럼)가 이름
    # 으로 오인되는 사고를 막기 위한 국가/지역명 일부(전수는 아니며, 흔히
    # 등장하는 것 위주)
    "France", "Germany", "Korea", "China", "Japan", "USA", "England",
    "Britain", "Italy", "Spain", "India", "Brazil", "Canada", "Australia",
    "Netherlands", "Belgium", "Switzerland", "Austria", "Poland", "Russia",
    "Mexico", "Egypt", "Thailand", "Vietnam", "Indonesia", "Malaysia",
    "Philippines", "Taiwan", "Turkey", "Iran", "Israel", "Nigeria", "Kenya",
    "Chile", "Argentina", "Sweden", "Norway", "Denmark", "Finland", "Ireland",
    "Portugal", "Greece",
)
_NON_NAME_WORD_REGEX = re.compile(
    r"\b(?:" + "|".join(re.escape(w) for w in _NON_NAME_WORDS) + r")\b",
    re.IGNORECASE,
)

def _contains_non_name_word(text):
    """저자 후보 문자열에 기관/회사/국가명이 섞여 있으면 True (=이름 아님)."""
    return bool(_NON_NAME_WORD_REGEX.search(text or ""))

def _looks_like_plausible_name_line(line):
    """추출된 '저자' 후보 줄이 실제 사람 이름처럼 보이는지 검증한다.
    (cid:...) 잔해나, 알파벳 비중이 지나치게 낮은 잡음 줄을 걸러낸다."""
    if not line:
        return False
    if CID_GARBLE_REGEX.search(line):
        return False
    letters = re.findall(r"[A-Za-z]", line)
    if len(letters) < 4:
        return False
    # 이름 줄이라면 대부분 알파벳/공백/콤마/마침표/하이픈으로 구성되어야 한다.
    # 그 외 특수문자(예: $, &, (, ) 등 잡음 기호) 비중이 높으면 깨진 텍스트로 본다.
    noise_chars = re.findall(r"[^A-Za-z0-9 ,.\-*∗†✉‡]", line)
    if len(noise_chars) > max(3, len(line) * 0.15):
        return False
    if _contains_non_name_word(line):
        return False
    return True

_SINGLE_NAME_TOKEN = re.compile(r"^[A-Z][A-Za-z.\-]*(?:\s+[A-Z][A-Za-z.\-]*){1,3}$")

def _looks_like_author_list(line):
    """구조까지 검증하는 더 엄격한 버전: 콤마/'and'로 나눈 조각들이 각각
    '대문자로 시작하는 단어 2~4개'라는 이름 형태를 갖추고 있는지 확인한다.
    유전자/클론 이름(예: "a13", "apT1")이 우연히 위첨자 패턴(letter+숫자)에
    걸려 저자 줄로 오인되는 사고를 막기 위해 도입했다(실제 1995년 논문
    OCR 테스트에서 확인된 문제)."""
    if not _looks_like_plausible_name_line(line):
        return False
    # 콤마뿐 아니라 " and "도 저자 구분자로 취급해서 나눈다. 콤마 없이
    # "A and B"처럼 정확히 2명만 " and "로만 이어진 경우(콤마가 전혀 없는
    # 경우), 콤마 기준으로만 나누면 전체가 한 덩어리로 남아 "and"가 중간에
    # 낀 문자열이 되어 이름 형태 검증에서 무조건 탈락하는 문제가 있었다
    # (실제 2인 저자 논문에서 확인). 콤마+and가 섞인 3인 이상 목록에서도
    # 결과가 동일하므로(예: "A, B and C" -> ["A","B","C"]) 부작용은 없다.
    parts = [re.sub(r"^\s*and\s+", "", p.strip(), flags=re.IGNORECASE).strip(" *∗†✉‡!?'\"")
             for p in re.split(r",|\s+and\s+", line, flags=re.IGNORECASE)]
    parts = [p for p in parts if p]
    if not parts:
        return False
    good = sum(1 for p in parts if _SINGLE_NAME_TOKEN.match(p))
    return good >= max(1, len(parts) - 1)

def find_author_line_by_superscript(lines):
    for idx, line in enumerate(lines[:20]):
        letter_hits = len(SUPERSCRIPT_RUN.findall(line))
        digit_hits = len(DIGIT_SUPERSCRIPT.findall(line))
        if letter_hits >= 2 or digit_hits >= 2:
            combined = line
            j = idx + 1
            while j < len(lines) and j < idx + 3 and (
                len(SUPERSCRIPT_RUN.findall(lines[j])) >= 1 or len(DIGIT_SUPERSCRIPT.findall(lines[j])) >= 1
            ):
                combined += " " + lines[j]
                j += 1
            return combined
    return None

# 논문 구조는 거의 항상 "제목(여러 줄로 감길 수 있음) → 저자 → 소속" 순서를
# 따른다. 제목이 여러 줄로 감기면 일반 후보 스캔이 제목의 이어지는 줄을
# 저자 줄로 오인하기 쉬운데(실제로 1990년대 논문 OCR 테스트에서 확인),
# 소속 줄(Department/University 등)은 훨씬 더 확실하게 식별되므로, 그
# 바로 위 줄(들)을 저자로 보는 편이 안전하다.
def find_author_line_before_affiliation(lines):
    affil_idx = None
    for idx, line in enumerate(lines[:25]):
        if any(h in line for h in AFFIL_HINTS):
            affil_idx = idx
            break
    if affil_idx is None or affil_idx == 0:
        return None

    first = lines[affil_idx - 1].strip()
    if not first or not re.search(r"[A-Z][A-Za-z.\-]+(?:\s+[A-Z][A-Za-z.\-]+)+", first) or _contains_non_name_word(first):
        return None
    collected = [first]

    # 소속 줄 바로 위 한 줄은 항상 채택한다. 그보다 더 위로 거슬러 올라가는
    # 것은, "그 위 줄" 자체가 쉼표로 끝나 다음 줄로 이어짐을 스스로 나타낼
    # 때만 허용한다(마지막에 채택한 줄의 끝이 아니라, 한 칸 더 위 줄의
    # 끝을 기준으로 판단해야, 제목의 마지막 줄이 우연히 이름처럼 보여도
    # 잘못 붙지 않는다 - 실제로 1995년 논문 OCR에서 이 문제가 확인됨).
    # 콤마뿐 아니라 " and"로 끝나는 경우도 이어짐 신호로 인정한다(실측:
    # "...Moo-Yong Eun and\nSeong-Joo Go"처럼 저자 목록 마지막이 콤마 없이
    # "and"로 줄바꿈되는 경우, 콤마만 보면 이어붙이지 못해 마지막 저자
    # 한 명만 채택되는 사고가 있었다).
    j = affil_idx - 2
    while j >= 0 and len(collected) < 3:
        candidate = lines[j].strip()
        if not candidate or not re.search(r",\s*$|\band\s*$", candidate, re.IGNORECASE):
            break
        if not re.search(r"[A-Z][A-Za-z.\-]+(?:\s+[A-Z][A-Za-z.\-]+)+", candidate) or _contains_non_name_word(candidate):
            break
        collected.insert(0, candidate)
        j -= 1
    return " ".join(collected)

def _clean_author_line(line):
    line = re.sub(r"\d+(,\d+)*[\*∗]?", "", line)
    line = re.sub(r"(?<=[A-Za-z])[a-z](?:,[a-z]){1,}(?=\s|,|$)", "", line)
    line = re.sub(r"[\*∗]", " ", line)
    line = re.sub(r"\s+,", ",", line)
    line = re.sub(r"\s+", " ", line).strip(" ,")
    # 저자 목록 맨 끝에 의미 없는 홑따옴표/굽은따옴표가 잔재로 남는 경우가
    # 있다(실측: "...Kyung-Ho Yun'" - 마지막 저자 이름 뒤에 따옴표 하나가
    # 그대로 붙음). 실제 성씨가 따옴표로 끝나는 경우는 사실상 없으므로,
    # 문자열 맨 끝의 따옴표류만 순수하게 제거한다("O'Brien"처럼 이름 중간에
    # 오는 정상적인 아포스트로피는 끝이 아니라서 영향받지 않는다).
    line = re.sub(r"['’`´]+\s*$", "", line).strip()
    return line

SINGLE_NAME_LINE = re.compile(r"[A-Z][A-Za-z.\-]+(?:\s+[A-Z][A-Za-z.\-]+){1,3}")

def _is_probable_single_author_line(line):
    """콤마나 ' and '가 없는 단독저자 줄(예: 'Kwan Hee Yoo*')도 저자 후보로
    인정하기 위한 보조 판별. 각주 기호(*, ∗, †, ✉, ‡)와 위첨자 숫자,
    콤마를 제거한 뒤 남는 문자열이 '대문자로 시작하는 단어 2~4개'로만
    이루어져 있으면(다른 글자가 섞여 있지 않으면) 단독 저자명으로 간주한다."""
    stripped = re.sub(r"[\*∗†✉‡\d,]", "", line).strip()
    return bool(re.fullmatch(SINGLE_NAME_LINE, stripped)) and not _contains_non_name_word(stripped)

def extract_authors_fallback(text, title=""):
    """저자 목록을 신뢰도 등급과 함께 반환한다: (저자문자열, 신뢰도).
    신뢰도는 "high"(위첨자 마커로 확인), "medium"(소속기관 줄 바로 위에서
    확인), "low"(그 외 - 본문을 훑다가 '이름처럼 생긴 줄'을 채택한 것으로,
    실제로 본문 문장이 잘못 채택된 사고가 있었던 가장 불확실한 경로)
    셋 중 하나이며, 확인 자체가 안 되면 ("확인필요", None)을 반환한다.
    호출부(process_pdf)는 "low"일 때 비고에 수동 확인 권장 표시를 남긴다."""
    lines = [l.strip() for l in text.splitlines() if l.strip()]

    sup_line = find_author_line_by_superscript(lines)
    if sup_line:
        cleaned = _clean_author_line(sup_line)
        if cleaned and _looks_like_author_list(cleaned):
            return cleaned, "high"

    affil_line = find_author_line_before_affiliation(lines)
    if affil_line:
        cleaned = _clean_author_line(affil_line)
        if cleaned and _looks_like_author_list(cleaned):
            return cleaned, "medium"

    title_words = set(w.lower() for w in re.findall(r"[A-Za-z]{4,}", title))
    candidates = []
    # 저널 배너/저작권 표기 줄(예: "The Journal of Microbiology, December 1999,
    # p.185-192", "Copyright (c) 1999, The Microbiological Society of Korea")은
    # 대문자 단어가 많고 콤마도 있어서 이름 줄처럼 오인되기 쉽다. 실제로
    # 이런 줄이 저자 후보로 잘못 채택되는 사고가 확인되어(1999년 논문
    # OCR 테스트), 명시적으로 제외한다.
    BANNER_LINE_REGEX = re.compile(
        r"copyright|journal of microbiology|microbiological society|"
        r"\bvol\.?\s*\d|\bp\.?\s*\d+\s*[–\-]\s*\d+|"
        r"\b(?:january|february|march|april|may|june|july|august|september|october|november|december)\s+(?:19|20)\d{2}",
        re.IGNORECASE,
    )
    for line in lines:
        if any(sw in line for sw in STOP_WORDS):
            break
        if BANNER_LINE_REGEX.search(line):
            continue
        line_words = set(w.lower() for w in re.findall(r"[A-Za-z]{4,}", line))
        if title_words and line_words:
            overlap = len(line_words & title_words) / len(line_words)
            if overlap > 0.5:
                continue
        has_name_pattern = bool(re.search(r"[A-Z][a-z]+(?:-[A-Z][a-z]+)?\s+[A-Z][a-z]+", line))
        cap_tokens = re.findall(r"\b[A-Z][A-Za-z]*\b", line)
        looks_like_names = has_name_pattern or len(cap_tokens) >= 3
        # 저자가 1명뿐이면 콤마도 " and "도 없는 줄이 되므로, 그 경우를 위한
        # 보조 조건을 추가한다(단, 그런 줄엔 소속 힌트가 없어야 함).
        single_author_line = _is_probable_single_author_line(line)
        if looks_like_names and (("," in line or " and " in line) or single_author_line) and not line.endswith("."):
            if any(h in line for h in AFFIL_HINTS):
                continue
            candidates.append(line)

    if candidates:
        # 이전에는 candidates[-1](마지막 매치)을 채택했는데, 옛 논문 포맷처럼
        # "Abstract" 같은 명시적 구분 단어가 없는 문서에서는 STOP_WORDS가
        # 걸리지 않아 스캔이 본문 깊숙이까지 계속되고, 우연히 이름처럼
        # 보이는 본문 문장(예: 약어가 많은 초록 문장)을 잘못 채택하는
        # 사고가 실제로 확인되었다(1995년 논문 OCR 테스트). 저자 표기는
        # 거의 항상 문서 맨 앞에 나오므로, 첫 번째로 유효성 검증을 통과하는
        # 후보를 우선한다.
        for candidate in candidates:
            cleaned = _clean_author_line(candidate)
            if _looks_like_author_list(cleaned):
                return cleaned, "low"
    return "확인필요", None

def split_first_author(authors_str):
    if not authors_str or authors_str == "확인필요":
        return "확인필요"
    first_chunk = re.split(r",| and ", authors_str)[0].strip()
    return first_chunk or "확인필요"

# =====================================================================
# [추가됨] 공동 1저자(equal contribution) 탐지
# =====================================================================
# 흔히 "†These two authors contributed equally to this work"처럼, 저자
# 위첨자 중 특정 기호(주로 †, 드물게 ‡)를 공유하는 저자 2명 이상이 공동
# 1저자임을 밝히는 각주가 따로 붙는다. 이 각주가 없으면 위첨자 기호가
# 우연히 겹친 것일 수 있으므로(예: 서로 다른 소속의 저자가 우연히 같은
# 번호), 반드시 이 각주 문구가 실제로 존재할 때만 시도한다.
EQUAL_CONTRIB_FOOTNOTE_REGEX = re.compile(
    r"(?:these\s+(?:two|three|four|\d+)?\s*authors?\s+contributed\s+equally"
    # "to this work"만 인식하던 것을 study/manuscript/paper/article 등
    # 흔한 변형까지 넓힌다(실측: 넓히기 전엔 0건 검출됐는데, 각주 자체가
    # 없어서가 아니라 문구가 좁게 잡혀서 놓쳤을 가능성이 제기됨).
    r"|contributed\s+equally\s+to\s+(?:this|the)\s+(?:work|study|manuscript|paper|article)"
    # "to this X" 없이 그냥 "contributed equally"로 끝나는 경우도 흔하다.
    r"|contributed\s+equally\b"
    r"|equally\s+contributed"
    r"|equal\s+contribution"
    r"|co-?first\s+authors?"
    r"|joint\s+first\s+authors?"
    r"|equal\s+first\s+authors?)",
    re.IGNORECASE,
)

# 저자 이름 줄 앞에 위첨자 코드만 있는 별도 줄이 오는 경우가 있다(실측:
# "1† 1† 2 3 2 1*\nJi-Eun Lee , Sang-Eun Lee , ..." - pdfplumber가 위첨자의
# 수직 위치 때문에 이름 줄과 분리해서 뽑아냄. 지금까지 확인한 "마커가
# 이름 줄과 분리되는" 문제의 연장선). 코드 줄의 각 토큰(예: "1†", "2", "1*")은
# 이름 줄의 같은 순서의 저자에 대응한다.
_EQUAL_CONTRIB_CODE_TOKEN = re.compile(r"^\d{1,2}[\*∗†✉‡]*$")

def _find_codes_and_names_lines(lines):
    for idx in range(len(lines) - 1):
        tokens = lines[idx].strip().split()
        if len(tokens) >= 2 and all(_EQUAL_CONTRIB_CODE_TOKEN.match(t) for t in tokens):
            next_line = lines[idx + 1].strip()
            if next_line and (("," in next_line) or (" and " in next_line.lower())):
                return tokens, next_line
    return None, None

def extract_equal_first_authors(full_text):
    """본문에 "contributed equally" 각주가 있을 때만, 위첨자 코드 줄과
    이름 줄을 순서대로 대응시켜 공동 1저자 이름 목록을 반환한다. 개수가
    안 맞거나 각주가 없으면 빈 리스트를 반환해서(추측하지 않고) 호출부가
    기존 로직을 그대로 쓰게 한다.

    각주 존재 여부는 저자 줄 근처(본문 앞부분)에서만 확인한다. 최근
    논문들은 본문 뒤쪽(2~3페이지)에 "Author Contributions"(CRediT) 섹션을
    따로 두고 "All authors contributed equally to the preparation of this
    manuscript"처럼 일반적인 협업 문구를 쓰는 경우가 흔한데, 이건 저자
    "순서"(누가 공동 1저자인가)와는 무관한 문구다. 검색 범위를 본문
    전체로 열어두면 이런 문구까지 걸려서 실제로는 문제 없는 논문 다수가
    "매칭 실패"로 잘못 표시되는 사고가 있었다(실측: 1,427건 배치에서
    0건 -> 121건으로 급증, 그 중 상당수가 일반 저자 목록만 있고 실제
    공동 1저자 표시는 없는 논문으로 추정). 진짜 공동 1저자 각주는 항상
    저자/소속 정보 근처(첫 페이지 앞부분)에 나오므로, 그 구간만 본다.
    """
    if not EQUAL_CONTRIB_FOOTNOTE_REGEX.search(full_text[:3500]):
        return []
    lines = full_text.splitlines()
    tokens, names_line = _find_codes_and_names_lines(lines)
    if not tokens or not names_line:
        return []
    name_parts = [p.strip(" ,") for p in re.split(r",|\band\b", names_line, flags=re.IGNORECASE) if p.strip(" ,")]
    if len(name_parts) != len(tokens):
        return []
    equal_authors = []
    for tok, name in zip(tokens, name_parts):
        if re.search(r"[†‡]", tok) and _SINGLE_NAME_TOKEN.match(name) and not _contains_non_name_word(name):
            equal_authors.append(name)
    return equal_authors if len(equal_authors) >= 2 else []
# =====================================================================

def metadata_from_fallback(filepath, first_page_text, full_text):
    filename = os.path.basename(filepath)

    year_confidence = None
    wiley_parsed = parse_wiley_entomological_filename(filename)
    if wiley_parsed:
        title, wiley_author, wiley_year = wiley_parsed
        journal = "Entomological Research"
        year = wiley_year
    else:
        title = clean_filename_to_title(filename)
        journal = detect_journal_fallback(first_page_text)
        year, year_confidence = guess_year_fallback(first_page_text)
        wiley_author = None

    volume, page, volpage_confidence = guess_volume_page_fallback(first_page_text)

    authors, author_confidence = extract_authors_fallback(full_text, title=title)
    if (not authors or authors == "확인필요") and wiley_author:
        authors = f"{wiley_author} 외 (파일명 기준 - 전체 저자는 본문 확인 필요)"
        first_author = f"{wiley_author} (파일명 기준)"
        author_confidence = "low"
    else:
        first_author = split_first_author(authors)

    return {
        "저자": authors,
        "년도": year,
        "제목": title,
        "저널 이름": journal,
        "볼륨": volume,
        "추가 정보(페이지)": page,
        "1저자": first_author,
        "DOI": "확인필요",
        "_author_confidence": author_confidence,
        "_year_confidence": year_confidence,
        "_volpage_confidence": volpage_confidence,
    }


# =====================================================================
# [추가됨] 이메일-저자 매칭을 위한 헬퍼 함수들
# =====================================================================
def extract_email_from_text(text):
    """본문에서 이메일 형태의 문자열을 1개 추출합니다."""
    m = re.search(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', text)
    return m.group(0) if m else None

def parse_author_list(authors_str):
    """'저자' 문자열에서 개별 이름만 분리하여 리스트로 만듭니다."""
    if not authors_str or authors_str == "확인필요":
        return []
    cleaned_text = re.sub(r'[\d\*∗†✉‡]', '', authors_str)
    raw_list = re.split(r',|\band\b|·', cleaned_text)
    return [name.strip() for name in raw_list if name.strip()]

def match_author_by_email(author_list, email):
    """이메일 아이디와 저자 이름 간의 유사도를 비교해 교신저자를 찾습니다."""
    if not email or not author_list:
        return None
    
    email_id = email.split('@')[0].lower()
    best_author = None
    highest_score = -1.0
    
    for author in author_list:
        clean_name = re.sub(r'[^a-zA-Z]', '', author.lower())
        if not clean_name:
            continue
            
        similarity = SequenceMatcher(None, email_id, clean_name).ratio()
        char_overlap = sum(1 for char in email_id if char in clean_name) / len(email_id) if len(email_id) > 0 else 0
        
        combined_score = (similarity * 0.5) + (char_overlap * 0.5)
        
        if combined_score > highest_score:
            highest_score = combined_score
            best_author = author
            
    # 매칭 점수가 0.45 이상일 때만 유효한 교신저자로 간주 (최후 수단이므로 임계값을 높임)
    return best_author if highest_score > 0.45 else None


def corresponding_matches_author_list(corresponding, authors_str):
    """교신저자로 채택된 값이 이미 확보된 '저자' 목록 중 누구와도 대응되지
    않으면 False를 반환한다(값을 고치는 게 아니라 "이상해 보이는지"만
    감지하는 용도). 실측 사고: 참고문헌 약칭("Bibl. Lichenol.")이나 논문
    섹션 단어("Treatment")가 교신저자로 잘못 채택됨 - 둘 다 저자 목록 어디
    에도 없는 값이었다. 이런 새로운 오탐 유형은 단어를 하나씩 블록리스트에
    추가해 막기보다("Sex", "Isolates" 때처럼), "저자 목록과 안 맞는다"는
    구조적 신호로 일반화해서 잡아 LLM 2차 검토로 넘기는 편이 낫다.

    저자 목록 자체가 확인필요이거나 비어 있으면 판단 근거가 없으므로
    True(문제없음으로 간주)를 반환해서 괜한 오탐을 늘리지 않는다.

    [수정됨] 예전에는 공백까지 다 지우고 문자만 이어붙인 뒤 부분 문자열
    포함 여부로 판단했는데(corr_norm in author_norm), 이 방식은 이름의
    일부가 잘려나간 경우를 못 잡는 사각지대가 있었다(실측: PDF 줄바꿈이
    이름 중간에 걸려 "Bong-Sik Yun"의 앞부분이 잘린 "Sik Yun"이 교신저자로
    채택됨 - "sikyun"이 "bongsikyun" 안에 글자 그대로 포함돼 있어서 "일치"로
    오판하고 통과시켰다). 이름은 단어(토큰) 단위로 의미를 가지므로, 이제는
    단어 경계를 유지한 채 토큰 시퀀스가 완전히 일치하는지로 비교한다."""
    author_list = parse_author_list(authors_str)
    if not author_list:
        return True
    corr_tokens = re.findall(r"[a-z]+", corresponding.lower())
    if not corr_tokens:
        return True
    for author in author_list:
        author_tokens = re.findall(r"[a-z]+", author.lower())
        if not author_tokens:
            continue
        if corr_tokens == author_tokens:
            return True
    return False


# =====================================================================
# [수정 완료] 강화된 교신저자 추출 함수 (이메일 매칭 알고리즘 포함)
# =====================================================================
def _extract_name_before_marker(text, marker_chars):
    """줄 안에서 marker_chars(별표 등) 바로 앞의 '한 명 분량'만 잘라낸다.
    정규식으로 이름을 욕심껏(chain) 매칭하지 않고, 가장 가까운 콤마나
    ' and ' 경계를 기준으로 잘라서 이전 저자 이름이 함께 붙는 것을 방지한다."""
    results = []
    marker_pattern = re.compile("[" + re.escape(marker_chars) + "]")
    lines = text.splitlines()
    for idx, line in enumerate(lines):
        for m in marker_pattern.finditer(line):
            marker_idx = m.start()
            before = line[:marker_idx].rstrip()
            if not before:
                continue
            # 이름 뒤에 붙은 소속기관 위첨자 숫자(예: "Kim1,2", "Kim1, 2")를 먼저
            # 제거한다. 안 그러면 위첨자 안의 콤마(1,2)를 저자 구분 콤마로
            # 착각해서 엉뚱하게 "2"만 잘라내는 사고가 난다.
            before_clean = re.sub(r"(?:\d+\s*,\s*)*\d+\s*$", "", before).rstrip()
            comma_idx = before_clean.rfind(",")
            and_idx = before_clean.rfind(" and ")
            boundary = max(comma_idx, and_idx + len(" and ") - 1 if and_idx != -1 else -1)
            candidate = before_clean[boundary + 1:].strip(" ,")
            # 토큰 수 자체는 제한하지 않는다(실제로 단일 성(姓)만 표기되는
            # 저자도 있을 수 있어, 무조건 2단어 이상을 요구하면 정상 케이스를
            # 놓칠 위험이 있다). 대신 표 헤더/섹션 제목류의 흔한 비이름
            # 단어(Sex, Isolates 등, 실측 사고 2건 모두 이 유형)만 블록리스트로
            # 걸러낸다.
            if candidate and re.fullmatch(r"[A-Z][A-Za-z.\-]+(?:\s+[A-Z][A-Za-z.\-]+){0,3}", candidate) and not _contains_non_name_word(candidate):
                if candidate.lower() not in NON_NAME_CANDIDATE_WORDS:
                    results.append(candidate)

        # 단독저자 논문은 위첨자 마커(*)가 이름과 같은 줄이 아니라 그 위의
        # 별도 줄로 분리되어 추출되는 경우가 있다(실측: "*\nMin-Kyung Kim"처럼
        # 마커 한 글자만 있는 줄이 이름 줄보다 먼저 나옴 - pdfplumber가 위첨자의
        # 수직 위치 때문에 별도 줄로 인식). 이런 "마커 문자로만 이루어진 줄"을
        # 만나면, 바로 다음 비어있지 않은 줄을 이름 후보로 확인한다. 다만
        # "*\nCorresponding author: ..."처럼 다음 줄이 사람 이름이 아닌 경우도
        # 있으므로, SINGLE_NAME_LINE 형태(대문자로 시작하는 단어 2~4개)를
        # 엄격히 만족할 때만 채택해서 오탐을 막는다.
        stripped = line.strip()
        if stripped and all((ch in marker_chars or ch.isspace()) for ch in stripped):
            j = idx + 1
            while j < len(lines) and not lines[j].strip():
                j += 1
            if j < len(lines):
                candidate = lines[j].strip()
                if re.fullmatch(SINGLE_NAME_LINE, candidate) and candidate.lower() not in (
                    NON_NAME_CANDIDATE_WORDS
                ) and not _contains_non_name_word(candidate):
                    results.append(candidate)
                elif _looks_like_author_list(candidate):
                    # 다수 저자 목록 전체 앞에 마커가 떨어져 나온 경우
                    # (실측: "*\nHee-Jong Yang, Su-Ji Jeong, ... and Do-Youn
                    # Jeong"). 처음에는 "마지막 저자 = 교신저자"로 추측했으나,
                    # 이미 마커로 확실히 확인된 29건을 대상으로 검증한 결과
                    # 79%(23/29)만 일치했고, 나머지는 오히려 "첫 번째 저자"가
                    # 교신저자인 경우가 많았다(예: "Sang Yeob Lee, Hang Yeon
                    # Weon, ..." -> 실제로는 Sang Yeob Lee). 4건 중 1건꼴로
                    # 틀리는 추측은 "확인필요"보다 위험하므로(사람이 틀린 값을
                    # 그대로 믿을 수 있음) 값을 추측하지 않는다. 대신 이 경우를
                    # 별도로 표시해두어(AMBIGUOUS_MARKER 전역 상태) 호출부
                    # (process_pdf)가 "확인필요"에 수동 확인이 필요하다는
                    # 안내를 비고에 남길 수 있게 한다.
                    _AMBIGUOUS_MARKER_FLAG["found"] = True
    return results


def _front_matter_window(full_text):
    """저자/소속/교신저자 마커(*, †, ✉, ‡)는 거의 항상 초록(Abstract) 앞
    (front matter)에만 등장한다. 본문 뒤쪽(Materials and Methods, Results
    등)에 나오는 무관한 별표는 대부분 통계적 유의성 표시(예: "p<0.05*")나
    소프트웨어/방법론 인용(예: "PAUP*", "ACE*")인데, 마커 기반 교신저자
    추출(_extract_name_before_marker)이 본문 전체(full_text, 최대 3페이지)를
    훑다 보니 이런 잡음까지 이름으로 오인해서 교신저자 필드에 섞여 들어가는
    사고가 실측 확인됐다(예: "Jong-Chun Cheong, FMB", "Seung Hun Yu, PAUP",
    "Young-Han You, PA, PA, PA, PD" - 전부 진짜 저자명 뒤에 본문 어딘가의
    무관한 약어가 콤마로 이어붙은 형태). "Abstract"가 시작되는 지점까지만
    (못 찾으면 앞 3000자까지만) 검색 범위를 제한해서 원천 차단한다."""
    m = re.search(r"\bABSTRACT\b", full_text[:6000], re.IGNORECASE)
    if m and m.start() > 200:
        return full_text[:m.start()]
    return full_text[:3000]


def extract_corresponding_author(full_text, authors_str):
    _AMBIGUOUS_MARKER_FLAG["found"] = False
    # 주의: 이름 사이 공백은 반드시 [ ,]+ 처럼 '줄바꿈을 포함하지 않는' 클래스만
    # 써야 한다. \s를 쓰면 줄바꿈까지 건너뛰어서, 이전 줄(예: 제목의 마지막 단어)과
    # 다음 줄(진짜 이름)이 하나로 이어붙는 오류가 생긴다.
    NAME = r"[A-Z][A-Za-z.\-]+"
    NAME_CHAIN = rf"{NAME}(?:[ ,]+{NAME}){{0,4}}"

    # 아래 패턴 1~5는 논문 본문에 '명시적으로' 표시된 신호(별표, CONTACT,
    # Corresponding author 등)를 근거로 하므로, 이메일 아이디의 문자열 유사도로
    # '추측'하는 것보다 항상 신뢰도가 높다. 따라서 이 패턴들을 먼저 시도하고,
    # 전부 실패했을 때만 마지막 수단으로 이메일 유사도 매칭을 사용한다.
    # (이메일 계정명이 개인 이름과 무관한 기관 공용 주소인 경우가 흔해서,
    #  이메일 매칭을 최우선으로 두면 엉뚱한 저자가 잘못 채택되는 사고가 난다.)

    # 패턴 1) Taylor & Francis 다중 교신저자
    # 일부 PDF는 작은대문자(small-caps) 폰트를 쓰는데, 시각적으로는 대문자로
    # 보여도 PDF 내부 텍스트 데이터 자체는 소문자로 저장되어 있는 경우가 있다
    # (예: "CONTACT Seung-yeol lee leesy1123@knu.ac.kr" - 실제 인쇄본은
    # "Seung-Yeol Lee"). 그러면 "대문자로 시작해야 함" 조건 때문에 이름
    # 사슬이 중간에서 끊겨 이메일과 연결이 안 되므로, 이 패턴만 대소문자를
    # 구분하지 않고 매칭한 뒤 결과를 .title()로 정상 표기로 되돌린다.
    NAME_CI = r"[A-Za-z][A-Za-z.\-]+"
    NAME_CHAIN_CI = rf"{NAME_CI}(?:[ ,]+{NAME_CI}){{0,4}}"
    m1 = re.search(r"CONTACT[ \t]+(.+?@[\w.\-]+(?:\s*;\s*.+?@[\w.\-]+)*)", full_text, re.IGNORECASE)
    if m1:
        contact_line = m1.group(1)
        authors_raw = contact_line.split(';')
        extracted_authors = []
        for author_info in authors_raw:
            name_match = re.search(rf"({NAME_CHAIN_CI})[ \t]+[\w.\-]+@", author_info.strip())
            if name_match:
                extracted_authors.append(name_match.group(1).strip().title())
        if extracted_authors:
            return ", ".join(extracted_authors)

    # 흔히 2단(2-column) 레이아웃 PDF에서 "Correspondence"(왼쪽 컬럼 제목)와
    # "Abstract"(오른쪽 컬럼 제목)이 텍스트 추출 시 같은 줄에 붙어버리는 경우가 있다.
    # 그래서 키워드 바로 다음 단어를 무조건 이름으로 채택하면 "Abstract" 같은
    # 섹션 제목을 이름으로 잘못 인식하게 된다. 이를 막기 위해 이름 후보를
    # (1) 최소 2단어 이상 (2) 흔한 섹션 제목 단어가 아님을 검증한다.
    def _valid_name(name):
        words = name.split()
        if len(words) < 2:
            return False
        if words[0].lower() in NON_NAME_CANDIDATE_WORDS:
            return False
        if _contains_non_name_word(name):
            return False
        return True

    # 패턴 2) "Corresponding author:" 또는 "Correspondence:" 명시형
    # 키워드 뒤에 흔한 섹션제목 단어가 붙어있으면 건너뛰고, 그 다음 줄까지 살펴본다.
    # 이름은 공백으로만 이어진 1~4단어("Alireza Askarianzadeh")로 본다.
    PERSON_NAME_STRICT = rf"{NAME}(?:\s+{NAME}){{0,3}}"
    for kw_match in re.finditer(r"(?:Corresponding author|Correspondence)", full_text, re.IGNORECASE):
        window = full_text[kw_match.end():kw_match.end() + 200]
        # 같은 줄에 겹쳐 붙은 섹션제목 단어(Abstract 등) 제거
        window = re.sub(r"^[ \t]*:?[ \t]*(?:Abstract|Keywords|Introduction)\b", "", window, flags=re.IGNORECASE)

        # 먼저 콤마 직전까지만(공백으로만 연결된) 엄격한 이름 매칭을 시도한다.
        # "Alireza Askarianzadeh, Plant Protection Department..."처럼 소속기관명이
        # 다음 줄로 넘어가는 바람에 ",\s*Department" 절단 조건이 걸리지 않아
        # "Plant Protection"까지 이름으로 잘못 삼키는 사고를 막기 위함이다.
        m2_strict = re.match(rf"^[ \t\n]*({PERSON_NAME_STRICT})\s*,", window)
        if m2_strict:
            name = m2_strict.group(1).strip(" ,*∗")
            if _valid_name(name):
                return name

        m2 = re.search(rf"^[ \t\n]*({NAME_CHAIN})", window)
        if m2:
            name = m2.group(1).strip()
            name = re.split(r",\s*(?:Department|Division|Institute|College|University|School|Center)", name)[0]
            name = name.strip(" ,*∗")
            if _valid_name(name):
                return name

    # [수정됨] 패턴 4, 5는 검색 범위를 초록(Abstract) 이전 구간으로 제한한다
    # (_front_matter_window 참고 - 본문 전체를 훑으면 무관한 별표/마커까지
    # 이름으로 오인해 교신저자에 섞여 들어가는 사고가 있었다).
    marker_window = _front_matter_window(full_text)

    # 패턴 4) 별표(*, ∗) 다중 추출 - 콤마 경계 기준으로 한 명씩만 추출
    cleaned_names = _extract_name_before_marker(marker_window, "*∗")
    if cleaned_names:
        return ", ".join(cleaned_names)

    # 패턴 5) 저자 목록 이메일 기호(†, ✉, ‡) 옆의 이름 추출
    cleaned_names = _extract_name_before_marker(marker_window, "†✉‡")
    if cleaned_names:
        return ", ".join(cleaned_names)

    # 마지막 수단) 명시적 신호가 전혀 없을 때만 이메일 유사도로 추측
    # (오탐 방지를 위해 임계값을 0.3 -> 0.45로 상향)
    email = extract_email_from_text(full_text)
    if email:
        author_list = parse_author_list(authors_str)
        matched_author = match_author_by_email(author_list, email)
        if matched_author:
            return matched_author

    return "확인필요"
# =====================================================================

# extract_corresponding_author 처리 중, "마커는 있는데 다수 저자라 누구인지
# 특정 불가"인 경우를 process_pdf에 알리기 위한 상태. 값을 추측해서 채우면
# 오답 위험이 있어(검증 결과 79%만 일치) 비워두는 대신, 비고에 "수동 확인
# 필요" 안내를 남기기 위한 용도로만 쓴다.
_AMBIGUOUS_MARKER_FLAG = {"found": False}


def process_pdf(filepath):
    filename = os.path.basename(filepath)
    # [수정됨] "파일명" 컬럼에는 파일명만이 아니라 PDF_DIR 기준 상대경로를
    # 저장한다. 하위 폴더(트리 구조)로 정리된 PDF를 main()이 재귀적으로는
    # 잘 찾아내지만, 이후 run_llm_second_pass가 "🔍 LLM 2차 검토 필요" 행을
    # 다시 열 때는 os.path.join(pdf_dir, 파일명)으로 경로를 평평하게(하위
    # 폴더 없이) 재조립했다. PDF가 하위 폴더 안에 있으면 이 재조립된 경로가
    # 실제 파일 위치와 달라져서 "파일을 찾을 수 없습니다"로 전부 건너뛰는
    # 사고가 실측 확인됐다(919건 중 251건 전부 스킵됨). 상대경로를 저장해
    # 두면 os.path.join(pdf_dir, 상대경로)가 하위 폴더 몇 단계든 정확히
    # 재구성된다. 제목 정리/문서유형 판별 등 "내용"에 관한 로직은 계속
    # filename(파일명만)을 그대로 쓴다 - 폴더 경로가 제목에 섞이면 안 되므로.
    relpath = os.path.relpath(filepath, PDF_DIR)
    row = {
        "참고문헌 종류": "Journal Article",
        "저자": "", "년도": "", "제목": "", "저널 이름": "", "볼륨": "",
        "추가 정보(페이지)": "", "1저자": "", "교신저자": "", "DOI": "",
        "비고": "", "파일명": relpath,
    }
    notes = []

    try:
        with pdfplumber.open(filepath) as pdf:
            full_text, first_page_text = get_pdf_text(pdf, max_pages=3)
    except Exception as e:
        row["비고"] = f"⚠️ 파일 열기 실패: {e}"
        return row

    real_text_len = len(re.sub(r"\s+", "", first_page_text))
    text_unusable = real_text_len < 150 or is_font_garbled(first_page_text)

    if real_text_len < 150:
        notes.append("⚠️ 텍스트 추출 불가 (스캔본으로 추정됨) - Vision LLM 2차 처리 필요")
    elif is_font_garbled(first_page_text):
        # 텍스트는 뽑혔지만 폰트에 유니코드 매핑이 없어 "(cid:123)" 같은
        # 내부 글리프 ID로만 채워진 경우. 글자 수 기준은 통과하지만 사실상
        # 읽을 수 없는 텍스트라 정규식을 태워봐야 의미가 없다.
        notes.append("⚠️ 폰트 인코딩 깨짐 ((cid:...) 형태) - Vision LLM 2차 처리 필요")

    if not text_unusable:
        doc_type = detect_document_type(first_page_text)
        if doc_type:
            row["참고문헌 종류"] = doc_type

    doi = extract_doi(full_text)
    crossref_data = None

    if doi:
        crossref_data = metadata_from_crossref(doi)
        time.sleep(0.2)

    # [추가됨] DOI를 못 찾았거나 CrossRef 조회가 실패한 경우, 기존 정규식
    # 폴백으로 바로 넘어가기 전에 GROBID로 먼저 헤더를 추출한다. GROBID가
    # 제목을 잡아주면 그 제목으로 CrossRef 서지검색을 한 번 더 시도해서,
    # DOI가 아예 없거나(오래된 호) 본문에서 못 찾은 논문도 CrossRef의 정식
    # 메타데이터를 최대한 끌어온다.
    grobid_fields = None
    source_tag = None  # [추가됨] 이 행의 메타데이터가 실제로 어디서 왔는지 추적
    doi_lookup_failed = bool(doi) and not crossref_data  # 진단용: DOI는 찾았는데 CrossRef 직접조회만 실패한 경우
    if not crossref_data and not text_unusable:
        grobid_fields = fetch_grobid_header(filepath)

        # [수정됨] 제목 서지검색에 쓸 제목을 고른다. 한국균학회지처럼 논문
        # 자체가 한글로 쓰인 경우 GROBID가 한글 제목을 그대로 뽑아오는데,
        # CrossRef 레코드는 거의 항상 영문 제목으로 등록돼 있어서 한글로
        # 검색하면 애초에 매칭될 수가 없다(crossref_search_by_title도 한글
        # 제목이면 아예 요청을 안 보내도록 방어해뒀지만, 여기서 애초에 더
        # 나은 후보를 고르는 게 먼저다). 업로드된 파일들의 파일명 자체가
        # 거의 항상 이미 영문 제목으로 정리되어 있음을 확인했으므로,
        # GROBID 제목이 한글이거나 아예 없으면 파일명을 대신 쓴다.
        #
        # 이렇게 하면 "GROBID가 실패한 경우"에도(grobid_fields가 None이어도)
        # 파일명 기반 영문 제목으로 CrossRef 재시도를 할 수 있다 - 이전에는
        # GROBID 성공이 전제조건이라, GROBID까지 실패하면 DOI 직접조회
        # 실패 건들이 CrossRef를 다시 시도해볼 기회 자체가 없었다.
        search_title = grobid_fields.get("title", "") if grobid_fields else ""
        if not search_title or _contains_hangul(search_title):
            search_title = clean_filename_to_title(filename)

        crossref_msg = crossref_search_by_title(search_title)
        if crossref_msg:
            crossref_data = _metadata_from_crossref_message(crossref_msg)
            time.sleep(0.2)
            source_tag = (
                "DOI 직접조회 실패 -> 영문 제목 서지검색으로 복구"
                if doi_lookup_failed else
                "제목 서지검색으로 CrossRef 메타데이터 확보 (DOI 없음)"
            )

    if crossref_data:
        row.update(crossref_data)
        # CrossRef는 성공했지만 페이지 범위 자체가 등록되어 있지 않은 경우
        # (Entomological Research 다수가 이런 케이스). 본문 맨 위 저널 배너에서
        # 직접 볼륨/페이지를 다시 추출해 빈 값만 채워 넣는다.
        # (CrossRef를 썼는지 텍스트 기반인지는 결과가 정상이면 비고에 남기지 않는다 -
        # 처리 과정보다 최종 결과의 정상/이상 여부만 비고에 표시하기로 함)
        volpage_confidence = "crossref"  # CrossRef가 준 값은 별도 신뢰도 등급을 매기지 않음
        if row.get("추가 정보(페이지)") in (None, "", "확인필요"):
            fallback_vol, fallback_page, volpage_confidence = guess_volume_page_fallback(first_page_text)
            if fallback_page != "확인필요":
                row["추가 정보(페이지)"] = fallback_page
                if row.get("볼륨") in (None, "", "확인필요"):
                    row["볼륨"] = fallback_vol
        # [추가됨] CrossRef 레코드에 container-title(저널 이름) 필드 자체가
        # 비어 있는 경우도 있다(실측: KJM DOI 132건 - 저자 필드 비는 것과
        # 같은 원인, 국내 학회 저널 특유의 CrossRef 등록 메타데이터가
        # 부실한 경우가 많음). 저자/볼륨과 동일한 패턴으로, 비어있을 때만
        # 본문 배너 기반 폴백으로 보충한다.
        if row.get("저널 이름") in (None, "", "확인필요"):
            fallback_journal = detect_journal_fallback(first_page_text)
            if fallback_journal != "확인필요":
                row["저널 이름"] = fallback_journal
        # CrossRef 레코드에 author 필드 자체가 비어 있는 경우가 있다(한국균학회지
        # 일부 DOI에서 실측 확인). 이 경우 CrossRef는 "성공"으로 처리되어
        # metadata_from_fallback이 아예 호출되지 않으므로, 본문에 저자명이
        # 멀쩡히 있어도 계속 "확인필요"로 남는다. 저자 필드만 비어 있으면
        # 본문 기반 폴백으로 별도 보충한다.
        author_confidence = "crossref"  # CrossRef가 준 저자는 별도 신뢰도 등급을 매기지 않음
        year_confidence = "crossref"
        if row.get("저자") in (None, "", "확인필요") and not text_unusable:
            fallback_authors, author_confidence = extract_authors_fallback(full_text, title=row.get("제목", ""))
            if fallback_authors and fallback_authors != "확인필요":
                row["저자"] = fallback_authors
                row["1저자"] = split_first_author(fallback_authors)
    elif text_unusable:
        # 텍스트 자체를 못 믿는 상황(스캔본/폰트 깨짐)이고 CrossRef도 실패했다면,
        # 정규식 폴백을 그냥 태우지 않는다 - 깨진 텍스트에서 나온 "그럴듯한
        # 값"은 확인필요보다 위험하다(조용히 틀린 채로 채워짐). 제목만
        # 파일명에서 채우고 나머지는 전부 확인필요로 남겨서, 2단계에서
        # Vision LLM이 이미지를 직접 보고 처리하게 한다.
        row["제목"] = clean_filename_to_title(filename)
        for f in ["저자", "년도", "저널 이름", "볼륨", "추가 정보(페이지)", "1저자", "DOI"]:
            row[f] = "확인필요"
        author_confidence = None
        year_confidence = None
        volpage_confidence = None
    elif grobid_fields:
        # [추가됨] CrossRef(DOI 직접조회+제목 서지검색) 둘 다 실패했지만
        # GROBID 헤더 추출은 성공한 경우 - 기존 정규식 폴백보다 이 결과를
        # 우선 채택한다. GROBID가 저자는 잘 뽑아도 저널명/연도는 한국
        # 학회지 다수에서 비어있는 경우가 실측으로 많이 확인됐다(430건,
        # "저자는 채워졌는데 저널명만 통째로 빈칸"인 행). CrossRef 성공
        # 분기에서만 넣었던 저널명/연도/볼륨/페이지 본문 배너 보완을
        # 여기에도 똑같이 적용한다 - 안 그러면 GROBID 경로를 탄 행일수록
        # 저널명이 빈칸으로 남는 사각지대가 생긴다.
        source_tag = "GROBID 헤더 추출 결과 채택 (CrossRef 실패)"
        row.update(metadata_from_grobid(grobid_fields))
        author_confidence = "grobid"
        year_confidence = "grobid"
        volpage_confidence = "grobid"
        if row.get("저널 이름") in (None, "", "확인필요"):
            fallback_journal = detect_journal_fallback(first_page_text)
            if fallback_journal != "확인필요":
                row["저널 이름"] = fallback_journal
        if row.get("년도") in (None, "", "확인필요"):
            fallback_year, fallback_year_conf = guess_year_fallback(first_page_text)
            if fallback_year != "확인필요":
                row["년도"] = fallback_year
                year_confidence = fallback_year_conf
        if row.get("추가 정보(페이지)") in (None, "", "확인필요"):
            fallback_vol, fallback_page, volpage_confidence = guess_volume_page_fallback(first_page_text)
            if fallback_page != "확인필요":
                row["추가 정보(페이지)"] = fallback_page
                if row.get("볼륨") in (None, "", "확인필요"):
                    row["볼륨"] = fallback_vol
        if doi:
            row["DOI"] = doi  # 본문에서 정규식으로 직접 찾은 DOI가 있으면 GROBID 값보다 우선 보존
    else:
        # [추가됨] GROBID까지 시도했는데도 실패했는지, 애초에 GROBID를
        # 시도할 필요조차 없었는지(텍스트는 멀쩡한데 GROBID 서버가 꺼져
        # 있는 등) 구분해서 표시한다.
        if not _grobid_enabled():
            source_tag = "정규식 폴백 (GROBID 서버 연결 불가)"
        else:
            source_tag = "정규식 폴백 (GROBID도 헤더 인식 실패)"
        fallback_data = metadata_from_fallback(filepath, first_page_text, full_text)
        author_confidence = fallback_data.pop("_author_confidence", None)
        year_confidence = fallback_data.pop("_year_confidence", None)
        volpage_confidence = fallback_data.pop("_volpage_confidence", None)
        row.update(fallback_data)
        # 본문에서는 DOI를 찾았지만 CrossRef 조회만 실패한 경우, 어렵게 찾은
        # DOI 값 자체는 살려둔다(폴백 경로가 무조건 "확인필요"로 덮어쓰던 버그).
        if doi:
            row["DOI"] = doi

    # [추가됨] 제목/저자에 한글이 섞여 있으면 영문으로 정리한다. 업로드된
    # 파일들의 파일명이 거의 항상 이미 영문 제목으로 정리되어 있음을
    # 확인했으므로, 제목은 그걸로 대체한다. 저자는 한글/영문이 나란히
    # 병기된 경우(예: "가강현, Kang-Hyeon Ka")가 많아, 한글 토큰만 걷어내고
    # 영문 표기는 최대한 살린다. 교신저자 추출(아래)이 이 정리된 저자
    # 목록을 근거로 매칭하므로, 반드시 그 전에 수행한다.
    if _contains_hangul(row.get("제목", "")):
        row["제목"] = clean_filename_to_title(filename)
        notes.append("ℹ️ 제목에 한글이 섞여 있어 파일명 기반 영문 제목으로 대체함")

    if row.get("저자") and row["저자"] != "확인필요" and _contains_hangul(row["저자"]):
        cleaned_authors = _strip_hangul_author_tokens(row["저자"])
        if cleaned_authors:
            row["저자"] = cleaned_authors
            row["1저자"] = split_first_author(cleaned_authors)
            notes.append("ℹ️ 저자 목록에서 한글 이름 토큰 제거, 영문 표기만 유지")
        else:
            # 영문 표기가 하나도 안 남았다면(=저자 전체가 한글로만 인쇄된
            # 논문) 억지로 비우지 않고 원래 값을 유지하되, 반드시 사람이
            # 확인하도록 표시한다.
            notes.append("⚠️ 저자가 한글로만 표기됨 - 영문 표기 확인 필요")

    # 수정됨: 앞서 확보한 `row["저자"]` 문자열을 인자로 전달하여 이메일 매칭 수행
    # (텍스트를 못 믿는 상황이면 이 역시 건너뛰고 확인필요로 둔다)
    if text_unusable and not crossref_data:
        corresponding = "확인필요"
        _AMBIGUOUS_MARKER_FLAG["found"] = False  # 건너뛴 경우 이전 파일의 값이 남지 않도록 명시적으로 초기화
    elif grobid_fields and grobid_fields.get("corresponding") and not _contains_hangul(grobid_fields["corresponding"]):
        # [추가됨] GROBID가 이메일 매칭으로 이미 교신저자를 특정해준 경우
        # (CrossRef 제목검색 실패로 GROBID 결과를 그대로 쓰는 상황) 그 값을
        # 우선 채택한다. 다만 아래 corresponding_matches_author_list 등
        # 안전장치는 출처와 무관하게 그대로 적용되므로 이상하면 여전히 표시된다.
        # 한글 이름이면(위 저자 정리와 일관되게) 영문 매칭을 다시 시도하도록
        # 건너뛰고 extract_corresponding_author로 넘긴다.
        corresponding = grobid_fields["corresponding"]
    else:
        corresponding = extract_corresponding_author(full_text, row["저자"])
    row["교신저자"] = corresponding
    if _contains_hangul(corresponding):
        notes.append("⚠️ 교신저자가 한글로만 표기됨 - 영문 표기 확인 필요")
    # 교신저자 인식 실패 여부는 "교신저자" 컬럼 자체가 "확인필요"로 이미 표시되므로
    # 비고에 별도로 중복 기재하지 않는다(비고는 구조적 이슈만 담당).
    # 다만 "마커는 감지됐지만 다수 저자라 누구인지 특정 불가"였던 경우는
    # 예외적으로 표시한다 - 이건 그냥 "완전히 못 찾음"과 달리 "본문에
    # 교신저자 표시 자체는 있으니 사람이 직접 보면 바로 알 수 있는" 케이스라,
    # 다른 확인필요 행들보다 우선적으로 봐야 함을 알려줄 가치가 있다.
    if corresponding == "확인필요" and _AMBIGUOUS_MARKER_FLAG["found"]:
        notes.append("⚠️ 교신저자 마커 감지됨 - 저자가 여럿이라 특정 불가, 수동 확인 필요")

    # 교신저자로 채택된 값이 이미 확보된 저자 목록 어디에도 없는 경우(실측:
    # 참고문헌 약칭 "Bibl. Lichenol.", 섹션 단어 "Treatment"가 교신저자로
    # 잘못 채택됨)를 감지해서 표시한다. 값을 되돌리거나 추측으로 고치지
    # 않고 - 그러다 또 다른 오탐이 날 수 있으므로 - 표시만 해서 LLM 2차
    # 검토로 넘긴다.
    if corresponding not in ("확인필요", "") and not corresponding_matches_author_list(corresponding, row.get("저자", "")):
        notes.append("⚠️ 교신저자가 저자 목록과 매칭되지 않음 - 오탐 의심, 확인 필요")

    # 문서 전체는 "폰트 깨짐"으로 분류될 만큼 심하지 않아도(is_font_garbled
    # 임계값 미달), 마커 글자 하나만 유니코드 매핑이 없어 "(cid:19)" 같은
    # 잔해가 저자/교신저자 필드에 그대로 새어 들어오는 경우가 있었다(실측:
    # "...Chang Soo Lee(cid:)"). 문서 단위 판정과 별개로 이 필드 자체에
    # 잔해가 남아있으면 항상 표시해서 새는 일이 없게 한다.
    if CID_GARBLE_ANY_REGEX.search(row.get("저자", "") or "") or CID_GARBLE_ANY_REGEX.search(row.get("교신저자", "") or ""):
        notes.append("⚠️ 저자/교신저자에 폰트 인코딩 잔해((cid:...)) 감지됨 - 확인 필요")

    # [추가됨] "(cid:...)" 패턴과는 다른 종류의 폰트 깨짐 - 매핑은 있는데
    # 표현이 불가능해 유니코드 대체문자(U+FFFD, " ")로 채워진 경우도 같은
    # 방식으로 감지한다(실측: 저자 필드 전체가 " " 범벅으로 나왔는데
    # 아무 표시 없이 나갈 뻔함).
    if REPLACEMENT_CHAR_REGEX.search(row.get("저자", "") or "") or REPLACEMENT_CHAR_REGEX.search(row.get("교신저자", "") or "") or REPLACEMENT_CHAR_REGEX.search(row.get("제목", "") or ""):
        notes.append("⚠️ 저자/교신저자/제목에 폰트 인코딩 깨짐(대체문자 ' ') 감지됨 - 확인 필요")

    # 저자 필드가 "본문 스캔으로 이름처럼 생긴 줄을 채택"(low 신뢰도) 경로로
    # 채워진 경우, 위첨자/소속기관 단서로 확인된 것보다 오탐 위험이 훨씬
    # 높다(실제로 본문 문장이 저자로 잘못 채택된 사고가 있었음). 값 자체는
    # 채워두되(공란보다는 낫다는 판단), 사람이 우선적으로 검수해야 함을
    # 표시해서 "채워져 있으니 맞겠지"라고 그냥 넘어가지 않게 한다.
    if author_confidence == "low":
        notes.append("⚠️ 저자 낮은 신뢰도 - 본문 스캔으로 추정된 값, 확인 권장")

    # 저자와 마찬가지로, 연도/볼륨/페이지도 신뢰도가 낮은 경로로 채워진
    # 경우가 있다(예: 연도 - 전용 배너 없이 그냥 첫 4자리 숫자를 채택;
    # 볼륨/페이지 - 서로 다른 구간에서 독립적으로 찾은 값을 조합). 이런
    # 필드도 채워져 있다고 그냥 넘기지 않도록 동일하게 표시한다.
    if year_confidence == "low":
        notes.append("⚠️ 년도 낮은 신뢰도 - 본문 추정값, 확인 권장")
    if volpage_confidence == "low":
        notes.append("⚠️ 볼륨/페이지 낮은 신뢰도 - 서로 다른 위치에서 조합한 값, 확인 권장")

    # 정정문(Erratum/Correction/Corrigendum) 여부를 제목/파일명으로 추정해서 태깅.
    # 제목이 채워진 뒤에 판단해야 하므로 여기서 수행한다. 결과가 맞고 틀리고와는
    # 무관하게 "이 문서의 성격이 무엇인지"를 알려주는 용도다.
    title_and_filename = f"{row.get('제목','')} {filename}"
    # Addendum(원 논문에 누락된 내용을 보완하는 공식 후속 문서)도
    # Erratum/Correction과 같은 성격("원 논문 1편에 딸린 문서")이라
    # 같은 태그로 묶는다.
    if re.search(r"\b(Erratum|Correction|Corrigendum|Addendum)\b", title_and_filename, re.IGNORECASE):
        notes.append("ℹ️ 정정문(Erratum/Correction/Addendum) 형태로 추정됨")

    row["저널 이름"] = normalize_journal_name(row.get("저널 이름"))

    # CrossRef/폴백 어느 경로든 상관없이, 본문에 "contributed equally" 각주가
    # 있으면 공동 1저자로 재표시한다(그동안 1저자 필드가 목록의 첫 번째
    # 사람만 기계적으로 채택해서, 공동 1저자 표시를 조용히 놓치고 있었음).
    # 텍스트를 못 믿는 상황(스캔본/폰트 깨짐)이고 CrossRef도 못 쓴 경우는
    # 이 각주 자체를 찾을 수 없으므로 건너뛴다.
    if not (text_unusable and not crossref_data):
        equal_authors = extract_equal_first_authors(full_text)
        if equal_authors:
            row["1저자"] = ", ".join(equal_authors) + " (공동 1저자)"
        elif EQUAL_CONTRIB_FOOTNOTE_REGEX.search(full_text[:3500]):
            # 본문에 "contributed equally" 각주는 분명히 있는데, 위첨자
            # 코드 줄과 이름 줄의 개수가 안 맞는 등 위치 매칭에 실패한
            # 경우다. 이때 그냥 조용히 "1저자 = 목록 맨 앞 사람"으로
            # 넘어가면, 실제로는 공동 1저자인데 한 명만 표시되는 틀린
            # 값이 아무 표시 없이 나갈 위험이 있다 - 반드시 표시해야 한다.
            notes.append("⚠️ 공동 1저자(contributed equally) 각주 감지됨 - "
                          "위치 매칭 실패, 1저자 수동 확인 필요")

    if source_tag:
        notes.insert(0, f"🔍 출처: {source_tag}")

    # [수정됨] 지금까지 쌓아온 상세 판단 근거(저자 낮은 신뢰도, 한글 섞임,
    # cid 잔해, 출처 등)는 row["_debug_notes"]에 그대로 보관한다(엑셀 컬럼
    # 목록에 없는 내부용 키라 최종 결과 파일에는 안 나간다 - 콘솔 요약
    # print_batch_summary에서만 계속 쓰인다). 화면에 보이는 "비고"는 이제
    # "LLM 2차 검토가 필요한지, 필요하면 어느 필드인지"만 담은 단순한
    # 표시로 바꾼다 - 판단 근거를 일일이 안 읽어도 무엇을 봐야 할지 바로
    # 알 수 있게 하기 위함이다.
    debug_notes = " / ".join(notes)
    row["_debug_notes"] = debug_notes

    vision_review = VISION_REVIEW_MARKER in debug_notes
    review_fields = [] if vision_review else _compute_review_fields(row, debug_notes)
    row["_vision_review"] = vision_review
    row["_review_fields"] = review_fields

    if vision_review:
        row["비고"] = "🔍 Vision LLM 2차 검토 필요 (스캔본/폰트 깨짐)"
    elif review_fields:
        row["비고"] = "🔍 LLM 2차 검토 필요: " + ", ".join(review_fields)
    else:
        row["비고"] = ""
    return row

FIELDS_FOR_DEDUP_CHECK = ["참고문헌 종류", "저자", "년도", "제목", "저널 이름", "볼륨",
                          "추가 정보(페이지)", "1저자", "교신저자", "DOI"]

def print_batch_summary(rows):
    """이번 배치 결과를 필드별 채움률 + 비고 사유별 건수로 요약해서 출력한다.
    표본 검증을 어디부터 봐야 할지(확인필요 행, 낮은 신뢰도 행, 마커 애매
    행 등) 한눈에 알 수 있게 하기 위한 용도다. 전체를 다 사람이 보지 않고
    이 요약이 가리키는 행들만 우선 검수하면 된다."""
    total = len(rows)
    if total == 0:
        return

    print("\n" + "=" * 60)
    print("배치 결과 요약")
    print("=" * 60)

    fields = ["년도", "제목", "저널 이름", "볼륨", "추가 정보(페이지)",
              "1저자", "교신저자", "DOI"]
    print("[필드별 채움률]")
    for f in fields:
        missing = sum(1 for r in rows if not str(r.get(f, "")).strip() or r.get(f) == "확인필요")
        filled_pct = round((total - missing) / total * 100, 1)
        print(f"  {f}: {filled_pct}% ({total - missing}/{total})")

    print("\n[비고 표시 - 우선 검수 권장 행]")
    # [수정됨] 화면에 보이는 "비고"는 이제 단순화된 표시라, 상세 사유별
    # 분류는 process_pdf가 저장해둔 내부용 _debug_notes(엑셀에는 안 나감)
    # 텍스트를 함께 봐서 판별한다.
    note_categories = [
        ("텍스트 추출 불가(스캔본 추정)", "텍스트 추출 불가"),
        ("폰트 인코딩 깨짐", "폰트 인코딩 깨짐"),
        ("교신저자 마커 애매", "교신저자 마커 감지됨"),
        ("저자 낮은 신뢰도", "저자 낮은 신뢰도"),
        ("년도 낮은 신뢰도", "년도 낮은 신뢰도"),
        ("볼륨/페이지 낮은 신뢰도", "볼륨/페이지 낮은 신뢰도"),
        ("공동 1저자 매칭 실패", "공동 1저자"),
        ("정정문 추정", "정정문"),
    ]
    any_notes = False
    for label, marker in note_categories:
        n = sum(1 for r in rows if marker in ((r.get("_debug_notes") or "") + " " + (r.get("비고") or "")))
        if n:
            any_notes = True
            print(f"  {label}: {n}건 ({round(n/total*100,1)}%)")
    n_dup = sum(1 for r in rows if r.get("_duplicate"))
    if n_dup:
        any_notes = True
        print(f"  중복 추정: {n_dup}건 ({round(n_dup/total*100,1)}%)")
    n_llm_updated = sum(1 for r in rows if r.get("_llm_updated"))
    n_llm_reviewed = sum(1 for r in rows if r.get("_llm_reviewed"))
    if n_llm_reviewed:
        any_notes = True
        print(f"  LLM 2차 검토로 갱신됨: {n_llm_updated}건 ({round(n_llm_updated/total*100,1)}%)")
        print(f"  LLM 2차 검토 완료(갱신 없음): {n_llm_reviewed - n_llm_updated}건 "
              f"({round((n_llm_reviewed - n_llm_updated)/total*100,1)}%)")
    if not any_notes:
        print("  없음 - 비고 표시된 행이 없습니다.")

    n_vision = sum(1 for r in rows if r.get("_vision_review"))
    n_text_review = sum(1 for r in rows if r.get("_review_fields") and not r.get("_vision_review"))
    n_clean = total - n_vision - n_text_review
    print("\n[최종 비고 기준 요약 - 실제 엑셀에 보이는 값]")
    print(f"  정상(검토 불필요): {n_clean}건 ({round(n_clean/total*100,1)}%)")
    print(f"  🔍 LLM 2차 검토 필요: {n_text_review}건 ({round(n_text_review/total*100,1)}%)")
    print(f"  🔍 Vision LLM 2차 검토 필요: {n_vision}건 ({round(n_vision/total*100,1)}%)")

    print("=" * 60)

def flag_duplicate_rows(rows):
    # 파일명만 다르고(혹은 같고) 나머지 필드가 전부 동일한 행은 진짜 중복일
    # 가능성이 높다. 다만 파일명만 같고 내용은 실제로 다른 별개의 논문인
    # 경우도 있었기 때문에(예: 같은 이름의 서로 다른 논문 2편), 파일명이
    # 아니라 '파일명을 제외한 나머지 필드 전체'가 완전히 일치하는 경우만
    # 중복으로 판단한다. 잘못 지웠다가 진짜 논문이 사라지는 사고를 막기 위해
    # 삭제하지 않고 비고에 표시만 해서 사람이 직접 확인하게 한다.
    signature_to_indices = {}
    for i, row in enumerate(rows):
        # 파일 열기 자체가 실패한 행은 모든 필드가 똑같이 빈 값이라,
        # 서로 무관한 실패 파일끼리 "중복"으로 오판될 수 있다. 제목이
        # 없는(=처리 자체가 안 된) 행은 애초에 비교 대상에서 제외한다.
        if not str(row.get("제목", "")).strip():
            continue
        signature = tuple(row.get(f, "") for f in FIELDS_FOR_DEDUP_CHECK)
        signature_to_indices.setdefault(signature, []).append(i)

    for indices in signature_to_indices.values():
        if len(indices) < 2:
            continue
        for i in indices:
            note = "⚠️ 중복 추정"
            existing = rows[i].get("비고", "") or ""
            rows[i]["비고"] = (existing + " / " if existing else "") + note
            rows[i]["_duplicate"] = True  # 콘솔 요약(print_batch_summary)용 내부 표시
    return rows


# =====================================================================
# 2단계: LLM(Gemini API, 무료 티어) 재검토
# =====================================================================
# 1단계(정규식/CrossRef)가 "확인필요"로 남겼거나 저신뢰도로 표시한 행만
# 골라서 재검토한다. 나머지(비고 없이 확실한 행)는 API를 아예 호출하지
# 않아 비용을 아낀다. regex는 레이아웃(줄바꿈 위치, 2단 배치 등)이 바뀔
# 때마다 새로 패턴을 짜야 하는 두더지잡기가 되지만, LLM은 본문을 사람처럼
# 통째로 읽으므로 이런 레이아웃 변형에 훨씬 안정적이다.

REVIEW_TARGET_FIELDS = ["년도", "제목", "저널 이름", "볼륨", "추가 정보(페이지)",
                        "저자", "1저자", "교신저자", "DOI"]

FIELD_DESCRIPTIONS = {
    "년도": "출판 연도 (4자리 숫자)",
    "제목": "논문 제목 (원문 그대로)",
    "저널 이름": "저널 이름 (정식 명칭)",
    "볼륨": "권(Volume) 번호",
    "추가 정보(페이지)": "페이지 범위 (예: 123-130)",
    "저자": "전체 저자 목록 - 논문에 인쇄된 순서와 표기 그대로, 콤마로 구분",
    "1저자": "제1저자 이름. 본문에 '두 저자가 동등하게 기여했다(contributed equally)'는"
             " 각주가 있으면 해당 저자 전원의 이름을 콤마로 나열하고 끝에 '(공동 1저자)'를 덧붙일 것",
    "교신저자": "교신저자(corresponding author) 이름 - 여러 명이면 콤마로 구분",
    "DOI": "논문 DOI",
}

def _compute_review_fields(row, note):
    """상세 판단 근거(note, 내부용 문자열)를 바탕으로 LLM 2차 검토가
    필요한 필드 목록을 계산한다. process_pdf가 행을 만드는 시점에 한 번
    호출해서 row["_review_fields"]에 저장해두고, 이후에는 이 계산을 다시
    하지 않는다(비고가 이제 단순화된 표시라 문자열을 다시 파싱할 수 없음).
    - 값 자체가 "확인필요"인 필드는 당연히 포함한다.
    - 값이 채워져 있어도 저신뢰도/애매 신호가 있었던 필드는 포함한다
      (채워져 있다고 다 맞는 게 아니라는 걸 이미 확인했기 때문 - "마지막
      저자=교신저자" 규칙이 79%만 맞았던 사례 참고)."""
    review = set()
    for f in REVIEW_TARGET_FIELDS:
        if str(row.get(f, "")).strip() in ("", "확인필요"):
            review.add(f)
    if "저자 낮은 신뢰도" in note:
        review.update({"저자", "1저자"})
    if "교신저자 마커 감지됨" in note:
        review.add("교신저자")
    if "년도 낮은 신뢰도" in note:
        review.add("년도")
    if "볼륨/페이지 낮은 신뢰도" in note:
        review.update({"볼륨", "추가 정보(페이지)"})
    if "공동 1저자" in note and "매칭 실패" in note:
        review.add("1저자")
    if "교신저자가 저자 목록과 매칭되지 않음" in note:
        review.add("교신저자")
    if "폰트 인코딩 잔해" in note:
        review.update({"저자", "교신저자"})
    if "폰트 인코딩 깨짐(대체문자" in note:
        review.update({"저자", "교신저자", "제목"})
    if "저자가 한글로만 표기됨" in note:
        review.add("저자")
    if "교신저자가 한글로만 표기됨" in note:
        review.add("교신저자")
    return sorted(review)

def identify_review_fields(row):
    """이 행에서 LLM 2차 검토가 필요한 필드 목록을 반환한다. process_pdf가
    이미 계산해서 row["_review_fields"]에 저장해둔 값을 그대로 돌려준다.
    (예전엔 비고 텍스트를 매번 다시 문자열 매칭했는데, 비고가 이제 단순한
    표시로 바뀌면서 그 방식이 더 이상 안 통한다.) 구버전 체크포인트처럼
    이 값이 없는 행이 섞여 있어도 죽지 않도록 빈 리스트를 기본값으로 둔다."""
    return row.get("_review_fields", [])

def _strip_json_fence(text):
    text = text.strip()
    text = re.sub(r"^```(?:json)?\s*", "", text)
    text = re.sub(r"\s*```$", "", text)
    return text.strip()

def call_gemini_for_review(full_text, filename, fields):
    """지정된 필드만 Gemini API에 다시 확인시킨다. 본문 텍스트를 통째로
    넘겨서 모델이 사람처럼 읽고 판단하게 하며, 확실하지 않으면 반드시
    "확인불가"라고 답하도록 강하게 지시한다(추측성 오답보다 "모르겠다"가
    낫다는 원칙 - 지난번 "마지막 저자" 추측 규칙이 79%만 맞아서 위험했던
    사례와 같은 이유).

    반환값: (결과dict 또는 None, 할당량_초과로_실패했는지 bool). 호출부
    (run_llm_second_pass)가 이 두 번째 값을 보고 "이건 코드로는 답이 없는
    할당량 문제다"를 판단해서, 남은 수백 건을 헛되이 계속 두드리지 않고
    조기에 멈출 수 있게 한다(실측: 이걸 못 알아채고 251건 내내 429만
    받다가 사용자가 직접 중단해야 했던 사고가 있었음)."""
    if not GEMINI_API_KEY:
        print("[LLM] ⚠️ GEMINI_API_KEY가 설정되지 않아 2차 검토를 건너뜁니다. "
              "실행 전에 os.environ[\"GEMINI_API_KEY\"] = \"AIza...\" 를 설정하세요.")
        return None, False

    field_lines = "\n".join(f'- "{f}": {FIELD_DESCRIPTIONS.get(f, f)}' for f in fields)
    prompt = f"""다음은 학술 논문 PDF에서 추출한 본문 텍스트입니다 (파일명: {filename}).
아래 필드들을 본문에서 찾아 JSON 객체 하나로만 답하세요. 다른 설명이나 코드블록 없이 순수 JSON만 출력하세요.

찾아야 할 필드:
{field_lines}

규칙:
- 본문에서 명확히 확인되는 값만 채우세요. 절대 추측하지 마세요.
- 확실하지 않거나 본문에서 찾을 수 없으면 반드시 "확인불가"라고 쓰세요. 틀린 값을 채우는 것보다 "확인불가"가 낫습니다.
- 저자/1저자/교신저자는 반드시 영문 표기로만 답하세요. 본문에 한글 이름이 함께 인쇄되어 있어도(예: "가강현, Kang-Hyeon Ka") 한글은 절대 쓰지 말고 영문 로마자 표기만 쓰세요. 여러 명이면 콤마로 구분하세요.
- 저자/1저자/교신저자에는 사람 이름만 쓰세요. 소속기관명, 대학교명, 주소, 우편번호 등은 절대 포함하지 마세요.
- JSON 키는 위에 나열된 필드명과 정확히 동일한 문자열을 쓰세요.

본문:
{full_text[:6000]}
"""

    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {"temperature": 0},
    }
    params = {"key": GEMINI_API_KEY}

    last_exc = None
    last_status = None
    for attempt in range(LLM_MAX_RETRIES):
        try:
            resp = requests.post(GEMINI_API_URL, params=params, json=payload, timeout=LLM_TIMEOUT)
            last_status = resp.status_code
            if resp.status_code == 200:
                data = resp.json()
                try:
                    raw = data["candidates"][0]["content"]["parts"][0]["text"]
                except (KeyError, IndexError):
                    print(f"[LLM] ⚠️ 응답 형식이 예상과 다름 ({filename}): {str(data)[:200]}")
                    return None, False
                cleaned = _strip_json_fence(raw)
                try:
                    return json.loads(cleaned), False
                except json.JSONDecodeError:
                    print(f"[LLM] ⚠️ JSON 파싱 실패 ({filename}): {raw[:200]}")
                    return None, False
            if resp.status_code == 429 or resp.status_code >= 500:
                # [수정됨] 지금까지는 재시도만 조용히 하고 원인을 안 찍어서,
                # 재시도가 다 소진돼도 "왜" 실패했는지(요청 한도 초과인지,
                # 키가 무효화됐는지, 서버 오류인지) 알 수가 없었다. 실측:
                # "재시도 소진, 실패 ... None"만 찍혀서 진단이 안 됐다.
                print(f"[LLM] ⚠️ {resp.status_code} 응답 ({filename}, 시도 {attempt+1}/{LLM_MAX_RETRIES}): {resp.text[:200]}")
                retry_after = resp.headers.get("Retry-After")
                wait = float(retry_after) if retry_after and retry_after.isdigit() else (2 ** attempt)
                time.sleep(wait)
                continue
            print(f"[LLM] ⚠️ API 오류 {resp.status_code} ({filename}): {resp.text[:200]}")
            return None, False
        except requests.exceptions.RequestException as e:
            last_exc = e
            time.sleep(2 ** attempt)
    print(f"[LLM] ⚠️ 재시도 소진, 실패 ({filename}): {last_exc}")
    return None, last_status == 429

VISION_REVIEW_MARKER = "Vision LLM 2차 처리 필요"

def identify_vision_review_files(rows):
    """스캔본/폰트 깨짐으로 텍스트 자체를 못 믿는 파일 목록만 모아서 반환한다.
    이런 파일은 텍스트 기반 정규식은 물론 텍스트 기반 LLM(run_llm_second_pass)도
    쓸모가 없다 - 애초에 읽을 수 있는 텍스트가 없기 때문이다. 이미지를 직접
    보는 Vision LLM으로 별도 처리해야 하므로, 여기서는 대상 파일명만 추려서
    반환한다(API 호출은 하지 않음 - 사용자가 별도 단계에서 처리).
    process_pdf가 저장해둔 row["_vision_review"] 플래그를 우선 쓰고, 구버전
    체크포인트처럼 그 값이 없는 행은 비고 텍스트로 대체 판별한다."""
    return [
        r.get("파일명") for r in rows
        if r.get("_vision_review", VISION_REVIEW_MARKER in (r.get("비고") or ""))
    ]

def _looks_like_bad_author_value(value):
    """LLM(2차 검토)이 돌려준 저자/1저자/교신저자 값이 실제 사람 이름이
    아닌 것(한글 병기, 소속기관/주소, 폰트 인코딩 깨짐 잔해)을 담고
    있으면 True. 1단계(GROBID/정규식 경로)에서 이미 검증된 신호들
    (_contains_non_name_word, CID_GARBLE_ANY_REGEX, REPLACEMENT_CHAR_REGEX)을
    LLM 결과에도 똑같이 적용한다 - 출처가 다르다고 오탐 패턴까지 달라지는
    건 아니기 때문이다."""
    return (
        _contains_hangul(value)
        or _contains_non_name_word(value)
        or bool(CID_GARBLE_ANY_REGEX.search(value))
        or bool(REPLACEMENT_CHAR_REGEX.search(value))
    )


def _strip_bad_author_tokens(authors_str):
    """저자 문자열에서 위 기준에 걸리는 개별 토큰만 제거하고 나머지
    (정상으로 보이는 이름)는 살린다."""
    parts = [p.strip(" ,") for p in re.split(r",|\band\b", authors_str, flags=re.IGNORECASE)]
    kept = [p for p in parts if p and not _looks_like_bad_author_value(p)]
    return ", ".join(kept)


def run_llm_second_pass(rows, pdf_dir=None):
    """1단계 결과(rows)에서 검토가 필요하다고 표시된 행만 골라 Gemini API로
    재확인시키고, row를 제자리에서 갱신한다(반환값은 같은 rows 리스트).
    표시가 전혀 없는 행은 API를 아예 호출하지 않는다.

    스캔본/폰트 깨짐(Vision LLM 대상) 행은 여기서 자동으로 건너뛴다 - 읽을
    수 있는 텍스트가 없어 텍스트 기반 프롬프트가 무의미하기 때문이다.
    이 파일들은 identify_vision_review_files(rows)로 따로 모아서, 이미지를
    직접 넣는 별도의 Vision LLM 처리 단계에서 다뤄야 한다.

    사용법 (Colab에서 1단계 실행 후 별도 셀에서):
        os.environ["GEMINI_API_KEY"] = "AIza..."
        rows = main()                       # 1단계
        rows = run_llm_second_pass(rows)     # 2단계(텍스트 기반)
        vision_targets = identify_vision_review_files(rows)  # 3단계 후보만 모으기
        pd.DataFrame(rows, columns=[...]).to_excel("결과_2차검토완료.xlsx", index=False)
    """
    if pdf_dir is None:
        pdf_dir = PDF_DIR

    to_review = [
        (i, r) for i, r in enumerate(rows)
        if identify_review_fields(r) and not r.get("_vision_review", VISION_REVIEW_MARKER in (r.get("비고") or ""))
    ]
    n_vision = sum(1 for r in rows if r.get("_vision_review", VISION_REVIEW_MARKER in (r.get("비고") or "")))
    print(f"[LLM 2차 검토] 전체 {len(rows)}건 중 {len(to_review)}건이 텍스트 기반 검토 대상입니다 "
          f"(신뢰도 높은 {len(rows) - len(to_review) - n_vision}건은 API 호출 생략, "
          f"스캔본/폰트깨짐 {n_vision}건은 Vision LLM 대상이라 여기서 건너뜀 - "
          f"identify_vision_review_files(rows)로 따로 확인하세요).")

    # [추가됨] 할당량(quota)이 완전히 막힌 상태(예: 모델 자체가 무료 티어에서
    # 제외됨)라면, 남은 수백 건을 계속 두드려봐야 전부 똑같이 실패한다.
    # 연속으로 QUOTA_ABORT_THRESHOLD번 할당량 초과가 나면, 남은 항목은
    # 건드리지 않고(오답으로 안 덮어씀) 조기에 멈춰서 안내만 남긴다 - 이전엔
    # 이걸 사람이 직접 Ctrl+C로 멈춰야 했다.
    QUOTA_ABORT_THRESHOLD = 3
    consecutive_quota_failures = 0

    for n, (i, row) in enumerate(to_review, start=1):
        try:
            fields = identify_review_fields(row)
            filename = row.get("파일명", "")
            filepath = os.path.join(pdf_dir, filename)
            print(f"[LLM {n}/{len(to_review)}] {filename} - 확인 필드: {fields}")

            if not os.path.exists(filepath):
                print(f"  ⚠️ 파일을 찾을 수 없어 건너뜁니다: {filepath}")
                continue

            try:
                with pdfplumber.open(filepath) as pdf:
                    full_text, _ = get_pdf_text(pdf, max_pages=3)
            except Exception as e:
                print(f"  ⚠️ 파일 열기 실패: {e}")
                continue

            # [수정됨] call_gemini_for_review가 (결과, 할당량초과여부) 튜플을
            # 반환하도록 바뀌었는데 호출부를 안 맞춰서 AttributeError가
            # 났었다("'tuple' object has no attribute 'get'"). 이제 정상적으로
            # 언패킹한다.
            result, quota_exhausted = call_gemini_for_review(full_text, filename, fields)
            if quota_exhausted:
                consecutive_quota_failures += 1
                if consecutive_quota_failures >= QUOTA_ABORT_THRESHOLD:
                    print(
                        f"\n⛔ 연속 {consecutive_quota_failures}건이 할당량 초과(429, limit 0 등)로 "
                        f"실패했습니다. 남은 {len(to_review) - n}건은 시도해도 똑같이 실패할 가능성이 "
                        f"높아 여기서 멈춥니다. https://ai.dev/rate-limit 에서 실제 한도/모델을 "
                        f"확인한 뒤, 다시 run_llm_second_pass(rows)를 호출하면 안 된 부분부터 "
                        f"이어서 시도합니다."
                    )
                    break
            else:
                consecutive_quota_failures = 0

            if not result:
                continue

            changed = []
            for f in fields:
                new_val = result.get(f)
                if not new_val or str(new_val).strip() in ("", "확인불가"):
                    continue
                # [수정됨] 프롬프트에서 영문/사람 이름만 답하라고 명시했지만,
                # LLM이 그래도 병기된 한글, 소속기관/주소, 또는 본문의 폰트
                # 깨짐 잔해를 그대로 옮겨오는 경우가 실측 확인됐다(1단계
                # GROBID/정규식 경로에서 겪었던 것과 같은 부류의 문제).
                # 같은 원칙으로 처리한다: 저자는 문제되는 토큰만 걷어내고
                # 나머지는 살리고, 1저자/교신저자처럼 한 사람뿐인 필드는
                # 통째로 문제 있으면 값을 아예 버리고(=기존값 유지) 넘어간다.
                if f == "저자" and _looks_like_bad_author_value(new_val):
                    new_val = _strip_bad_author_tokens(new_val)
                    if not new_val:
                        continue
                elif f in ("1저자", "교신저자") and _looks_like_bad_author_value(new_val):
                    continue
                old_val = row.get(f)
                row[f] = new_val
                if old_val != new_val:
                    changed.append(f)

            # [수정됨] "검토 필요" 상태는 이 시점에 해소됐으므로(갱신했든,
            # 갱신할 게 없어서 그대로 뒀든) 이전 비고에 계속 덧붙이지 않고
            # 결과로 교체한다. 갱신할 내용이 없으면 더 이상 사람이 볼 필요가
            # 없다고 보고 비고를 비운다(=다른 정상 행들과 동일하게 취급).
            row["_llm_reviewed"] = True
            row["_llm_updated"] = bool(changed)
            if changed:
                row["비고"] = f"✅ 2차 검토로 갱신됨: {', '.join(changed)}"
            else:
                row["비고"] = ""
        finally:
            # [수정됨] 파일을 못 찾았거나 API 호출이 실패해서 continue로
            # 넘어가는 경우에도 반드시 대기한다. try/finally로 감싸서
            # continue가 어디서 발생하든 항상 실행되게 한다. 예전에는
            # 실패 시 대기 없이 곧바로 다음 요청을 보내서, 429(할당량 초과)를
            # 만났을 때 오히려 더 빠르게 다음 요청을 쏴 상황을 악화시켰다.
            time.sleep(LLM_REQUEST_INTERVAL)

    return rows
# =====================================================================

# =====================================================================
# [추가됨] 체크포인트 (10,000개 이상 대량 배치용)
# =====================================================================
_checkpoint_lock = threading.Lock()

def load_checkpoint(path):
    """이미 처리된 파일들의 결과를 체크포인트(JSON Lines)에서 불러온다.
    대량 배치는 중간에 끊기는 일이 흔하므로, 재실행 시 이미 끝난 파일은
    건너뛰고 나머지만 이어서 처리하기 위한 것이다. 반환값은
    {파일명: 결과행dict} - 파일명이 키라서 중복 없이 자연스럽게 병합된다."""
    done_rows = {}
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                try:
                    row = json.loads(line)
                except json.JSONDecodeError:
                    continue  # 중간에 끊겨서 마지막 줄이 깨진 경우 등 - 그 줄만 버리고 계속
                if row.get("파일명"):
                    done_rows[row["파일명"]] = row
    return done_rows

def append_checkpoint(path, row):
    """결과 행 하나를 체크포인트 파일에 즉시 이어 쓴다(파일 하나 처리
    끝날 때마다 호출). 스레드 여러 개가 동시에 쓰므로 락으로 보호한다."""
    with _checkpoint_lock:
        with open(path, "a", encoding="utf-8") as f:
            f.write(json.dumps(row, ensure_ascii=False) + "\n")
# =====================================================================


def run_sample_test(n=50, seed=None, workers=None):
    """10,000개 본 배치를 돌리기 전에, 무작위로 n개만 뽑아서 미리 검증하는
    함수. main()과 로직은 동일하지만(GROBID -> CrossRef -> 정규식 폴백,
    병렬 처리) 체크포인트/결과 파일을 본 배치용과 완전히 분리된 별도
    경로에 써서, 샘플 테스트가 본 배치의 진행 상황(체크포인트)을 절대
    건드리지 않는다.

    사용 예:
        rows = run_sample_test(50)                # 시드 없이 매번 다른 50개
        rows = run_sample_test(50, seed=42)        # 같은 50개로 재현 테스트
    """
    import random

    pdf_files = sorted(glob.glob(os.path.join(PDF_DIR, "**/*.pdf"), recursive=True))
    if not pdf_files:
        print(f"⚠️ {PDF_DIR} 아래에서 PDF를 찾지 못했습니다.")
        return []

    rng = random.Random(seed)
    sample = rng.sample(pdf_files, min(n, len(pdf_files)))
    total = len(sample)
    print(f"전체 {len(pdf_files)}개 중 무작위 {total}개 샘플 테스트 시작")

    _grobid_enabled()  # 최초 1회만 접속 확인

    sample_checkpoint = "/content/checkpoint_sample.jsonl"
    sample_output = "/content/샘플_테스트_결과.xlsx"
    # 매번 새로 테스트한다는 가정이 자연스러우므로(오래된 샘플 결과가 섞여
    # 헷갈리지 않도록), 이전 샘플 체크포인트는 시작 시 초기화한다. 본
    # 배치용 CHECKPOINT_PATH는 절대 건드리지 않는다.
    if os.path.exists(sample_checkpoint):
        os.remove(sample_checkpoint)

    rows = []
    start_time = time.time()
    progress_lock = threading.Lock()

    def _run_and_checkpoint(fp):
        row = process_pdf(fp)
        append_checkpoint(sample_checkpoint, row)
        return row

    with ThreadPoolExecutor(max_workers=workers or MAX_WORKERS) as executor:
        futures = {executor.submit(_run_and_checkpoint, fp): fp for fp in sample}
        for future in as_completed(futures):
            fp = futures[future]
            try:
                row = future.result()
            except Exception as e:
                row = {"파일명": os.path.relpath(fp, PDF_DIR), "비고": f"⚠️ 처리 중 예외 발생: {e}"}
                append_checkpoint(sample_checkpoint, row)
            with progress_lock:
                rows.append(row)
                completed = len(rows)
                elapsed = time.time() - start_time
                avg = elapsed / completed
                print(f"[{completed}/{total}] {row.get('파일명')} (평균 {avg:.1f}초/건)")

    rows = flag_duplicate_rows(rows)
    columns = ["참고문헌 종류", "저자", "년도", "제목", "저널 이름", "볼륨",
               "추가 정보(페이지)", "1저자", "교신저자", "DOI", "비고", "파일명"]
    df = pd.DataFrame(rows, columns=columns)
    df.to_excel(sample_output, index=False)

    total_elapsed = time.time() - start_time
    avg_per_file = total_elapsed / total
    print(f"\n샘플 테스트 완료: {sample_output} ({total}행)")
    print(f"평균 {avg_per_file:.1f}초/건 -> 전체 {len(pdf_files)}개라면 약 "
          f"{avg_per_file * len(pdf_files) / 60:.0f}분 예상 (동시성 {workers or MAX_WORKERS}워커 기준)")
    print_batch_summary(rows)
    return rows


def main():
    # 하위 폴더(트리 구조)의 PDF까지 모두 찾기 (recursive=True)
    pdf_files = sorted(glob.glob(os.path.join(PDF_DIR, "**/*.pdf"), recursive=True))
    total = len(pdf_files)
    print(f"총 {total}개 PDF 발견")

    # GROBID 접속 확인은 여기서 한 번만 미리 해둔다(파일마다 확인하면
    # 서버가 꺼져 있을 때 매번 타임아웃을 기다리게 되어 배치 전체가 느려짐).
    _grobid_enabled()

    done = load_checkpoint(CHECKPOINT_PATH)
    todo = [fp for fp in pdf_files if os.path.relpath(fp, PDF_DIR) not in done]
    if done:
        print(f"체크포인트에서 {len(done)}개 이어받음 - 새로 처리할 파일: {len(todo)}개")

    rows = list(done.values())
    start_time = time.time()
    completed = len(done)
    progress_lock = threading.Lock()

    def _run_and_checkpoint(fp):
        row = process_pdf(fp)
        append_checkpoint(CHECKPOINT_PATH, row)
        return row

    if todo:
        # GROBID/CrossRef 네트워크 호출을 기다리는 I/O 위주 작업이라 스레드
        # 병렬화 효과가 크다. 워커 수는 GROBID 서버 코어 수에 맞춰 조절.
        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
            futures = {executor.submit(_run_and_checkpoint, fp): fp for fp in todo}
            for future in as_completed(futures):
                fp = futures[future]
                try:
                    row = future.result()
                except Exception as e:
                    # 파일 하나가 완전히 죽어도 배치 전체가 멈추면 안 된다.
                    row = {"파일명": os.path.relpath(fp, PDF_DIR), "비고": f"⚠️ 처리 중 예외 발생: {e}"}
                    append_checkpoint(CHECKPOINT_PATH, row)
                with progress_lock:
                    rows.append(row)
                    completed += 1
                    elapsed = time.time() - start_time
                    avg = elapsed / max(1, completed - len(done))
                    remaining = avg * (total - completed)
                    print(
                        f"[{completed}/{total}] {row.get('파일명')} "
                        f"(평균 {avg:.1f}초/건 | 남은시간 약 {remaining/60:.1f}분)"
                    )

    # 병렬 처리로 완료 순서가 섞이므로, 결과를 파일 탐색 순서와 다시
    # 맞춰서(재현성 유지) 저장한다.
    order = {os.path.relpath(fp, PDF_DIR): i for i, fp in enumerate(pdf_files)}
    rows.sort(key=lambda r: order.get(r.get("파일명"), 10**9))

    rows = flag_duplicate_rows(rows)

    columns = ["참고문헌 종류", "저자", "년도", "제목", "저널 이름", "볼륨",
               "추가 정보(페이지)", "1저자", "교신저자", "DOI", "비고", "파일명"]
    df = pd.DataFrame(rows, columns=columns)
    df.to_excel(OUTPUT_XLSX, index=False)
    n_dup = sum(1 for r in rows if "중복" in (r.get("비고") or ""))
    total_elapsed = time.time() - start_time
    print(f"\n총 소요시간: {total_elapsed/60:.1f}분")
    print(f"\n완료: {OUTPUT_XLSX} ({len(df)}행, 그 중 중복 추정 {n_dup}행)")
    print_batch_summary(rows)
    return rows  # 2단계(run_llm_second_pass)에서 이어서 쓸 수 있도록 반환

RESULT_COLUMNS = ["참고문헌 종류", "저자", "년도", "제목", "저널 이름", "볼륨",
                  "추가 정보(페이지)", "1저자", "교신저자", "DOI", "비고", "파일명"]

def flag_bad_author_fields_for_review(rows):
    """이미 저장된 결과에서 "1저자"/"교신저자" 필드에 사람 이름이 아닌
    값(소속기관·주소, 폰트 깨짐 잔해)이 섞인 행만 찾아서, 그 값을
    "확인필요"로 되돌리고 그 필드만 다시 LLM 2차 검토 대상으로 표시한다.
    나머지 필드/행은 전혀 건드리지 않으므로, run_llm_second_pass(rows)를
    다시 돌려도 이번에 표시된 필드만 API로 재확인되고 이미 정상인 656건
    이상은 그대로 건너뛴다(비용/시간 절약).

    판단 기준:
    - 기관/회사/국가명 단어 포함 (University, Institute, Department, 국가명 등)
    - 숫자 포함 (실제 사람 이름에는 숫자가 없음 - 주소의 번지수/우편번호,
      소속기관의 위첨자 잔해 등으로 추정)
    - 폰트 인코딩 잔해 ((cid:...), 유니코드 대체문자 ' ')

    사용 예:
        rows = load_rows_from_excel("/content/결과_2차검토완료.xlsx")
        rows = flag_bad_author_fields_for_review(rows)
        rows = run_llm_second_pass(rows)   # 표시된 필드만 재확인됨
        save_rows_to_excel(rows, "/content/결과_2차검토완료.xlsx")
    """
    def _looks_bad(val):
        if not val or val == "확인필요":
            return False
        if _contains_non_name_word(val):
            return True
        if CID_GARBLE_ANY_REGEX.search(val) or REPLACEMENT_CHAR_REGEX.search(val):
            return True
        if re.search(r"\d", val):
            return True
        return False

    n_rows_flagged = 0
    for row in rows:
        bad_fields = [f for f in ("1저자", "교신저자") if _looks_bad(row.get(f, ""))]
        if not bad_fields:
            continue
        n_rows_flagged += 1
        for f in bad_fields:
            row[f] = "확인필요"
        # 기존에 이미 검토 표시가 남아있던 필드가 있으면 합치고, 없으면
        # 이번에 찾은 것만으로 새로 만든다.
        review_fields = sorted(set(row.get("_review_fields", [])) | set(bad_fields))
        row["_review_fields"] = review_fields
        row["_vision_review"] = False
        row["_llm_reviewed"] = False
        row["_llm_updated"] = False
        row["비고"] = "🔍 LLM 2차 검토 필요: " + ", ".join(review_fields)

    print(f"오염된 1저자/교신저자 값 발견: {n_rows_flagged}행 -> "
          f"해당 필드만 확인필요로 되돌리고 재검토 대상으로 표시했습니다.")
    return rows


def clean_hangul_from_saved_results(rows):
    """이미 저장된 결과(특히 LLM 2차 검토를 거친 행)에서 저자/1저자/교신저자에
    섞여 들어온 문제(한글 병기, 소속기관/주소, 폰트 인코딩 깨짐 잔해)를,
    API를 다시 호출하지 않고 정리한다(_looks_like_bad_author_value 기준
    - run_llm_second_pass의 검증 로직과 동일). 저자는 문제되는 토큰만
    걷어내고, 1저자/교신저자는 통째로 문제 있으면 "확인필요"로 되돌려서
    (값을 추측하지 않고) 사람이 확인하게 한다.

    사용 예:
        rows = load_rows_from_excel("/content/결과_2차검토완료.xlsx")
        rows = clean_hangul_from_saved_results(rows)
        save_rows_to_excel(rows, "/content/결과_2차검토완료.xlsx")
    """
    cleaned, flagged = 0, 0
    for row in rows:
        touched = False
        authors = row.get("저자", "")
        if authors and authors != "확인필요" and _looks_like_bad_author_value(authors):
            fixed = _strip_bad_author_tokens(authors)
            if fixed:
                row["저자"] = fixed
                row["1저자"] = split_first_author(fixed)
                cleaned += 1
                touched = True
            else:
                # 전부 문제 토큰이라 정리 후 남는 게 없는 경우. 옛 값을
                # 그대로 방치하면 조용히 틀린 채로 남아 재검토 대상에도
                # 안 잡히므로, 확인필요로 되돌려서 검토 대상에 포함시킨다.
                row["저자"] = "확인필요"
                row["1저자"] = "확인필요"
                flagged += 1
                touched = True
        for f in ("1저자", "교신저자"):
            val = row.get(f, "")
            if val and val != "확인필요" and _looks_like_bad_author_value(val):
                row[f] = "확인필요"
                flagged += 1
                touched = True

        # [수정됨] 값만 고치고 "이 행은 재검토 대상이다"라는 내부 표시
        # (_review_fields)와 화면상 비고를 그대로 두면, run_llm_second_pass가
        # 옛 표시만 보고 검토 대상에서 빠뜨리는 사고가 난다(실측: 23건 중
        # 대부분이 빠지고 1건만 검토 대상으로 잡힘). 값을 건드린 행은
        # 반드시 표시도 같이 갱신한다.
        if touched:
            review_fields = _compute_review_fields(row, "")
            row["_review_fields"] = review_fields
            if review_fields:
                row["비고"] = "🔍 LLM 2차 검토 필요: " + ", ".join(review_fields)
            else:
                row["비고"] = ""

    print(f"정리 완료: 저자 필드 문제 토큰 제거 {cleaned}건 / "
          f"1저자·교신저자 문제값 확인필요로 되돌림 {flagged}건")
    return rows


def repair_filepaths_for_tree(rows, pdf_dir=None):
    """구버전(파일명만 저장, 하위 폴더 정보 없음)으로 만들어진 rows를,
    다시 GROBID/CrossRef를 돌리지 않고 "파일명"만 실제 폴더 구조에 맞는
    상대경로로 고쳐서 이어갈 수 있게 한다. 1차 결과(저자/제목/저널/DOI
    등)는 이미 계산이 끝나 정확한 값이라 다시 계산할 필요가 없고, 망가진
    건 run_llm_second_pass가 파일을 다시 찾을 때 쓰는 "파일명" 필드뿐이기
    때문이다.

    pdf_dir(기본값 PDF_DIR) 아래를 재귀적으로 스캔해서 각 파일의 basename
    -> 실제 상대경로 매핑을 만든 뒤, row["파일명"]이 그 매핑에서 basename
    으로 못 찾아지면(=폴더 정보가 빠진 구버전 값이면) 올바른 상대경로로
    바꿔치기한다. 같은 이름의 파일이 서로 다른 폴더에 여러 개 있으면
    자동으로 결정할 수 없으므로, 그런 행은 건드리지 않고 경고만 남긴다.

    사용 예 (PDF를 트리 구조 그대로 다시 업로드한 뒤):
        rows = load_rows_from_excel("/content/결과_2차검토완료.xlsx")
        rows = repair_filepaths_for_tree(rows)
        rows = run_llm_second_pass(rows)   # 이제 파일을 정상적으로 찾음
    """
    if pdf_dir is None:
        pdf_dir = PDF_DIR

    basename_to_relpaths = {}
    for fp in glob.glob(os.path.join(pdf_dir, "**/*.pdf"), recursive=True):
        rel = os.path.relpath(fp, pdf_dir)
        basename_to_relpaths.setdefault(os.path.basename(fp), []).append(rel)

    fixed, ambiguous, missing, already_ok = 0, 0, 0, 0
    for row in rows:
        current = row.get("파일명", "")
        if not current:
            continue
        # 이미 올바른 상대경로면(파일이 그 자리에 실제로 있으면) 그대로 둔다.
        if os.path.exists(os.path.join(pdf_dir, current)):
            already_ok += 1
            continue
        basename = os.path.basename(current)
        candidates = basename_to_relpaths.get(basename, [])
        if len(candidates) == 1:
            row["파일명"] = candidates[0]
            fixed += 1
        elif len(candidates) > 1:
            ambiguous += 1
            print(f"⚠️ 동일 파일명이 여러 폴더에 있어 자동 결정 불가: {basename} -> {candidates}")
        else:
            missing += 1

    print(f"복구 완료: 이미 정상 {already_ok}건 / 경로 복구 {fixed}건 / "
          f"모호(수동 확인 필요) {ambiguous}건 / 파일 자체를 못 찾음 {missing}건")
    return rows


def load_rows_from_excel(path):
    """1단계(main()) 결과를 저장한 엑셀 파일을 다시 불러와 LLM 2차 검토를
    이어갈 수 있게 한다. process_pdf가 만드는 row["_review_fields"],
    row["_vision_review"] 같은 내부 판단 필드는 엑셀 컬럼에는 저장되지
    않으므로(최종 결과 파일을 깔끔하게 유지하려는 의도적 설계), 세션이
    끊겨 rows 변수를 잃어버린 뒤 엑셀만 갖고 재개하려 하면 이 필드들이
    없어서 identify_review_fields가 전부 빈 리스트를 반환하는 문제가
    있었다(=검토 대상이 0건으로 잘못 나옴).

    "비고" 컬럼이 이미 process_pdf가 만든 단순화된 형식
    ("🔍 LLM 2차 검토 필요: 저자, 교신저자" 등) 그대로 저장되어 있으므로,
    이 텍스트를 다시 파싱해서 내부 필드를 복구한다. PDF를 다시 처리할
    필요 없이 저장된 엑셀만으로 정확히 이어갈 수 있다.

    한계: _debug_notes(상세 판단 근거)는 엑셀에 애초에 없으므로 복구되지
    않는다 - print_batch_summary의 "저자 낮은 신뢰도" 등 세부 카테고리
    집계는 이 경로로 불러온 rows에서는 0으로 나온다(핵심 기능인 검토
    대상 판별 자체에는 영향 없음).

    사용 예:
        rows = load_rows_from_excel("/content/참고문헌_메타데이터.xlsx")
        rows = run_llm_second_pass(rows)
    """
    df = pd.read_excel(path).fillna("")
    rows = df.to_dict("records")
    for row in rows:
        note = str(row.get("비고", "") or "")
        vision = "Vision LLM 2차 검토 필요" in note
        review_fields = []
        if not vision:
            m = re.search(r"LLM 2차 검토 필요:\s*([^/]+)", note)
            if m:
                review_fields = [f.strip() for f in m.group(1).split(",") if f.strip()]
        row["_vision_review"] = vision
        row["_review_fields"] = review_fields
        row["_duplicate"] = "중복 추정" in note
    return rows


def save_rows_to_excel(rows, path):
    """2단계(LLM 재검토) 이후 결과를 저장할 때 쓰는 헬퍼.
    사용 예: save_rows_to_excel(rows, "/content/결과_2차검토완료.xlsx")"""
    df = pd.DataFrame(rows, columns=RESULT_COLUMNS)
    df.to_excel(path, index=False)
    print(f"저장 완료: {path} ({len(df)}행)")

if __name__ == "__main__":
    main()