# -*- coding: utf-8 -*-
"""
TTA 인증서 발급용 데이터 추출기 (Streamlit)

- 관리 목록(xlsx) + 시험결과요약서(PDF, 여러 개) + (선택) 회의록(hwp)을 업로드하면
- 인증번호를 기준으로 매칭해서 인증서 제작에 필요한 필드를 표로 뽑아주고
- 표는 화면에서 직접 수정 가능하며, 엑셀로 다운로드할 수 있습니다.

배포: Streamlit Community Cloud 등에 그대로 올리면 됩니다.
"""
import io
import re
import subprocess
import tempfile
from datetime import date, timedelta

import pandas as pd
import pdfplumber
import streamlit as st

st.set_page_config(page_title="TTA 인증서 데이터 추출기", layout="wide")
st.title("📄 TTA 인증서 발급용 데이터 추출기")

st.markdown(
    """
    1. **관리 목록 엑셀**을 업로드하세요. (여러 시트를 자동으로 합칩니다)
    2. **시험결과요약서 PDF**를 여러 개 업로드하세요. 표지에 있는 **인증번호**를 기준으로
       관리 목록과 매칭합니다. (인증번호가 없는 요약서는 제외됩니다)
    3. 인터넷전화(MMoIP) 관련 건이 있다면 **인증심의위원회 회의록(.hwp)** 도 업로드하세요.
       인증범위 문구를 회의록에서 가져옵니다.
    4. 결과 표를 화면에서 바로 수정한 뒤, 엑셀로 다운로드하세요.
    """
)

# ──────────────────────────────────────────────────────────────
# 업로드
# ──────────────────────────────────────────────────────────────
col1, col2, col3 = st.columns(3)
with col1:
    mgmt_file = st.file_uploader("① 관리 목록 (xlsx)", type=["xlsx"])
with col2:
    summary_files = st.file_uploader(
        "② 시험결과요약서 (PDF, 여러 개)", type=["pdf"], accept_multiple_files=True
    )
with col3:
    minutes_file = st.file_uploader("③ 회의록 (.hwp, 선택)", type=["hwp"])


# ──────────────────────────────────────────────────────────────
# 관리 목록 로드
# ──────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_management_list(file_bytes):
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    frames = []
    for sheet in xls.sheet_names:
        # 이 관리목록 양식은 1행이 시트 제목, 2행이 실제 컬럼 헤더인 구조
        df = xls.parse(sheet, header=1)
        # 혹시 헤더 행에 '인증번호' 컬럼이 없으면(양식이 다르면) header=0으로 재시도
        if "인증번호" not in df.columns:
            df = xls.parse(sheet, header=0)
        df["__시트"] = sheet
        frames.append(df)
    all_df = pd.concat(frames, ignore_index=True, sort=False)
    # 인증번호 공백/개행 정리
    if "인증번호" in all_df.columns:
        all_df["인증번호"] = all_df["인증번호"].astype(str).str.strip()
    return all_df


# ──────────────────────────────────────────────────────────────
# 시험결과요약서 표지 필드 추출
# ──────────────────────────────────────────────────────────────
LABEL_PATTERNS = {
    "업체명": [r"업체명"],
    "제조자": [r"제조자", r"제조사"],
    "제조국가": [r"제조국가"],
    "인증연월일": [r"인증\s*연월일"],
    "인증번호": [r"인증\s*번호"],
    "인증기준": [r"인증\s*기준\s*번호", r"인증기준\s*번호"],
    "인증범위": [r"인증\s*범위"],
}

# 라벨 뒤에 값이 바로 붙어 나오는 한 줄짜리 패턴들
# 예) "인증 번호 TTA-IT-C-26-007" / "제조국가: 대한민국" / "제조자/제조국가: 에이엠(주)/대한민국"
def _search_value(text, label_variants, combined_with=None):
    for lab in label_variants:
        # "라벨1/라벨2: 값1/값2" 형태 우선 시도
        if combined_with:
            pat = rf"{lab}\s*(?:및|/)\s*{combined_with}\s*[:\-]?\s*([^\n]+)"
            m = re.search(pat, text)
            if m:
                return m.group(1).strip(), True  # combined=True
        pat = rf"{lab}\s*[:\-]?\s*([^\n]+)"
        m = re.search(pat, text)
        if m:
            return m.group(1).strip(), False
    return None, False


def extract_test_highlights_from_page(page, x_split=195, y_tol=3.0):
    """표지가 2단 레이아웃(좌: 목차/ABOUT TTA, 우: Test Highlights 박스)인 문서에서
    x좌표 기준으로 우측 칼럼 단어만 모아 줄 단위로 재구성한 뒤,
    'Test Highlights' 다음부터 '개요' 헤딩 전까지를 하이라이트로 추출한다."""
    words = [w for w in page.extract_words() if w["x0"] >= x_split]
    if not words:
        return ""
    words.sort(key=lambda w: (round(w["top"] / y_tol), w["x0"]))

    lines = []
    cur_top, cur_words = None, []
    for w in words:
        if cur_top is None or abs(w["top"] - cur_top) <= y_tol:
            cur_words.append(w)
            cur_top = w["top"] if cur_top is None else cur_top
        else:
            lines.append((cur_top, cur_words))
            cur_words = [w]
            cur_top = w["top"]
    if cur_words:
        lines.append((cur_top, cur_words))

    started = False
    out_lines = []
    for _, lw in lines:
        lw_sorted = sorted(lw, key=lambda w: w["x0"])
        text = " ".join(w["text"] for w in lw_sorted)
        compact = re.sub(r"\s+", "", text)
        if not started:
            if "TestHighlights" in compact.replace(":", ""):
                # 같은 줄에 제목 뒤로 내용이 더 있으면 잘라서 시작
                m = re.search(r"Highlights\s*:?\s*(.*)", text)
                started = True
                if m and m.group(1).strip():
                    out_lines.append(m.group(1).strip())
            continue
        if compact in {"개요", "시험목적", "시험환경", "시험방법", "시험절차", "시험결과"}:
            break
        out_lines.append(text.strip())
    return "\n".join(l for l in out_lines if l)


@st.cache_data(show_spinner=False)
def extract_summary_fields(file_bytes, filename):
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
        tmp.write(file_bytes)
        tmp_path = tmp.name
    with pdfplumber.open(tmp_path) as pdf:
        page = pdf.pages[0]
        text = page.extract_text() or ""
        highlights = extract_test_highlights_from_page(page)

    result = {"__파일명": filename}

    # 제조자/제조국가가 한 줄로 붙어 있는 경우 우선 탐지
    combined_val, is_combined = _search_value(text, LABEL_PATTERNS["제조자"], combined_with="제조국가")
    manu, country = None, None
    if is_combined and combined_val:
        # "에이엠(주)/대한민국" 또는 "에이엠(주) / 대한민국" 둘 다 처리
        parts = re.split(r"\s*/\s*", combined_val)
        if len(parts) >= 2:
            manu, country = parts[0].strip(), parts[1].strip()

    if manu is None:
        manu, _ = _search_value(text, LABEL_PATTERNS["제조자"])
    if country is None:
        country, _ = _search_value(text, LABEL_PATTERNS["제조국가"])

    company, _ = _search_value(text, LABEL_PATTERNS["업체명"])
    if manu is None:
        manu = company  # 제조자 표기가 없으면 업체명으로 대체

    cert_no, _ = _search_value(text, LABEL_PATTERNS["인증번호"])
    cert_date, _ = _search_value(text, LABEL_PATTERNS["인증연월일"])
    criteria, _ = _search_value(text, LABEL_PATTERNS["인증기준"])
    scope, _ = _search_value(text, LABEL_PATTERNS["인증범위"])

    # 시험번호 (문서 상단 "No. TTA-26-XXXXX-TS00")
    m = re.search(r"No\.\s*(TTA-\d{2}-\d+)", text)
    test_no = m.group(1) if m else None

    result.update(
        {
            "업체명(추출)": company,
            "제조자(추출)": manu,
            "제조국가(추출)": country,
            "인증번호": cert_no,
            "시험번호": test_no,
            "인증연월일(추출)": cert_date,
            "인증기준(추출)": criteria,
            "인증범위(추출, 있는경우)": scope,
            "Test_Highlights": highlights,
        }
    )
    return result


# ──────────────────────────────────────────────────────────────
# 회의록(.hwp) 표 파싱 → 인증범위 매칭용
# ──────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def extract_minutes_table(file_bytes):
    """hwp5html로 변환 후 표를 대략적으로 파싱. 실패하면 빈 DataFrame 반환."""
    try:
        with tempfile.NamedTemporaryFile(suffix=".hwp", delete=False) as tmp:
            tmp.write(file_bytes)
            tmp_path = tmp.name
        out_dir = tempfile.mkdtemp()
        subprocess.run(
            ["hwp5html", "--output=" + out_dir, tmp_path],
            check=True, capture_output=True, timeout=60,
        )
        with open(out_dir + "/index.xhtml", encoding="utf-8") as f:
            html = f.read()

        from bs4 import BeautifulSoup
        soup = BeautifulSoup(html, "lxml")
        table = soup.find("table")
        if table is None:
            return pd.DataFrame()
        rows = []
        for tr in table.find_all("tr"):
            cells = [td.get_text(strip=True) for td in tr.find_all(["td", "th"])]
            if cells:
                rows.append(cells)
        if len(rows) < 2:
            return pd.DataFrame()
        header, body = rows[0], rows[1:]
        # 컬럼 수가 안 맞으면 잘라서 맞춤
        ncol = len(header)
        body = [r[:ncol] + [""] * (ncol - len(r)) for r in body]
        return pd.DataFrame(body, columns=header)
    except Exception as e:
        st.warning(f"회의록(.hwp) 파싱 실패 — 수동으로 입력해 주세요. ({e})")
        return pd.DataFrame()


def calc_expiry(date_str):
    """'2026. 8. 21.' 형태 문자열 → (시작일 문자열, 만료일 문자열)"""
    m = re.search(r"(\d{4})\D+(\d{1,2})\D+(\d{1,2})", date_str or "")
    if not m:
        return date_str, ""
    y, mo, d = map(int, m.groups())
    start = date(y, mo, d)
    end = start.replace(year=start.year + 5) - timedelta(days=1)
    fmt = lambda dt: f"{dt.year}. {dt.month}. {dt.day}."
    return fmt(start), fmt(end)


# ──────────────────────────────────────────────────────────────
# 메인 처리
# ──────────────────────────────────────────────────────────────
if mgmt_file and summary_files:
    mgmt_df = load_management_list(mgmt_file.getvalue())

    with st.expander("회의록 표 미리보기 (MMoIP 인증범위 참고용)"):
        minutes_df = pd.DataFrame()
        if minutes_file:
            minutes_df = extract_minutes_table(minutes_file.getvalue())
            st.dataframe(minutes_df, use_container_width=True)
        else:
            st.caption("회의록을 업로드하지 않았습니다.")

    rows = []
    for f in summary_files:
        extracted = extract_summary_fields(f.getvalue(), f.name)
        cert_no = extracted.get("인증번호")

        mgmt_row = None
        if cert_no and "인증번호" in mgmt_df.columns:
            match = mgmt_df[mgmt_df["인증번호"] == cert_no]
            if not match.empty:
                mgmt_row = match.iloc[0]

        업체명_국문 = mgmt_row["업체명"] if mgmt_row is not None and "업체명" in mgmt_row else extracted.get("업체명(추출)")
        업체명_영문 = mgmt_row["영문명"] if mgmt_row is not None and "영문명" in mgmt_row else ""
        제품명 = mgmt_row["제품명"] if mgmt_row is not None and "제품명" in mgmt_row else ""
        모델명 = mgmt_row["모델명"] if mgmt_row is not None and "모델명" in mgmt_row else ""
        파생모델명 = mgmt_row["파생모델명"] if mgmt_row is not None and "파생모델명" in mgmt_row else ""
        if pd.notna(파생모델명) and str(파생모델명).strip():
            모델명_전체 = f"{모델명}\n(파생모델명: {파생모델명})"
        else:
            모델명_전체 = 모델명
        인증기준 = mgmt_row["인증기준"] if mgmt_row is not None and "인증기준" in mgmt_row else extracted.get("인증기준(추출)")

        manu = extracted.get("제조자(추출)") or 업체명_국문
        country = extracted.get("제조국가(추출)") or ""
        제조자국가_국문 = f"{manu} / {country}".strip(" /")

        # 인증범위: 시험결과요약서에 있으면 그것을, 없으면 회의록에서 시험번호로 검색
        scope = extracted.get("인증범위(추출, 있는경우)")
        if not scope and minutes_file:
            minutes_df = extract_minutes_table(minutes_file.getvalue())
            if not minutes_df.empty:
                proj_col = next((c for c in minutes_df.columns if "프로젝트" in c or "시험번호" in c), None)
                scope_col = next((c for c in minutes_df.columns if "인증범위" in c), None)
                if proj_col and scope_col:
                    m = minutes_df[minutes_df[proj_col].astype(str).str.contains(extracted.get("시험번호") or "!!!", na=False)]
                    if not m.empty:
                        scope = m.iloc[0][scope_col]

        시작일, 만료일 = calc_expiry(extracted.get("인증연월일(추출)"))

        rows.append(
            {
                "인증번호": cert_no,
                "시험번호": extracted.get("시험번호"),
                "업체명_국문": 업체명_국문,
                "업체명_영문": 업체명_영문,
                "제품명": 제품명,
                "모델명": 모델명_전체,
                "제조자및제조국가_국문": 제조자국가_국문,
                "인증연월일": extracted.get("인증연월일(추출)"),
                "유효기간_시작일": 시작일,
                "유효기간_만료일": 만료일,
                "인증기준": 인증기준,
                "인증범위": scope,
                "Test_Highlights": extracted.get("Test_Highlights"),
                "__원본파일": extracted.get("__파일명"),
            }
        )

    result_df = pd.DataFrame(rows)
    result_df = result_df.sort_values("인증번호", na_position="last").reset_index(drop=True)
    result_df.insert(0, "순번", range(1, len(result_df) + 1))

    st.subheader("결과 (직접 수정 가능)")
    edited_df = st.data_editor(result_df, use_container_width=True, num_rows="dynamic", height=500)

    # 엑셀 다운로드
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        edited_df.to_excel(writer, index=False, sheet_name="인증서데이터")
    st.download_button(
        "📥 엑셀로 다운로드",
        data=buf.getvalue(),
        file_name="인증서_발급용_데이터.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("① 관리 목록과 ② 시험결과요약서를 업로드하면 결과가 표시됩니다.")
