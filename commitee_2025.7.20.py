import datetime
import sqlite3
import time
import zipfile
from io import BytesIO

import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="양산시 위원회 관리 시스템",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
<style>
@import url('https://fonts.googleapis.com/css2?family=SUIT:wght@400;500;700;800&display=swap');

:root {
    --bg-1: #f4f8f7;
    --bg-2: #e9f0ee;
    --ink: #0f2f2b;
    --muted: #55716c;
    --card: #ffffff;
    --line: #d9e6e3;
    --mint: #1d8f7a;
    --alert: #ef4444;
}

html, body, [class*="css"] {
    font-family: 'SUIT', 'Noto Sans KR', sans-serif;
    color: var(--ink);
}

.stApp {
    background:
        radial-gradient(1200px 500px at 0% -10%, #d7ebe6 0%, transparent 55%),
        radial-gradient(1000px 450px at 100% -20%, #f4e9d8 0%, transparent 55%),
        linear-gradient(180deg, var(--bg-1) 0%, var(--bg-2) 100%);
}

.block-container {
    padding-top: 2rem;
    padding-bottom: 2rem;
}

.hero {
    background: linear-gradient(135deg, #0f4f44, #1f6f60);
    border: 1px solid #2f7d70;
    color: white;
    padding: 1.4rem 1.6rem;
    border-radius: 18px;
    box-shadow: 0 10px 28px rgba(11, 54, 47, 0.2);
}

.hero h1 {
    margin: 0;
    font-size: 1.8rem;
    font-weight: 800;
}

.hero p {
    margin: 0.45rem 0 0 0;
    opacity: 0.9;
}

.kpi {
    background: var(--card);
    border: 1px solid var(--line);
    border-radius: 16px;
    padding: 1rem 1.1rem;
    box-shadow: 0 8px 20px rgba(20, 45, 40, 0.06);
}

.kpi-label {
    color: var(--muted);
    font-size: 0.85rem;
}

.kpi-value {
    font-size: 1.7rem;
    font-weight: 800;
    line-height: 1.2;
    margin-top: 0.2rem;
}

.section-title {
    margin-top: 0.5rem;
    margin-bottom: 0.6rem;
    font-size: 1.1rem;
    font-weight: 800;
    color: #174740;
}

.notice {
    background: #fff;
    border: 1px solid #f3d3d3;
    border-left: 6px solid var(--alert);
    border-radius: 12px;
    padding: 0.7rem 0.9rem;
}

.alert-center {
    background: #fffdf6;
    border: 1px solid #f2e5b8;
    border-left: 6px solid #e0a700;
    border-radius: 12px;
    padding: 0.8rem 1rem;
    margin-bottom: 0.7rem;
}

.profile-card {
    background: #ffffff;
    border: 1px solid var(--line);
    border-radius: 14px;
    padding: 1rem;
    box-shadow: 0 6px 16px rgba(20, 45, 40, 0.06);
}

div[data-testid="stDataFrame"] {
    border-radius: 12px;
    border: 1px solid var(--line);
    overflow: hidden;
}
</style>
""",
    unsafe_allow_html=True,
)


@st.cache_resource
def get_conn() -> sqlite3.Connection:
    return sqlite3.connect("organization.db", check_same_thread=False)


def init_tables() -> None:
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        """
        CREATE TABLE IF NOT EXISTS departments (
            순서 INTEGER,
            부서명 TEXT PRIMARY KEY,
            갱신일 TEXT
        )
        """
    )
    c.execute(
        """
        CREATE TABLE IF NOT EXISTS committees (
            순서 INTEGER,
            위원회명 TEXT PRIMARY KEY,
            부서명 TEXT,
            담당자 TEXT,
            연락처 TEXT,
            갱신일 TEXT,
            FOREIGN KEY(부서명) REFERENCES departments(부서명)
        )
        """
    )
    c.execute(
        """
        CREATE TABLE IF NOT EXISTS commissioners (
            순서 INTEGER,
            위원회명 TEXT,
            구분 TEXT CHECK(구분 IN ('당연직', '위촉직')),
            소속 TEXT,
            직위 TEXT,
            성명 TEXT,
            생년월일 TEXT,
            위촉일자 TEXT,
            만료일자 TEXT,
            임기 INTEGER,
            성별 TEXT,
            갱신일 TEXT,
            PRIMARY KEY (위원회명, 생년월일, 성명),
            FOREIGN KEY(위원회명) REFERENCES committees(위원회명)
        )
        """
    )
    c.execute(
        """
        CREATE TABLE IF NOT EXISTS admin_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            작업일시 TEXT,
            작업자 TEXT,
            작업유형 TEXT,
            대상 TEXT,
            상세 TEXT,
            건수 INTEGER
        )
        """
    )
    c.execute(
        """
        CREATE TABLE IF NOT EXISTS appointment_issues (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            발급번호 TEXT UNIQUE,
            발급일시 TEXT,
            부서명 TEXT,
            위원회명 TEXT,
            성명 TEXT,
            생년월일 TEXT,
            구분 TEXT,
            소속 TEXT,
            직위 TEXT,
            용도 TEXT,
            발급자 TEXT
        )
        """
    )
    conn.commit()


def fetch_df(query: str, params: tuple = ()) -> pd.DataFrame:
    return pd.read_sql_query(query, get_conn(), params=params)


def execute(query: str, params: tuple = ()) -> None:
    conn = get_conn()
    conn.execute(query, params)
    conn.commit()


def executemany(query: str, seq) -> None:
    conn = get_conn()
    conn.executemany(query, seq)
    conn.commit()


def to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Sheet1") -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()


def to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8-sig")


def log_action(actor: str, action_type: str, target: str, detail: str, count: int = 0) -> None:
    execute(
        """
        INSERT INTO admin_logs (작업일시, 작업자, 작업유형, 대상, 상세, 건수)
        VALUES (?, ?, ?, ?, ?, ?)
        """,
        (
            datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            actor.strip() if actor else "관리자",
            action_type,
            target,
            detail,
            int(count),
        ),
    )


def fetch_recent_logs(limit: int = 100) -> pd.DataFrame:
    return fetch_df(
        """
        SELECT 작업일시, 작업자, 작업유형, 대상, 상세, 건수
        FROM admin_logs
        ORDER BY id DESC
        LIMIT ?
        """,
        (limit,),
    )


def next_issue_numbers(count: int) -> list[str]:
    today = datetime.datetime.now().strftime("%Y%m%d")
    prefix = f"YS-{today}-"
    latest = fetch_df(
        """
        SELECT 발급번호
        FROM appointment_issues
        WHERE 발급번호 LIKE ?
        ORDER BY 발급번호 DESC
        LIMIT 1
        """,
        (f"{prefix}%",),
    )
    if latest.empty:
        start = 1
    else:
        last_no = str(latest["발급번호"].iloc[0]).split("-")[-1]
        start = int(last_no) + 1
    return [f"{prefix}{i:04d}" for i in range(start, start + count)]


def issue_numbers(selected_df: pd.DataFrame, issuer: str, purpose: str) -> pd.DataFrame:
    if selected_df.empty:
        return pd.DataFrame()
    numbers = next_issue_numbers(len(selected_df))
    issued_at = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    work = selected_df.copy().reset_index(drop=True)
    work["발급번호"] = numbers
    work["발급일시"] = issued_at
    work["용도"] = purpose.strip() if purpose else "위촉확인서"
    work["발급자"] = issuer.strip() if issuer else "담당자"

    insert_cols = ["발급번호", "발급일시", "부서명", "위원회명", "성명", "생년월일", "구분", "소속", "직위", "용도", "발급자"]
    executemany(
        """
        INSERT INTO appointment_issues
        (발급번호, 발급일시, 부서명, 위원회명, 성명, 생년월일, 구분, 소속, 직위, 용도, 발급자)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        [tuple(row) for row in work[insert_cols].itertuples(index=False, name=None)],
    )
    log_action(issuer, "발급번호 생성", "appointment_issues", work["용도"].iloc[0], len(work))
    return work[["발급번호", "발급일시", "부서명", "위원회명", "성명", "생년월일", "구분", "용도", "발급자"]]


def load_base_data() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    departments = fetch_df("SELECT 순서, 부서명, 갱신일 FROM departments ORDER BY 순서")
    committees = fetch_df(
        """
        SELECT c.순서, c.위원회명, c.부서명, c.담당자, c.연락처, c.갱신일
        FROM committees c
        LEFT JOIN departments d ON d.부서명 = c.부서명
        ORDER BY d.순서, c.순서, c.위원회명
        """
    )
    commissioners = fetch_df(
        """
        SELECT m.순서, c.부서명, m.위원회명, m.구분, m.소속, m.직위, m.성명,
               m.생년월일, m.위촉일자, m.만료일자, m.임기, m.성별, m.갱신일
        FROM commissioners m
        LEFT JOIN committees c ON c.위원회명 = m.위원회명
        """
    )
    return departments, committees, commissioners


def update_departments(df: pd.DataFrame, actor: str = "관리자") -> int:
    work = df.copy()
    work.columns = [str(col).strip() for col in work.columns]
    if "부서명" not in work.columns:
        raise ValueError("엑셀 파일에 [부서명] 컬럼이 있어야 합니다.")
    if "순서" not in work.columns:
        work["순서"] = range(1, len(work) + 1)

    work = work[["순서", "부서명"]].dropna(subset=["부서명"])
    work["부서명"] = work["부서명"].astype(str).str.strip()
    work = work[work["부서명"] != ""]
    work["순서"] = pd.to_numeric(work["순서"], errors="coerce").fillna(0).astype(int)
    work["갱신일"] = datetime.datetime.now().strftime("%Y-%m-%d")

    execute("DELETE FROM departments")
    executemany(
        "INSERT INTO departments (순서, 부서명, 갱신일) VALUES (?, ?, ?)",
        [tuple(row) for row in work.itertuples(index=False, name=None)],
    )
    log_action(actor, "부서 일괄갱신", "departments", "엑셀 업로드 반영", len(work))
    return len(work)


def delete_department(display_name: str, actor: str = "관리자") -> None:
    dept = display_name.split(". ")[-1].strip()
    execute("DELETE FROM departments WHERE 부서명 = ?", (dept,))
    log_action(actor, "부서 삭제", "departments", dept, 1)


def upsert_committees(df: pd.DataFrame, actor: str = "관리자") -> int:
    work = df.copy()
    work.columns = [str(col).strip() for col in work.columns]
    for col in ["부서명", "위원회명"]:
        if col not in work.columns:
            raise ValueError("엑셀 파일에는 [부서명, 위원회명] 컬럼이 모두 필요합니다.")

    for optional in ["순서", "담당자", "연락처"]:
        if optional not in work.columns:
            work[optional] = ""

    work = work[["순서", "부서명", "위원회명", "담당자", "연락처"]].dropna(subset=["부서명", "위원회명"])
    work["순서"] = pd.to_numeric(work["순서"], errors="coerce").fillna(0).astype(int)
    for col in ["부서명", "위원회명", "담당자", "연락처"]:
        work[col] = work[col].astype(str).str.strip()

    dept_df = fetch_df("SELECT 부서명 FROM departments")
    valid = set(dept_df["부서명"].tolist())
    work = work[work["부서명"].isin(valid)]
    today = datetime.datetime.now().strftime("%Y-%m-%d")

    q = """
    INSERT INTO committees (순서, 위원회명, 부서명, 담당자, 연락처, 갱신일)
    VALUES (?, ?, ?, ?, ?, ?)
    ON CONFLICT(위원회명) DO UPDATE SET
        순서=excluded.순서,
        부서명=excluded.부서명,
        담당자=excluded.담당자,
        연락처=excluded.연락처,
        갱신일=excluded.갱신일
    """
    rows = [
        (r.순서, r.위원회명, r.부서명, r.담당자, r.연락처, today)
        for r in work.itertuples(index=False)
    ]
    executemany(q, rows)
    log_action(actor, "위원회 일괄갱신", "committees", "엑셀 업로드 반영", len(rows))
    return len(rows)


def delete_committee(display_name: str, actor: str = "관리자") -> None:
    name = display_name.split("_")[-1].strip()
    execute("DELETE FROM committees WHERE 위원회명 = ?", (name,))
    log_action(actor, "위원회 삭제", "committees", name, 1)


def register_committee_members(df: pd.DataFrame, actor: str = "관리자") -> int:
    work = df.copy()
    work.columns = [str(col).strip() for col in work.columns]
    required = ["순서", "위원회명", "구분", "소속", "직위", "성명", "생년월일", "위촉일자", "만료일자", "임기", "성별"]
    missing = [col for col in required if col not in work.columns]
    if missing:
        raise ValueError(f"필수 컬럼 누락: {', '.join(missing)}")

    work = work[required].copy()
    work["순서"] = pd.to_numeric(work["순서"], errors="coerce").fillna(0).astype(int)
    work["임기"] = pd.to_numeric(work["임기"], errors="coerce")
    for dcol in ["생년월일", "위촉일자", "만료일자"]:
        work[dcol] = pd.to_datetime(work[dcol], errors="coerce").dt.strftime("%Y-%m-%d")

    for col in ["위원회명", "구분", "소속", "직위", "성명", "성별"]:
        work[col] = work[col].astype(str).str.strip()

    today = datetime.datetime.now().strftime("%Y-%m-%d")
    work["갱신일"] = today

    q = """
    INSERT INTO commissioners (순서, 위원회명, 구분, 소속, 직위, 성명, 생년월일, 위촉일자, 만료일자, 임기, 성별, 갱신일)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(위원회명, 생년월일, 성명) DO UPDATE SET
        순서=excluded.순서,
        구분=excluded.구분,
        소속=excluded.소속,
        직위=excluded.직위,
        위촉일자=excluded.위촉일자,
        만료일자=excluded.만료일자,
        임기=excluded.임기,
        성별=excluded.성별,
        갱신일=excluded.갱신일
    """

    rows = [tuple(row) for row in work[["순서", "위원회명", "구분", "소속", "직위", "성명", "생년월일", "위촉일자", "만료일자", "임기", "성별", "갱신일"]].itertuples(index=False, name=None)]
    executemany(q, rows)
    target = work["위원회명"].dropna().astype(str).iloc[0] if not work.empty else "미지정"
    log_action(actor, "위원 일괄갱신", "commissioners", target, len(rows))
    return len(rows)


def delete_members_by_committee(committee_name: str, actor: str = "관리자") -> None:
    old_cnt = fetch_df(
        "SELECT COUNT(*) AS cnt FROM commissioners WHERE 위원회명 = ?",
        (committee_name,),
    )
    cnt = int(old_cnt["cnt"].iloc[0]) if not old_cnt.empty else 0
    execute("DELETE FROM commissioners WHERE 위원회명 = ?", (committee_name,))
    log_action(actor, "위원회 위원 삭제", "commissioners", committee_name, cnt)


def generate_excel_for_committee(committee_name: str) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        cdf = fetch_df(
            "SELECT 순서, 위원회명, 부서명, 담당자, 연락처, 갱신일 FROM committees WHERE 위원회명 = ?",
            (committee_name,),
        )
        mdf = fetch_df(
            """
            SELECT 순서, 위원회명, 구분, 소속, 직위, 성명, 생년월일, 위촉일자, 만료일자, 임기, 성별
            FROM commissioners
            WHERE 위원회명 = ?
            ORDER BY 순서
            """,
            (committee_name,),
        )
        cdf.to_excel(writer, index=False, sheet_name="부서별_위원회_현황")
        mdf.to_excel(writer, index=False, sheet_name="위원회_명단")
    return output.getvalue()


def generate_zip_all_committees() -> bytes | None:
    base = fetch_df(
        """
        SELECT c.위원회명, d.순서, d.부서명
        FROM committees c
        LEFT JOIN departments d ON d.부서명 = c.부서명
        ORDER BY d.순서, c.순서, c.위원회명
        """
    )
    if base.empty:
        return None

    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for row in base.itertuples(index=False):
            name = row.위원회명
            dept = row.부서명 if pd.notna(row.부서명) else "미지정"
            seq = int(row.순서) if pd.notna(row.순서) else 0
            zf.writestr(f"{seq}.{dept}_{name}.xlsx", generate_excel_for_committee(name))

    zip_buffer.seek(0)
    return zip_buffer.getvalue()


def build_policy_tables(commissioners: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    appoint = commissioners[
        (commissioners["구분"] == "위촉직") & (commissioners["소속"] != "양산시의회")
    ].copy()

    dup_names = (
        appoint.groupby(["성명", "생년월일"], dropna=False)["위원회명"]
        .nunique()
        .reset_index(name="위원회수")
    )
    dup_names = dup_names[dup_names["위원회수"] >= 4]
    dup_table = appoint.merge(dup_names[["성명", "생년월일", "위원회수"]], on=["성명", "생년월일"], how="inner")
    dup_table = dup_table[["부서명", "위원회명", "성명", "생년월일", "위원회수"]].sort_values(["위원회수", "성명"], ascending=[False, True])

    g = appoint.groupby(["부서명", "위원회명"], dropna=False)
    gender = g["성별"].value_counts().unstack(fill_value=0).reset_index()
    if "남" not in gender.columns:
        gender["남"] = 0
    if "여" not in gender.columns:
        gender["여"] = 0
    gender["전체"] = gender["남"] + gender["여"]
    gender = gender[gender["전체"] > 0]
    gender["남성(%)"] = (gender["남"] / gender["전체"] * 100).round(1)
    gender["여성(%)"] = (gender["여"] / gender["전체"] * 100).round(1)
    gender_bias = gender[(gender["남성(%)"] > 60) | (gender["여성(%)"] > 60)][["부서명", "위원회명", "남성(%)", "여성(%)", "전체"]].sort_values("전체", ascending=False)

    today = pd.Timestamp.now().normalize()
    cut = today + pd.Timedelta(days=90)
    exp = appoint.copy()
    exp["만료일자"] = pd.to_datetime(exp["만료일자"], errors="coerce")
    expiring = exp[(exp["만료일자"].notna()) & (exp["만료일자"] <= cut)][["부서명", "위원회명", "성명", "만료일자", "직위", "소속"]].sort_values("만료일자")

    return dup_table, gender_bias, expiring


def build_alert_center(dup_table: pd.DataFrame, gender_bias: pd.DataFrame, expiring: pd.DataFrame) -> pd.DataFrame:
    today = pd.Timestamp.now().normalize()
    exp = expiring.copy()
    if not exp.empty:
        exp["만료일자"] = pd.to_datetime(exp["만료일자"], errors="coerce")
    else:
        exp["만료일자"] = pd.Series(dtype="datetime64[ns]")

    within_30 = int((exp["만료일자"] <= (today + pd.Timedelta(days=30))).sum())
    within_60 = int((exp["만료일자"] <= (today + pd.Timedelta(days=60))).sum())
    within_90 = int((exp["만료일자"] <= (today + pd.Timedelta(days=90))).sum())

    alerts = [
        ("높음", "중복 위촉 금지 대상", f"{len(dup_table):,}건", "위촉직 중 4개 이상 위원회 중복"),
        ("중간", "특정 성별 60% 초과", f"{len(gender_bias):,}개 위원회", "성별 구성 편중 점검 필요"),
        ("높음", "30일 내 만료", f"{within_30:,}명", "즉시 재위촉/교체 검토"),
        ("중간", "60일 내 만료", f"{within_60:,}명", "사전 검토 권장"),
        ("안내", "90일 내 만료", f"{within_90:,}명", "분기 내 갱신 대상"),
    ]
    return pd.DataFrame(alerts, columns=["우선순위", "항목", "현황", "안내"])


def render_home_tab(departments: pd.DataFrame, committees: pd.DataFrame, commissioners: pd.DataFrame) -> None:
    dup_table, gender_bias, expiring = build_policy_tables(commissioners)
    alert_center = build_alert_center(dup_table, gender_bias, expiring)

    st.markdown(
        f"""
<div class="hero">
  <h1>양산시 위원회 통합 현황</h1>
  <p>{datetime.datetime.now().strftime('%Y-%m-%d')} 기준으로 위원회 운영 현황과 정책 점검 결과를 한눈에 확인합니다.</p>
</div>
""",
        unsafe_allow_html=True,
    )
    st.write("")

    high_count = int((alert_center["우선순위"] == "높음").sum())
    mid_count = int((alert_center["우선순위"] == "중간").sum())
    st.markdown(
        f"""
<div class="alert-center">
  <b>자동 알림센터</b> · 높은 우선순위 {high_count}건, 중간 우선순위 {mid_count}건이 감지되었습니다.
</div>
""",
        unsafe_allow_html=True,
    )
    with st.expander("알림센터 상세 보기", expanded=False):
        st.dataframe(alert_center, hide_index=True, use_container_width=True, height=220)

    k1, k2, k3, k4 = st.columns(4)
    with k1:
        st.markdown(f'<div class="kpi"><div class="kpi-label">등록 부서</div><div class="kpi-value">{len(departments):,}</div></div>', unsafe_allow_html=True)
    with k2:
        st.markdown(f'<div class="kpi"><div class="kpi-label">등록 위원회</div><div class="kpi-value">{committees["위원회명"].nunique():,}</div></div>', unsafe_allow_html=True)
    with k3:
        st.markdown(f'<div class="kpi"><div class="kpi-label">전체 위원</div><div class="kpi-value">{len(commissioners):,}</div></div>', unsafe_allow_html=True)
    with k4:
        st.markdown(f'<div class="kpi"><div class="kpi-label">3개월 내 만료(위촉직)</div><div class="kpi-value">{len(expiring):,}</div></div>', unsafe_allow_html=True)

    st.write("")
    left, right = st.columns([1.15, 1])

    with left:
        st.markdown('<div class="section-title">부서별 위원회 분포</div>', unsafe_allow_html=True)
        dept_count = committees.groupby("부서명", dropna=False)["위원회명"].nunique().reset_index(name="위원회 수")
        dept_count = dept_count.sort_values("위원회 수", ascending=False).head(15)
        if not dept_count.empty:
            st.bar_chart(dept_count.set_index("부서명"), horizontal=True)
        else:
            st.info("표시할 위원회 데이터가 없습니다.")

        st.markdown('<div class="section-title">빠른 탐색</div>', unsafe_allow_html=True)
        col_a, col_b, col_c = st.columns([1.1, 1.1, 1.4])

        dept_options = ["전체"] + sorted([x for x in committees["부서명"].dropna().unique()])
        selected_dept = col_a.selectbox("부서", dept_options, label_visibility="collapsed", key="home_dept")

        sub = committees.copy()
        if selected_dept != "전체":
            sub = sub[sub["부서명"] == selected_dept]

        comm_options = ["전체"] + sorted([x for x in sub["위원회명"].dropna().unique()])
        selected_comm = col_b.selectbox("위원회", comm_options, label_visibility="collapsed", key="home_comm")
        keyword = col_c.text_input("검색", placeholder="위원 이름 검색", label_visibility="collapsed", key="home_kw")

        detail = commissioners.copy()
        if selected_dept != "전체":
            detail = detail[detail["부서명"] == selected_dept]
        if selected_comm != "전체":
            detail = detail[detail["위원회명"] == selected_comm]
        if keyword.strip():
            detail = detail[detail["성명"].fillna("").str.contains(keyword.strip(), na=False)]

        detail = detail[["부서명", "위원회명", "구분", "소속", "직위", "성명", "위촉일자", "만료일자", "성별"]].sort_values(["부서명", "위원회명", "성명"])
        st.dataframe(detail, use_container_width=True, hide_index=True, height=420)
        st.download_button(
            "현재 필터 결과 CSV 다운로드",
            data=to_csv_bytes(detail),
            file_name=f"위원_조회_{datetime.datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv",
            use_container_width=True,
            key="home_download",
        )

    with right:
        st.markdown('<div class="section-title">점검 및 알림</div>', unsafe_allow_html=True)
        st.markdown('<div class="notice"><b>중복 위촉위원(4개 이상)</b></div>', unsafe_allow_html=True)
        st.dataframe(dup_table, use_container_width=True, hide_index=True, height=180)

        st.markdown('<div class="notice"><b>특정 성별 60% 초과 위원회</b></div>', unsafe_allow_html=True)
        st.dataframe(gender_bias, use_container_width=True, hide_index=True, height=180)

        st.markdown('<div class="notice"><b>3개월 내 만료 대상(위촉직)</b></div>', unsafe_allow_html=True)
        exp_show = expiring.copy()
        if not exp_show.empty:
            exp_show["만료일자"] = pd.to_datetime(exp_show["만료일자"], errors="coerce").dt.strftime("%Y-%m-%d")
        st.dataframe(exp_show, use_container_width=True, hide_index=True, height=220)


def render_search_tab(departments: pd.DataFrame, committees: pd.DataFrame, commissioners: pd.DataFrame) -> None:
    st.subheader("부서별/위원회별 조회")
    dept_options = sorted([d for d in departments["부서명"].dropna().tolist()])
    if not dept_options:
        st.info("부서 데이터가 없습니다. 관리자페이지에서 먼저 등록해 주세요.")
        return

    c1, c2 = st.columns([2.5, 7.5])
    with c1:
        selected_dept = st.selectbox("부서", dept_options, label_visibility="collapsed", key="search_dept")
        left_df = committees[committees["부서명"] == selected_dept].drop_duplicates(subset="위원회명")
        st.dataframe(left_df[["순서", "부서명", "위원회명", "담당자", "연락처", "갱신일"]], hide_index=True, use_container_width=True, height=300)

    with c2:
        cand = commissioners[commissioners["부서명"] == selected_dept]
        comm_options = sorted([x for x in cand["위원회명"].dropna().unique()])
        if comm_options:
            selected_comm = st.selectbox("위원회", comm_options, label_visibility="collapsed", key="search_comm")
            right_df = cand[cand["위원회명"] == selected_comm].sort_values("순서")
            st.dataframe(right_df[["순서", "부서명", "위원회명", "구분", "소속", "직위", "성명", "생년월일", "위촉일자", "만료일자", "임기", "성별"]], hide_index=True, use_container_width=True, height=300)
        else:
            st.info("선택한 부서에 위원회/위원 데이터가 없습니다.")

    st.subheader("양산시 위원 이름 검색")
    kw = st.text_input("검색", placeholder="예: 홍길동", label_visibility="collapsed", key="search_name")
    if kw.strip():
        result = commissioners[commissioners["성명"].fillna("").str.contains(kw.strip(), na=False)][["부서명", "위원회명", "구분", "소속", "직위", "성명", "생년월일", "위촉일자", "만료일자", "임기", "성별"]]
        st.dataframe(result, hide_index=True, use_container_width=True, height=320)
        if not result.empty:
            options = (
                result[["성명", "생년월일"]]
                .drop_duplicates()
                .fillna("")
                .apply(lambda r: f"{r['성명']} ({r['생년월일']})", axis=1)
                .tolist()
            )
            picked = st.selectbox("이력 확인 대상", options, key="member_profile_pick")
            picked_name = picked.split(" (")[0].strip()
            picked_birth = picked.split("(")[-1].rstrip(")").strip()
            history = commissioners[
                (commissioners["성명"] == picked_name)
                & (commissioners["생년월일"].fillna("") == picked_birth)
            ].copy()
            history["만료일자_dt"] = pd.to_datetime(history["만료일자"], errors="coerce")
            today = pd.Timestamp.now().normalize()
            active_count = int((history["만료일자_dt"].isna() | (history["만료일자_dt"] >= today)).sum())
            expired_count = int((history["만료일자_dt"] < today).sum())
            latest_update = (
                history["갱신일"].dropna().astype(str).sort_values().iloc[-1]
                if not history["갱신일"].dropna().empty
                else "-"
            )
            st.markdown(
                f"""
<div class="profile-card">
  <b>위원 이력 카드</b><br/>
  성명: {picked_name} / 생년월일: {picked_birth}<br/>
  전체 참여 위원회: {history['위원회명'].nunique()}개 · 현재 활동: {active_count}건 · 만료: {expired_count}건 · 최근 갱신: {latest_update}
</div>
""",
                unsafe_allow_html=True,
            )
            history_show = history[
                ["부서명", "위원회명", "구분", "소속", "직위", "위촉일자", "만료일자", "임기", "성별"]
            ].sort_values(["부서명", "위원회명"])
            st.dataframe(history_show, hide_index=True, use_container_width=True, height=220)


def render_issue_tab(departments: pd.DataFrame, committees: pd.DataFrame, commissioners: pd.DataFrame) -> None:
    st.subheader("위촉위원 발급번호 생성")
    st.caption("위촉직 위원을 선택한 뒤 발급번호를 생성합니다. 형식: YS-YYYYMMDD-0001")

    left, right = st.columns([3, 2])
    with left:
        dept_options = ["전체"] + sorted([d for d in departments["부서명"].dropna().tolist()])
        selected_dept = st.selectbox("부서", dept_options, key="issue_dept")
        filtered_comm = committees.copy()
        if selected_dept != "전체":
            filtered_comm = filtered_comm[filtered_comm["부서명"] == selected_dept]

        comm_options = ["전체"] + sorted([x for x in filtered_comm["위원회명"].dropna().unique()])
        selected_comm = st.selectbox("위원회", comm_options, key="issue_comm")
        name_kw = st.text_input("이름 검색", placeholder="예: 홍길동", key="issue_name_kw")

    with right:
        issuer = st.text_input("발급자", value="담당자", key="issue_issuer")
        purpose = st.text_input("발급용도", value="위촉확인서", key="issue_purpose")

    candidates = commissioners[commissioners["구분"] == "위촉직"].copy()
    if selected_dept != "전체":
        candidates = candidates[candidates["부서명"] == selected_dept]
    if selected_comm != "전체":
        candidates = candidates[candidates["위원회명"] == selected_comm]
    if name_kw.strip():
        candidates = candidates[candidates["성명"].fillna("").str.contains(name_kw.strip(), na=False)]

    show_cols = ["부서명", "위원회명", "성명", "생년월일", "구분", "소속", "직위", "위촉일자", "만료일자"]
    base = candidates[show_cols].copy().sort_values(["부서명", "위원회명", "성명"])
    if base.empty:
        st.info("발급 가능한 위촉위원 데이터가 없습니다.")
    else:
        base.insert(0, "선택", False)
        edited = st.data_editor(
            base,
            hide_index=True,
            use_container_width=True,
            height=340,
            key="issue_editor",
            column_config={"선택": st.column_config.CheckboxColumn("선택")},
            disabled=show_cols,
        )
        selected = edited[edited["선택"]].copy()
        selected = selected.drop(columns=["선택"])

        if st.button("선택한 위원 발급번호 생성", use_container_width=True, key="issue_btn"):
            if selected.empty:
                st.warning("먼저 발급 대상을 선택해 주세요.")
            else:
                issued = issue_numbers(selected, issuer=issuer, purpose=purpose)
                st.success(f"발급번호 {len(issued)}건 생성 완료")
                st.dataframe(issued, hide_index=True, use_container_width=True, height=220)
                st.download_button(
                    "이번 생성 결과 CSV 다운로드",
                    data=to_csv_bytes(issued),
                    file_name=f"발급번호_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv",
                    use_container_width=True,
                    key="issue_download_new",
                )

    st.write("")
    st.subheader("발급 이력 조회")
    history = fetch_df(
        """
        SELECT 발급번호, 발급일시, 부서명, 위원회명, 성명, 생년월일, 구분, 용도, 발급자
        FROM appointment_issues
        ORDER BY id DESC
        LIMIT 500
        """
    )
    st.dataframe(history, hide_index=True, use_container_width=True, height=260)
    st.download_button(
        "발급 이력 CSV 다운로드",
        data=to_csv_bytes(history),
        file_name=f"발급이력_{datetime.datetime.now().strftime('%Y%m%d')}.csv",
        mime="text/csv",
        use_container_width=True,
        key="issue_download_history",
    )


def render_admin_tab(departments: pd.DataFrame, committees: pd.DataFrame, commissioners: pd.DataFrame) -> None:
    password = st.text_input("관리자 암호 입력", type="password")
    if password != "1234":
        st.info("관리자 암호를 입력하면 관리자 기능을 사용할 수 있습니다.")
        return

    st.title("관리자 기능")
    today = datetime.datetime.now().strftime("%Y-%m-%d")
    actor = st.text_input("작업자명", value="관리자", help="변경 이력에 기록됩니다.")
    col1, col2, col3 = st.columns([2.6, 3.1, 4.3])

    with col1:
        st.subheader("1. 부서현황 관리")
        st.dataframe(departments, hide_index=True, use_container_width=True, height=300)
        st.download_button(
            "부서정보 다운로드",
            data=to_excel_bytes(departments[["순서", "부서명"]], "부서현황"),
            file_name=f"부서현황_{today}.xlsx",
            use_container_width=True,
            key="admin_dept_download",
        )

        upload = st.file_uploader("부서 엑셀 업로드", type=["xlsx"], key="admin_dept_upload")
        if upload is not None and st.button("부서정보 반영하기", use_container_width=True):
            try:
                cnt = update_departments(pd.read_excel(upload), actor=actor)
                st.success(f"부서 {cnt}건 반영 완료")
                time.sleep(0.7)
                st.rerun()
            except Exception as e:
                st.error(str(e))

        dept_list = [f"{r['순서']}. {r['부서명']}" for _, r in departments.iterrows()]
        pick = st.selectbox("삭제할 부서", dept_list if dept_list else ["선택 가능한 부서 없음"], label_visibility="collapsed", key="admin_dept_delete_select")
        if st.button("선택 부서 삭제하기", use_container_width=True, disabled=not dept_list):
            delete_department(pick, actor=actor)
            st.success(f"{pick} 삭제 완료")
            time.sleep(0.7)
            st.rerun()

    with col2:
        st.subheader("2. 부서별 위원회 현황 관리")
        committee_view = committees[["순서", "부서명", "위원회명", "담당자", "연락처", "갱신일"]]
        st.dataframe(committee_view, hide_index=True, use_container_width=True, height=300)
        st.download_button(
            "양산시 위원회현황 다운로드",
            data=to_excel_bytes(committee_view, "부서별_위원회_현황"),
            file_name=f"양산시_위원회현황_{today}.xlsx",
            use_container_width=True,
            key="admin_comm_download",
        )

        c_upload = st.file_uploader("위원회 엑셀 업로드", type=["xlsx"], key="admin_comm_upload")
        if c_upload is not None and st.button("위원회자료 반영하기", use_container_width=True):
            try:
                xls = pd.ExcelFile(c_upload)
                target_sheet = "부서별_위원회_현황" if "부서별_위원회_현황" in xls.sheet_names else xls.sheet_names[0]
                cnt = upsert_committees(pd.read_excel(c_upload, sheet_name=target_sheet), actor=actor)
                st.success(f"위원회 {cnt}건 반영 완료")
                time.sleep(0.7)
                st.rerun()
            except Exception as e:
                st.error(str(e))

        comm_delete_list = [f"{r['순서']}. {r['부서명']}_{r['위원회명']}" for _, r in committee_view.iterrows()]
        c_pick = st.selectbox("삭제할 위원회", comm_delete_list if comm_delete_list else ["선택 가능한 위원회 없음"], label_visibility="collapsed", key="admin_comm_delete_select")
        if st.button("선택한 위원회 삭제하기", use_container_width=True, disabled=not comm_delete_list):
            delete_committee(c_pick, actor=actor)
            st.success(f"{c_pick} 삭제 완료")
            time.sleep(0.7)
            st.rerun()

    with col3:
        st.subheader("3. 위원회 위원 현황 관리")

        dept_options = sorted([d for d in departments["부서명"].dropna().tolist()])
        selected_dept = st.selectbox("부서 선택", dept_options if dept_options else [""], key="admin_member_dept")
        sub = commissioners[commissioners["부서명"] == selected_dept] if selected_dept else commissioners.iloc[0:0]
        comm_options = sorted([x for x in sub["위원회명"].dropna().unique()])
        selected_comm = st.selectbox("위원회 선택", comm_options if comm_options else [""], key="admin_member_comm")

        show = sub[sub["위원회명"] == selected_comm].sort_values("순서") if selected_comm else sub.iloc[0:0]
        st.dataframe(show[["순서", "구분", "소속", "직위", "성명", "생년월일", "위촉일자", "만료일자", "임기", "성별"]], hide_index=True, use_container_width=True, height=240)

        if selected_comm:
            st.download_button(
                "개별 위원회 현황 다운로드",
                data=generate_excel_for_committee(selected_comm),
                file_name=f"{selected_dept}({selected_comm})_{today}.xlsx",
                use_container_width=True,
                key="admin_member_excel_download",
            )

        m_upload = st.file_uploader("개별 위원회 엑셀 업로드", type=["xlsx"], key="admin_member_upload")
        if m_upload is not None and st.button("위원회 갱신자료 반영", use_container_width=True):
            try:
                uploaded = pd.read_excel(m_upload, sheet_name="위원회_명단")
                uploaded_comm = uploaded["위원회명"].dropna().astype(str).str.strip().unique().tolist()
                if not selected_comm:
                    raise ValueError("먼저 대상 위원회를 선택해 주세요.")
                if not uploaded_comm or uploaded_comm[0] != selected_comm:
                    raise ValueError("업로드 파일의 위원회명과 선택된 위원회명이 일치하지 않습니다.")
                delete_members_by_committee(selected_comm, actor=actor)
                cnt = register_committee_members(uploaded, actor=actor)
                st.success(f"{selected_comm} 위원 {cnt}건 갱신 완료")
                time.sleep(0.7)
                st.rerun()
            except Exception as e:
                st.error(str(e))

        if st.button("개별 위원회 현황 삭제하기", use_container_width=True, disabled=not selected_comm):
            delete_members_by_committee(selected_comm, actor=actor)
            st.success(f"{selected_comm} 위원 현황 삭제 완료")
            time.sleep(0.7)
            st.rerun()

        zip_data = generate_zip_all_committees()
        st.download_button(
            "부서별 위원회 갱신 파일 다운로드(zip)",
            data=zip_data if zip_data is not None else b"",
            file_name=f"부서별_위원회_위원회현황_{today}.zip",
            mime="application/zip",
            use_container_width=True,
            disabled=zip_data is None,
            key="admin_zip_download",
        )

    st.write("")
    st.subheader("최근 변경 이력")
    log_df = fetch_recent_logs(200)
    st.dataframe(log_df, hide_index=True, use_container_width=True, height=260)
    st.download_button(
        "변경 이력 CSV 다운로드",
        data=to_csv_bytes(log_df),
        file_name=f"변경이력_{today}.csv",
        mime="text/csv",
        use_container_width=True,
        key="admin_log_download",
    )


def main() -> None:
    init_tables()
    departments, committees, commissioners = load_base_data()

    left, right = st.columns([8, 2], vertical_alignment="bottom")
    with left:
        st.title("양산시 위원회 관리 시스템")
    with right:
        st.page_link(
            "https://www.elis.go.kr/allalr/selectAlrBdtOne?alrNo=48330102202009&histNo=006&menuNm=main",
            label="**양산시 각종 위원회 운영 조례**",
        )

    tab1, tab2, tab3, tab4 = st.tabs(["한눈에 보는 현황", "위원회 검색", "발급번호 생성", "관리자페이지"])
    with tab1:
        render_home_tab(departments, committees, commissioners)
    with tab2:
        render_search_tab(departments, committees, commissioners)
    with tab3:
        render_issue_tab(departments, committees, commissioners)
    with tab4:
        render_admin_tab(departments, committees, commissioners)


if __name__ == "__main__":
    main()
