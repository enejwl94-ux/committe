import streamlit as st
import pandas as pd
import sqlite3
from io import BytesIO
import datetime
from datetime import timedelta
import time
import zipfile
import html
import altair as alt

# 페이지 제목과 레이아웃 설정
st.set_page_config(layout="wide", page_title="양산시 위원회 관리 시스템", initial_sidebar_state="collapsed")

# 전역 데이터베이스 연결 설정 (organization-2.db 사용)
if 'conn' not in st.session_state:
    st.session_state.conn = sqlite3.connect("organization.db", check_same_thread=False)
    st.session_state.c = st.session_state.conn.cursor()

conn = st.session_state.conn
c = st.session_state.c

공통업무매뉴얼_KEY = "__GENERAL_COMMITTEE_MANUAL__"


def refresh_db_connection():
    """위원/위원회 갱신 직후 동일 세션에서 최신 DB 상태를 읽도록 연결을 재생성한다."""
    global conn, c

    existing_cursor = st.session_state.get("c")
    existing_conn = st.session_state.get("conn")

    try:
        if existing_cursor is not None:
            existing_cursor.close()
    except sqlite3.Error:
        pass

    try:
        if existing_conn is not None:
            existing_conn.close()
    except sqlite3.Error:
        pass

    st.session_state.conn = sqlite3.connect("organization.db", check_same_thread=False)
    st.session_state.c = st.session_state.conn.cursor()
    conn = st.session_state.conn
    c = st.session_state.c


def 위원_핵심지표_조회():
    query = """
        SELECT
            SUM(CASE WHEN COALESCE(TRIM(cm.성명), '') != '' OR cm.구분 = '위촉직' THEN 1 ELSE 0 END) AS 총위원수,
            SUM(CASE WHEN cm.구분 = '위촉직' THEN 1 ELSE 0 END) AS 위촉직수,
            SUM(CASE WHEN cm.구분 = '위촉직' AND cm.성별 = '여' THEN 1 ELSE 0 END) AS 위촉직_여성수
        FROM commissioners cm
        INNER JOIN committees c ON cm.위원회명 = c.위원회명
    """
    row = conn.execute(query).fetchone()
    if not row:
        return 0, 0, 0
    return tuple(int(value or 0) for value in row)


def normalize_excel_date_cell(value):
    """엑셀에서 읽은 날짜 셀을 DB 저장용 YYYY-MM-DD 문자열로 정리한다."""
    if pd.isna(value):
        return None

    if isinstance(value, str):
        value = value.strip()
        if not value:
            return None
        parsed = pd.to_datetime(value, errors="coerce")
        return parsed.strftime("%Y-%m-%d") if pd.notna(parsed) else value

    if isinstance(value, (pd.Timestamp, datetime.datetime, datetime.date)):
        return pd.Timestamp(value).strftime("%Y-%m-%d")

    if isinstance(value, (int, float)):
        parsed = pd.to_datetime(value, unit="D", origin="1899-12-30", errors="coerce")
        if pd.notna(parsed):
            return parsed.strftime("%Y-%m-%d")

    parsed = pd.to_datetime(value, errors="coerce")
    if pd.notna(parsed):
        return parsed.strftime("%Y-%m-%d")

    return str(value).strip() or None


def ensure_committee_columns():
    """
    기존 DB에 설치 관련 컬럼이 없을 경우 자동으로 추가
    """
    cur = conn.cursor()
    cur.execute("PRAGMA table_info(committees)")
    existing_columns = {row[1] for row in cur.fetchall()}
    cur.close()

    extra_columns = [
        ("설치일자", "TEXT"),
        ("설치근거", "TEXT"),
        ("법령상 의무설치 여부", "TEXT"),
        ("상설여부", "TEXT"),
        ("근거법령", "TEXT")
    ]

    for column_name, column_type in extra_columns:
        if column_name not in existing_columns:
            c.execute(f"ALTER TABLE committees ADD COLUMN `{column_name}` {column_type}")

    if "강행/임의" in existing_columns and "법령상 의무설치 여부" not in existing_columns:
        c.execute("""
            UPDATE committees
            SET `법령상 의무설치 여부` = `강행/임의`
            WHERE `강행/임의` IS NOT NULL
              AND TRIM(`강행/임의`) != ''
        """)
    conn.commit()


# 테이블 생성
def create_tables():
    c.execute("""
    CREATE TABLE IF NOT EXISTS departments (
        순서 INTEGER,
        부서명 TEXT PRIMARY KEY,
        갱신일 TEXT
    )
    """)

    c.execute("""
    CREATE TABLE IF NOT EXISTS committees (
        순서 INTEGER,
        위원회명 TEXT PRIMARY KEY,
        부서명 TEXT,
        설치일자 TEXT,
        설치근거 TEXT,
        `법령상 의무설치 여부` TEXT,
        상설여부 TEXT,
        근거법령 TEXT,
        담당자 TEXT,
        연락처 TEXT,
        갱신일 TEXT,
        FOREIGN KEY(부서명) REFERENCES departments(부서명)
    )
    """)

    c.execute("""
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
    """)

    c.execute("""
    CREATE TABLE IF NOT EXISTS committee_manuals (
        위원회명 TEXT PRIMARY KEY,
        파일명 TEXT NOT NULL,
        파일데이터 BLOB NOT NULL,
        파일형식 TEXT,
        파일크기 INTEGER,
        업로드일 TEXT,
        FOREIGN KEY(위원회명) REFERENCES committees(위원회명)
    )
    """)

    c.execute("""
    CREATE TABLE IF NOT EXISTS committee_yearly_stats (
        연도 INTEGER,
        위원회명 TEXT,
        대면회의 INTEGER DEFAULT 0,
        서면회의 INTEGER DEFAULT 0,
        운영경비천원 INTEGER DEFAULT 0,
        갱신일 TEXT,
        PRIMARY KEY (연도, 위원회명),
        FOREIGN KEY(위원회명) REFERENCES committees(위원회명)
    )
    """)
    conn.commit()
    ensure_committee_columns()


create_tables()

# DataFrame을 엑셀 파일로 변환
def to_excel(df, sheet_name="Sheet1"):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()


def fetch_data(query, params=None, columns=None):
    cur = conn.cursor()
    try:
        if params is None:
            params = ()
        elif isinstance(params, (list, tuple)):
            params = tuple(params)
        else:
            params = (params,)
        cur.execute(query, params)
        rows = cur.fetchall()
    finally:
        cur.close()
    return pd.DataFrame(rows, columns=columns or [])


기준일_고정값 = pd.Timestamp("2026-01-01").normalize()


def 기준일_옵션_목록():
    return [
        ("기준일", 기준일_고정값, "base"),
        ("1개월 전", 기준일_고정값 - pd.DateOffset(months=1), "one_month"),
        ("6개월 전", 기준일_고정값 - pd.DateOffset(months=6), "six_months"),
        ("1년 전", 기준일_고정값 - pd.DateOffset(years=1), "one_year"),
    ]


def 시의원_위촉현황_원본데이터():
    council_members_query = """
    SELECT
        d.부서명,
        c.위원회명,
        cm.성명,
        cm.구분,
        cm.직위,
        cm.생년월일,
        cm.위촉일자,
        cm.만료일자,
        cm.임기,
        cm.성별,
        cm.소속
    FROM commissioners cm
    JOIN committees c ON cm.위원회명 = c.위원회명
    LEFT JOIN departments d ON c.부서명 = d.부서명
    WHERE cm.소속 LIKE '%양산시의회%'
    ORDER BY d.순서, c.위원회명, cm.성명
    """
    df = pd.read_sql_query(council_members_query, conn)

    if df.empty:
        return pd.DataFrame(columns=[
            "부서명", "위원회명", "성명", "구분", "직위",
            "생년월일", "위촉일자", "만료일자", "임기", "성별", "소속"
        ])

    for column in ["생년월일", "위촉일자", "만료일자"]:
        df[column] = pd.to_datetime(df[column], errors="coerce")

    return df


def 시의원_위촉현황_데이터(기준일):
    df = 시의원_위촉현황_원본데이터()
    if df.empty:
        return df

    기준시점 = pd.Timestamp(기준일).normalize()
    active_mask = (
        (df["위촉일자"].isna() | (df["위촉일자"] <= 기준시점)) &
        (df["만료일자"].isna() | (df["만료일자"] >= 기준시점))
    )
    df = df.loc[active_mask].copy()

    for column in ["생년월일", "위촉일자", "만료일자"]:
        df[column] = df[column].dt.strftime("%Y-%m-%d")
        df[column] = df[column].where(df[column].notna(), None)

    return df.reset_index(drop=True)


def 시의원_위촉현황_요약(df_council):
    columns = ["부서명", "위원회명", "시의원수", "명단"]
    if df_council.empty:
        return pd.DataFrame(columns=columns)

    summary = (
        df_council.groupby(["부서명", "위원회명"], dropna=False)
        .agg(
            시의원수=("성명", "nunique"),
            명단=("성명", lambda names: ", ".join(sorted(name for name in names if isinstance(name, str))))
        )
        .reset_index()
        .sort_values(by=["부서명", "위원회명"], na_position="last")
    )
    return summary


def 시의원_위촉현황_다운로드_파일(df_council, summary, 기준일, 기준라벨):
    info_df = pd.DataFrame([{
        "기준구분": 기준라벨,
        "기준일": pd.Timestamp(기준일).strftime("%Y-%m-%d"),
        "산정방식": "위촉일자 <= 기준일 <= 만료일자(만료일자 공란 포함)"
    }])

    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        info_df.to_excel(writer, index=False, sheet_name="기준정보")
        df_council.to_excel(writer, index=False, sheet_name="개별명단")
        summary.to_excel(writer, index=False, sheet_name="위원회별현황")

        writer.sheets["기준정보"].set_column("A:C", 28)
        writer.sheets["개별명단"].set_column("A:K", 16)
        writer.sheets["위원회별현황"].set_column("A:D", 24)

    return output.getvalue()


def 기준일_위원회_위촉현황_데이터(부서명=None, 위원회명=None, 기준일=None):
    query = """
    SELECT
        m.순서,
        d.부서명,
        c.위원회명,
        m.구분,
        m.소속,
        m.직위,
        m.성명,
        m.생년월일,
        m.위촉일자,
        m.만료일자,
        m.임기,
        m.성별
    FROM committees AS c
    LEFT JOIN departments AS d ON c.부서명 = d.부서명
    LEFT JOIN commissioners AS m ON c.위원회명 = m.위원회명
    WHERE 1 = 1
    """
    params = []

    if 부서명:
        query += " AND d.부서명 = ?"
        params.append(부서명)
    if 위원회명:
        query += " AND c.위원회명 = ?"
        params.append(위원회명)

    query += " ORDER BY d.순서, c.위원회명, m.순서, m.성명"

    columns = ["순서", "부서명", "위원회명", "구분", "소속", "직위", "성명", "생년월일", "위촉일자", "만료일자", "임기", "성별"]
    df = fetch_data(query, params=params, columns=columns)

    if df.empty:
        return df

    name_series = df["성명"].astype("string").fillna("").str.strip()
    df = df[name_series.ne("") | df["구분"].eq("위촉직")].copy()
    if df.empty:
        return df

    for column in ["생년월일", "위촉일자", "만료일자"]:
        df[column] = pd.to_datetime(df[column], errors="coerce")

    if 기준일 is not None:
        기준시점 = pd.Timestamp(기준일).normalize()
        active_mask = (
            (df["위촉일자"].isna() | (df["위촉일자"] <= 기준시점)) &
            (df["만료일자"].isna() | (df["만료일자"] >= 기준시점))
        )
        df = df.loc[active_mask].copy()

    if df.empty:
        return df

    for column in ["생년월일", "위촉일자", "만료일자"]:
        df[column] = df[column].dt.strftime("%Y-%m-%d")
        df[column] = df[column].where(df[column].notna(), None)

    return df.reset_index(drop=True)


def 회의운영현황_컬럼정리(df):
    df = df.copy()
    rename_map = {
        "대면회의횟수": "대면회의",
        "서면회의횟수": "서면회의",
        "운영경비": "운영경비(천원)",
        "운영경비천원": "운영경비(천원)",
    }
    applicable_map = {
        old_name: new_name
        for old_name, new_name in rename_map.items()
        if old_name in df.columns and new_name not in df.columns
    }
    if applicable_map:
        df = df.rename(columns=applicable_map)
    return df


def 회의운영현황_입력양식():
    return pd.DataFrame(columns=["연도", "위원회명", "대면회의", "서면회의", "운영경비(천원)"])


def 회의운영현황_기준연도(df=None):
    current_year = datetime.datetime.now().year
    default_years = list(range(current_year - 3, current_year + 1))

    if df is None or df.empty or "연도" not in df.columns:
        return default_years

    years = pd.to_numeric(df["연도"], errors="coerce").dropna().astype(int).tolist()
    if not years:
        return default_years

    return sorted(set(default_years).union(years))


def 회의운영현황_와이드컬럼(years):
    columns = ["연번", "소관부서", "위원회명"]
    for year in years:
        columns.extend([f"{year}년_합계", f"{year}년_서면", f"{year}년_대면"])
    for year in years:
        columns.append(f"{year}년_운영경비(천원)")
    return columns


def 회의운영현황_기본행(df_long=None):
    base_query = """
        SELECT
            d.순서,
            c.부서명 AS 소관부서,
            c.위원회명
        FROM committees AS c
        LEFT JOIN departments AS d ON c.부서명 = d.부서명
        ORDER BY d.순서, c.부서명, c.위원회명
    """
    base_df = fetch_data(base_query, columns=["순서", "소관부서", "위원회명"])

    extra_df = pd.DataFrame(columns=["순서", "소관부서", "위원회명"])
    if df_long is not None and not df_long.empty and "위원회명" in df_long.columns:
        extra_df = df_long.copy()
        if "부서명" in extra_df.columns:
            extra_df = extra_df.rename(columns={"부서명": "소관부서"})
        if "소관부서" not in extra_df.columns:
            extra_df["소관부서"] = None
        extra_df["순서"] = pd.NA
        extra_df = extra_df[["순서", "소관부서", "위원회명"]]

    combined_df = pd.concat([base_df, extra_df], ignore_index=True)
    if combined_df.empty:
        return pd.DataFrame(columns=["연번", "소관부서", "위원회명"])

    combined_df["위원회명"] = combined_df["위원회명"].astype("string").fillna("").str.strip()
    combined_df["소관부서"] = combined_df["소관부서"].astype("string").fillna("").str.strip()
    combined_df = combined_df[combined_df["위원회명"].ne("")].copy()
    combined_df["소관부서"] = combined_df["소관부서"].replace("", "미등록")
    combined_df["정렬순서"] = pd.to_numeric(combined_df["순서"], errors="coerce")
    combined_df = (
        combined_df
        .sort_values(["정렬순서", "소관부서", "위원회명"], na_position="last")
        .drop_duplicates(subset=["위원회명"], keep="first")
        .reset_index(drop=True)
    )
    combined_df["연번"] = range(1, len(combined_df) + 1)
    return combined_df[["연번", "소관부서", "위원회명"]]


def 회의운영현황_와이드데이터(df_long=None):
    years = 회의운영현황_기준연도(df_long)
    wide_df = 회의운영현황_기본행(df_long)

    for year in years:
        wide_df[f"{year}년_합계"] = pd.NA
        wide_df[f"{year}년_서면"] = pd.NA
        wide_df[f"{year}년_대면"] = pd.NA
        wide_df[f"{year}년_운영경비(천원)"] = pd.NA

    if df_long is None or df_long.empty:
        return wide_df[회의운영현황_와이드컬럼(years)]

    df_long = df_long.copy()
    if "부서명" in df_long.columns and "소관부서" not in df_long.columns:
        df_long = df_long.rename(columns={"부서명": "소관부서"})

    numeric_columns = ["연도", "대면회의", "서면회의", "운영경비(천원)"]
    for column in numeric_columns:
        if column in df_long.columns:
            df_long[column] = pd.to_numeric(df_long[column], errors="coerce")

    if "위원회명" not in df_long.columns:
        return wide_df[회의운영현황_와이드컬럼(years)]

    df_long["위원회명"] = df_long["위원회명"].astype("string").fillna("").str.strip()
    df_long = df_long[df_long["위원회명"].ne("")].copy()

    for year in years:
        year_df = df_long[df_long["연도"] == year].copy()
        if year_df.empty:
            continue

        year_df[f"{year}년_합계"] = year_df["서면회의"].fillna(0) + year_df["대면회의"].fillna(0)
        year_df = year_df.rename(
            columns={
                "서면회의": f"{year}년_서면",
                "대면회의": f"{year}년_대면",
                "운영경비(천원)": f"{year}년_운영경비(천원)",
            }
        )
        year_df = (
            year_df[["위원회명", f"{year}년_합계", f"{year}년_서면", f"{year}년_대면", f"{year}년_운영경비(천원)"]]
            .drop_duplicates(subset=["위원회명"], keep="last")
        )
        wide_df = wide_df.merge(year_df, on="위원회명", how="left", suffixes=("", "__new"))

        for column in [f"{year}년_합계", f"{year}년_서면", f"{year}년_대면", f"{year}년_운영경비(천원)"]:
            new_column = f"{column}__new"
            if new_column in wide_df.columns:
                wide_df[column] = wide_df[new_column].combine_first(wide_df[column])
                wide_df = wide_df.drop(columns=[new_column])

    return wide_df[회의운영현황_와이드컬럼(years)]


def 회의운영현황_와이드헤더_파싱(df_raw):
    if df_raw.shape[0] < 4:
        return None

    header_top = df_raw.iloc[0].ffill()
    header_middle = df_raw.iloc[1].ffill()
    header_bottom = df_raw.iloc[2]
    data_df = df_raw.iloc[3:].reset_index(drop=True).copy()

    columns = []
    for idx in range(data_df.shape[1]):
        top = "" if pd.isna(header_top.iloc[idx]) else str(header_top.iloc[idx]).strip()
        middle = "" if pd.isna(header_middle.iloc[idx]) else str(header_middle.iloc[idx]).strip()
        bottom = "" if pd.isna(header_bottom.iloc[idx]) else str(header_bottom.iloc[idx]).strip()

        if bottom in {"연번", "소관부서", "위원회명"}:
            columns.append(bottom)
        elif "회의 개최 횟수" in top and middle:
            columns.append(f"{middle}_{bottom}")
        elif "운영경비" in top and middle:
            columns.append(f"{middle}_운영경비(천원)")
        else:
            columns.append(bottom or middle or top or f"col_{idx}")

    data_df.columns = columns
    data_df = data_df.dropna(axis=1, how="all")
    return data_df


def 회의운영현황_와이드에서_변환(df):
    df = df.copy()
    df.columns = [str(column).strip() for column in df.columns]
    if "1. 위원회명" in df.columns and "위원회명" not in df.columns:
        df = df.rename(columns={"1. 위원회명": "위원회명"})
    if "부서명" in df.columns and "소관부서" not in df.columns:
        df = df.rename(columns={"부서명": "소관부서"})

    years = sorted(
        {
            int(column.split("년_")[0])
            for column in df.columns
            if "년_" in column and column.split("년_")[0].isdigit()
        }
    )

    if not years or "위원회명" not in df.columns:
        return 회의운영현황_컬럼정리(df)

    records = []
    for _, row in df.iterrows():
        committee_name = row.get("위원회명")
        if pd.isna(committee_name):
            continue
        committee_name = str(committee_name).strip()
        if not committee_name:
            continue

        department_name = row.get("소관부서")
        department_name = None if pd.isna(department_name) else str(department_name).strip()

        for year in years:
            raw_total = row.get(f"{year}년_합계")
            raw_document = row.get(f"{year}년_서면")
            raw_face_to_face = row.get(f"{year}년_대면")
            raw_budget = row.get(f"{year}년_운영경비(천원)")

            raw_values = [raw_total, raw_document, raw_face_to_face, raw_budget]
            if all(pd.isna(value) or str(value).strip() == "" for value in raw_values):
                continue

            document_count = pd.to_numeric(pd.Series([raw_document]), errors="coerce").iloc[0]
            face_to_face_count = pd.to_numeric(pd.Series([raw_face_to_face]), errors="coerce").iloc[0]
            budget_amount = pd.to_numeric(pd.Series([raw_budget]), errors="coerce").iloc[0]

            records.append(
                {
                    "연도": year,
                    "부서명": department_name,
                    "위원회명": committee_name,
                    "대면회의": 0 if pd.isna(face_to_face_count) else int(round(face_to_face_count)),
                    "서면회의": 0 if pd.isna(document_count) else int(round(document_count)),
                    "운영경비(천원)": 0 if pd.isna(budget_amount) else int(round(budget_amount)),
                }
            )

    return pd.DataFrame(records, columns=["연도", "부서명", "위원회명", "대면회의", "서면회의", "운영경비(천원)"])


def 회의운영현황_엑셀읽기(uploaded_file):
    try:
        workbook = pd.ExcelFile(uploaded_file)
        sheet_name = "입력양식" if "입력양식" in workbook.sheet_names else workbook.sheet_names[0]

        df_raw = pd.read_excel(workbook, sheet_name=sheet_name, header=None)
        wide_df = 회의운영현황_와이드헤더_파싱(df_raw)
        if wide_df is not None and "위원회명" in wide_df.columns:
            parsed_df = 회의운영현황_와이드에서_변환(wide_df)
            if not parsed_df.empty:
                return parsed_df

        df = pd.read_excel(workbook, sheet_name=sheet_name)
        return 회의운영현황_컬럼정리(df)
    except Exception as e:
        st.error(f"회의 운영현황 파일을 읽는 중 오류가 발생했습니다: {e}")
        return 회의운영현황_입력양식()


def 파일크기_표시(size):
    if size in (None, "") or pd.isna(size):
        return "-"

    value = float(size)
    units = ["B", "KB", "MB", "GB"]
    unit_index = 0
    while value >= 1024 and unit_index < len(units) - 1:
        value /= 1024
        unit_index += 1

    if unit_index == 0:
        return f"{int(value)} {units[unit_index]}"
    return f"{value:.1f} {units[unit_index]}"


def 업무매뉴얼_mime_type(file_name):
    lower_name = str(file_name).lower()
    if lower_name.endswith(".hwp"):
        return "application/haansofthwp"
    if lower_name.endswith(".hwpx"):
        return "application/octet-stream"
    return "application/octet-stream"


def 공통_업무매뉴얼_저장(uploaded_file):
    if uploaded_file is None:
        return False

    file_name = str(uploaded_file.name)
    lower_name = file_name.lower()
    if not lower_name.endswith((".hwp", ".hwpx")):
        st.error("한글 파일은 .hwp 또는 .hwpx 형식만 업로드할 수 있습니다.")
        return False

    file_bytes = uploaded_file.getvalue()
    if not file_bytes:
        st.error("업로드된 파일을 읽을 수 없습니다.")
        return False

    uploaded_at = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    mime_type = 업무매뉴얼_mime_type(file_name)

    try:
        c.execute(
            """
            INSERT INTO committee_manuals (위원회명, 파일명, 파일데이터, 파일형식, 파일크기, 업로드일)
            VALUES (?, ?, ?, ?, ?, ?)
            ON CONFLICT(위원회명) DO UPDATE SET
                파일명 = excluded.파일명,
                파일데이터 = excluded.파일데이터,
                파일형식 = excluded.파일형식,
                파일크기 = excluded.파일크기,
                업로드일 = excluded.업로드일
            """,
            (공통업무매뉴얼_KEY, file_name, file_bytes, mime_type, len(file_bytes), uploaded_at)
        )
        conn.commit()
        return True
    except Exception as e:
        st.error(f"업무 매뉴얼 저장 중 오류가 발생했습니다: {e}")
        return False


def 공통_업무매뉴얼_파일():
    cur = conn.cursor()
    try:
        cur.execute(
            """
            SELECT 위원회명, 파일명, 파일데이터, 파일형식, 파일크기, 업로드일
            FROM committee_manuals
            WHERE 위원회명 = ?
            """,
            (공통업무매뉴얼_KEY,)
        )
        row = cur.fetchone()
    finally:
        cur.close()

    if not row:
        return None

    return {
        "위원회명": row[0],
        "파일명": row[1],
        "파일데이터": row[2],
        "파일형식": row[3] or 업무매뉴얼_mime_type(row[1]),
        "파일크기": row[4],
        "업로드일": row[5],
    }


def 부서현황_갱신(table_name, df, key_column):
    try:
        c.execute(f"DELETE FROM {table_name}")
        today = datetime.datetime.now().strftime("%Y-%m-%d")
        df['갱신일'] = today
        fields = ", ".join(df.columns)
        placeholders = ", ".join(["?"] * len(df.columns))
        insert_sql = f"INSERT INTO {table_name} ({fields}) VALUES ({placeholders})"

        for _, row in df.iterrows():
            values = [None if pd.isna(value) else value for value in row.tolist()]
            c.execute(insert_sql, tuple(values))

        conn.commit()
        return True
    except Exception as e:
        st.error(f"{table_name} 데이터 업로드 중 오류: {e}")
        return False


def 부서_삭제(부서명):
    try:
        실제_부서명 = 부서명.split(". ")[-1]
        delete_sql = "DELETE FROM departments WHERE 부서명 = ?"
        c.execute(delete_sql, (실제_부서명,))
        conn.commit()
    except Exception as e:
        st.error(f"부서 삭제 중 오류: {e}")



        

def 부서별_위원회_갱신(df):
    required_columns = {"위원회명", "부서명"}
    if not required_columns.issubset(df.columns):
        st.error("부서명과 위원회명 컬럼이 모두 포함되어야 합니다.")
        return 0

    today = datetime.datetime.now().strftime("%Y-%m-%d")
    df = df.copy()
    df["위원회명"] = df["위원회명"].astype("string").fillna("").str.strip()
    df["부서명"] = df["부서명"].astype("string").fillna("").str.strip()
    df = df[df["위원회명"].ne("") & df["부서명"].ne("")].copy()
    if "설치일자" in df.columns:
        df["설치일자"] = df["설치일자"].apply(normalize_excel_date_cell)
    df['갱신일'] = today

    if "강행/임의" in df.columns and "법령상 의무설치 여부" not in df.columns:
        df = df.rename(columns={"강행/임의": "법령상 의무설치 여부"})

    optional_columns = ["순서", "설치일자", "설치근거", "법령상 의무설치 여부", "상설여부", "근거법령", "담당자", "연락처"]
    insert_columns = ["위원회명", "부서명", "갱신일"]
    insert_columns += [col for col in optional_columns if col in df.columns and col not in insert_columns]
    insert_columns = list(dict.fromkeys(insert_columns))

    update_columns = [col for col in insert_columns if col != "위원회명"]

    placeholders = ", ".join(["?"] * len(insert_columns))
    fields = ", ".join([f"`{col}`" for col in insert_columns])
    update_sql = ", ".join([f"`{col}`=excluded.`{col}`" for col in update_columns])

    insert_sql = f"""
        INSERT INTO committees ({fields})
        VALUES ({placeholders})
        ON CONFLICT(위원회명) DO UPDATE SET
            {update_sql}
    """

    처리_카운터 = 0
    try:
        for _, row in df.iterrows():
            insert_values = []
            for col in insert_columns:
                value = row.get(col)
                if pd.isna(value):
                    value = None
                insert_values.append(value)

            c.execute(insert_sql, tuple(insert_values))
            처리_카운터 += 1

        conn.commit()
        return 처리_카운터
    except Exception as e:
        st.error(f"committees 데이터 업로드 중 오류: {e}")
        return 처리_카운터


def 연도별_회의운영_갱신(df):
    df = 회의운영현황_컬럼정리(df)
    required_columns = {"연도", "위원회명", "대면회의", "서면회의", "운영경비(천원)"}
    if not required_columns.issubset(df.columns):
        st.error("엑셀 파일에는 [연도, 위원회명, 대면회의, 서면회의, 운영경비(천원)] 컬럼이 모두 포함되어야 합니다.")
        return 0

    today = datetime.datetime.now().strftime("%Y-%m-%d")
    df = df.copy()
    df["위원회명"] = df["위원회명"].astype("string").fillna("").str.strip()
    df["연도"] = pd.to_numeric(df["연도"], errors="coerce")

    numeric_columns = ["대면회의", "서면회의", "운영경비(천원)"]
    for column in numeric_columns:
        df[column] = pd.to_numeric(df[column], errors="coerce").fillna(0)

    df = df[df["연도"].notna() & df["위원회명"].ne("")].copy()
    if df.empty:
        st.error("반영할 수 있는 유효한 데이터가 없습니다. 연도와 위원회명을 확인해 주세요.")
        return 0

    df["연도"] = df["연도"].round(0).astype(int)
    for column in numeric_columns:
        df[column] = df[column].round(0).astype(int)
    df["갱신일"] = today

    insert_sql = """
        INSERT INTO committee_yearly_stats (연도, 위원회명, 대면회의, 서면회의, 운영경비천원, 갱신일)
        VALUES (?, ?, ?, ?, ?, ?)
        ON CONFLICT(연도, 위원회명) DO UPDATE SET
            대면회의 = excluded.대면회의,
            서면회의 = excluded.서면회의,
            운영경비천원 = excluded.운영경비천원,
            갱신일 = excluded.갱신일
    """

    처리_카운터 = 0
    try:
        for _, row in df.iterrows():
            c.execute(
                insert_sql,
                (
                    row["연도"],
                    row["위원회명"],
                    row["대면회의"],
                    row["서면회의"],
                    row["운영경비(천원)"],
                    row["갱신일"],
                )
            )
            처리_카운터 += 1

        conn.commit()
        return 처리_카운터
    except Exception as e:
        st.error(f"연도별 회의 운영현황 반영 중 오류가 발생했습니다: {e}")
        return 처리_카운터

# def 부서별_위원회_갱신(df):
#     처리_카운터 = 0
#     try:
#         today = datetime.datetime.now().strftime("%Y-%m-%d")

#         for col in ['설치일자', '설치근거', '근거법령']:
#             if col not in df.columns:
#                 df[col] = None

#         df['갱신일'] = today

#         insert_columns = ["순서", "위원회명", "부서명", "설치일자", "설치근거", "근거법령", "담당자", "연락처", "갱신일"]
#         placeholders = ", ".join(["?"] * len(insert_columns))
#         fields = ", ".join(insert_columns)

#         updatable_columns = ["순서", "부서명", "설치일자", "설치근거", "근거법령", "담당자", "연락처", "갱신일"]
#         update_sql = ", ".join([f"{col}=excluded.{col}" for col in updatable_columns])

#         insert_sql = f"""
#         INSERT INTO committees ({fields})
#         VALUES ({placeholders})
#         ON CONFLICT(위원회명) DO UPDATE SET
#             {update_sql}
#         """

#         for _, row in df.iterrows():
#             insert_values = []
#             for col in insert_columns:
#                 value = row.get(col)
#                 if pd.isna(value):
#                     value = None
#                 insert_values.append(value)

#             c.execute(insert_sql, tuple(insert_values))
#             처리_카운터 += 1

#         conn.commit()
#         return 처리_카운터
#     except Exception as e:
#         st.error(f"committees 데이터 업로드 중 오류: {e}")
#         return 처리_카운터

def 위원회_구성_데이터():
    query = """
        SELECT
            c.부서명,
            c.위원회명,
            SUM(CASE WHEN COALESCE(TRIM(cm.성명), '') != '' OR cm.구분 = '위촉직' THEN 1 ELSE 0 END) AS 총계,
            SUM(CASE WHEN cm.구분 = '당연직' THEN 1 ELSE 0 END) AS 당연직,
            SUM(CASE WHEN cm.구분 = '위촉직' THEN 1 ELSE 0 END) AS 위촉직,
            SUM(CASE WHEN cm.구분 = '당연직' AND cm.성별 = '남' THEN 1 ELSE 0 END) AS 당연직_남성,
            SUM(CASE WHEN cm.구분 = '당연직' AND cm.성별 = '여' THEN 1 ELSE 0 END) AS 당연직_여성,
            SUM(CASE WHEN cm.구분 = '위촉직' AND cm.성별 = '남' THEN 1 ELSE 0 END) AS 위촉직_남성,
            SUM(CASE WHEN cm.구분 = '위촉직' AND cm.성별 = '여' THEN 1 ELSE 0 END) AS 위촉직_여성
        FROM committees c
        LEFT JOIN commissioners cm ON c.위원회명 = cm.위원회명
        GROUP BY c.부서명, c.위원회명
        ORDER BY c.부서명, c.위원회명
    """
    columns = [
        "부서명", "위원회명", "총계", "당연직", "위촉직",
        "당연직_남성", "당연직_여성", "위촉직_남성", "위촉직_여성"
    ]
    return fetch_data(query, columns=columns)

def 위원회_성별_비율():
    query = """
        SELECT
            c.부서명,
            c.위원회명,
            SUM(CASE WHEN cm.구분 = '위촉직' AND cm.성별 = '남' THEN 1 ELSE 0 END) AS 위촉직_남성,
            SUM(CASE WHEN cm.구분 = '위촉직' AND cm.성별 = '여' THEN 1 ELSE 0 END) AS 위촉직_여성
        FROM committees c
        LEFT JOIN commissioners cm ON c.위원회명 = cm.위원회명
        GROUP BY c.부서명, c.위원회명
        HAVING 위촉직_남성 + 위촉직_여성 > 0
        ORDER BY c.부서명, c.위원회명
    """
    columns = ["부서명", "위원회명", "위촉직_남성", "위촉직_여성"]
    df = fetch_data(query, columns=columns)
    df_long = df.melt(
        id_vars=["부서명", "위원회명"],
        value_vars=["위촉직_남성", "위촉직_여성"],
        var_name="성별",
        value_name="인원"
    )
    df_long["성별"] = df_long["성별"].map({"위촉직_남성": "남", "위촉직_여성": "여"})
    return df_long

def 위원회_구성_집계():
    query = """
        SELECT
            COALESCE(d.순서, c.순서) AS 정렬순서,
            c.부서명,
            c.위원회명,
            SUM(CASE WHEN COALESCE(TRIM(cm.성명), '') != '' OR cm.구분 = '위촉직' THEN 1 ELSE 0 END) AS 총계,
            SUM(CASE WHEN cm.구분 = '당연직' THEN 1 ELSE 0 END) AS 당연직,
            SUM(CASE WHEN cm.구분 = '위촉직' THEN 1 ELSE 0 END) AS 위촉직,
            SUM(CASE WHEN cm.구분 = '위촉직' AND cm.성별 = '남' THEN 1 ELSE 0 END) AS 위촉직_남성,
            SUM(CASE WHEN cm.구분 = '위촉직' AND cm.성별 = '여' THEN 1 ELSE 0 END) AS 위촉직_여성,
            ROUND(
                CASE WHEN SUM(CASE WHEN cm.구분 = '위촉직' THEN 1 ELSE 0 END) = 0 THEN 0
                ELSE SUM(CASE WHEN cm.구분 = '위촉직' AND cm.성별 = '남' THEN 1 ELSE 0 END)
                     * 100.0 / SUM(CASE WHEN cm.구분 = '위촉직' THEN 1 ELSE 0 END)
                END
            , 1) AS 위촉직_남성_비율,
            ROUND(
                CASE WHEN SUM(CASE WHEN cm.구분 = '위촉직' THEN 1 ELSE 0 END) = 0 THEN 0
                ELSE SUM(CASE WHEN cm.구분 = '위촉직' AND cm.성별 = '여' THEN 1 ELSE 0 END)
                     * 100.0 / SUM(CASE WHEN cm.구분 = '위촉직' THEN 1 ELSE 0 END)
                END
            , 1) AS 위촉직_여성_비율
        FROM committees c
        LEFT JOIN departments d ON c.부서명 = d.부서명
        LEFT JOIN commissioners cm ON c.위원회명 = cm.위원회명
        GROUP BY COALESCE(d.순서, c.순서), c.부서명, c.위원회명
        ORDER BY COALESCE(d.순서, c.순서), c.위원회명
    """
    columns = [
        "정렬순서", "부서명", "위원회명", "총계", "당연직", "위촉직",
        "위촉직_남성", "위촉직_여성", "위촉직_남성_비율", "위촉직_여성_비율"
    ]
    df = fetch_data(query, columns=columns)
    if df.empty:
        return df.drop(columns=["정렬순서"], errors="ignore")
    return (
        df.sort_values(["정렬순서", "위원회명"], na_position="last")
        .drop(columns=["정렬순서"])
        .reset_index(drop=True)
    )


def 위원회_기본정보_데이터():
    query = """
        SELECT
            순서,
            부서명,
            위원회명,
            설치일자,
            설치근거,
            `법령상 의무설치 여부`,
            상설여부,
            근거법령,
            담당자,
            연락처,
            갱신일
        FROM committees
        ORDER BY 순서, 부서명, 위원회명
    """
    columns = ["순서", "부서명", "위원회명", "설치일자", "설치근거", "법령상 의무설치 여부", "상설여부", "근거법령", "담당자", "연락처", "갱신일"]
    df = fetch_data(query, columns=columns)
    if not df.empty and "설치일자" in df.columns:
        df["설치일자"] = df["설치일자"].apply(normalize_excel_date_cell)
    return df


def 부서_대시보드_요약():
    query = """
        SELECT
            c.부서명,
            COUNT(DISTINCT c.위원회명) AS 위원회수,
            SUM(CASE WHEN COALESCE(TRIM(cm.성명), '') != '' OR cm.구분 = '위촉직' THEN 1 ELSE 0 END) AS 위원수,
            SUM(CASE WHEN cm.구분 = '위촉직' THEN 1 ELSE 0 END) AS 위촉직수,
            ROUND(
                CASE WHEN SUM(CASE WHEN cm.구분 = '위촉직' THEN 1 ELSE 0 END) = 0 THEN 0
                ELSE SUM(CASE WHEN cm.구분 = '위촉직' AND cm.성별 = '여' THEN 1 ELSE 0 END)
                     * 100.0 / SUM(CASE WHEN cm.구분 = '위촉직' THEN 1 ELSE 0 END)
                END
            , 1) AS 위촉직_여성비율
        FROM committees c
        LEFT JOIN commissioners cm ON c.위원회명 = cm.위원회명
        GROUP BY c.부서명
        ORDER BY 위원회수 DESC, 위원수 DESC, c.부서명
    """
    columns = ["부서명", "위원회수", "위원수", "위촉직수", "위촉직_여성비율"]
    return fetch_data(query, columns=columns)


def 회의운영현황_데이터():
    query = """
        SELECT
            COALESCE(d.순서, c.순서) AS 순서,
            COALESCE(d.부서명, c.부서명, '미등록') AS 부서명,
            y.위원회명,
            y.연도,
            y.대면회의,
            y.서면회의,
            y.대면회의 + y.서면회의 AS 총회의,
            y.운영경비천원 AS "운영경비(천원)",
            y.갱신일
        FROM committee_yearly_stats AS y
        LEFT JOIN committees AS c ON y.위원회명 = c.위원회명
        LEFT JOIN departments AS d ON c.부서명 = d.부서명
        ORDER BY y.연도 DESC, COALESCE(d.순서, c.순서), y.위원회명
    """
    columns = ["순서", "부서명", "위원회명", "연도", "대면회의", "서면회의", "총회의", "운영경비(천원)", "갱신일"]
    df = fetch_data(query, columns=columns)

    if df.empty:
        return df

    numeric_columns = ["연도", "대면회의", "서면회의", "총회의", "운영경비(천원)"]
    for column in numeric_columns:
        df[column] = pd.to_numeric(df[column], errors="coerce").fillna(0).astype(int)

    df["부서명"] = df["부서명"].fillna("미등록")
    df["갱신일"] = pd.to_datetime(df["갱신일"], errors="coerce").dt.date.astype("string")
    return df


def 연도별_회의운영_요약(df):
    columns = ["연도", "위원회수", "대면회의", "서면회의", "총회의", "운영경비(천원)"]
    if df.empty:
        return pd.DataFrame(columns=columns)

    summary = (
        df.groupby("연도", as_index=False)
        .agg(
            위원회수=("위원회명", "nunique"),
            대면회의=("대면회의", "sum"),
            서면회의=("서면회의", "sum"),
            총회의=("총회의", "sum"),
            운영경비_천원=("운영경비(천원)", "sum"),
        )
        .sort_values("연도")
        .reset_index(drop=True)
    )
    summary = summary.rename(columns={"운영경비_천원": "운영경비(천원)"})
    return summary[columns]


def 중복_위촉위원_데이터():
    query = """
        SELECT
            commissioners.성명,
            commissioners.생년월일,
            COUNT(DISTINCT commissioners.위원회명) AS 참여위원회수,
            GROUP_CONCAT(DISTINCT committees.부서명) AS 소관부서,
            GROUP_CONCAT(DISTINCT commissioners.위원회명) AS 위원회목록
        FROM commissioners
        JOIN committees ON commissioners.위원회명 = committees.위원회명
        WHERE commissioners.구분 = '위촉직'
          AND commissioners.소속 != '양산시의회'
          AND commissioners.소속 != '양산시'
        GROUP BY commissioners.성명, commissioners.생년월일
        HAVING COUNT(DISTINCT commissioners.위원회명) >= 4
        ORDER BY 참여위원회수 DESC, commissioners.성명
    """
    columns = ["성명", "생년월일", "참여위원회수", "소관부서", "위원회목록"]
    return fetch_data(query, columns=columns)


def 성별_편중_위원회_데이터():
    query = """
        SELECT
            COALESCE(d.순서, c.순서) AS 정렬순서,
            c.부서명,
            cm.위원회명,
            COUNT(*) AS 위촉직수,
            ROUND(SUM(CASE WHEN cm.성별 = '남' THEN 1 ELSE 0 END) * 100.0 / COUNT(*), 1) AS 남성비율,
            ROUND(SUM(CASE WHEN cm.성별 = '여' THEN 1 ELSE 0 END) * 100.0 / COUNT(*), 1) AS 여성비율
        FROM commissioners cm
        JOIN committees c ON cm.위원회명 = c.위원회명
        LEFT JOIN departments d ON c.부서명 = d.부서명
        WHERE cm.구분 = '위촉직'
        GROUP BY COALESCE(d.순서, c.순서), c.부서명, cm.위원회명
        HAVING 남성비율 > 60 OR 여성비율 > 60
        ORDER BY COALESCE(d.순서, c.순서), cm.위원회명
    """
    columns = ["정렬순서", "부서명", "위원회명", "위촉직수", "남성비율", "여성비율"]
    df = fetch_data(query, columns=columns)
    if df.empty:
        return df.drop(columns=["정렬순서"], errors="ignore")
    return (
        df.sort_values(["정렬순서", "위원회명"], na_position="last")
        .drop(columns=["정렬순서"])
        .reset_index(drop=True)
    )


def 임기만료_예정_데이터(days=90):
    start_date = datetime.datetime.now().strftime("%Y-%m-%d")
    end_date = (datetime.datetime.now() + timedelta(days=days)).strftime("%Y-%m-%d")
    query = """
        SELECT
            c.부서명,
            c.위원회명,
            cm.성명,
            cm.만료일자,
            c.담당자,
            c.연락처
        FROM committees AS c
        JOIN commissioners AS cm ON c.위원회명 = cm.위원회명
        WHERE cm.구분 = '위촉직'
          AND cm.만료일자 >= ?
          AND cm.만료일자 <= ?
        ORDER BY cm.만료일자, c.부서명, c.위원회명
    """
    columns = ["부서명", "위원회명", "성명", "만료일자", "담당자", "연락처"]
    return fetch_data(query, params=(start_date, end_date), columns=columns)


def 청년위원회_비율_데이터(df_committees_status_1, df_committees_status_2):
    target_committees = [
        '시보운영위원회',
        '시민통합위원회',
        '청년정책위원회',
        '노사민정협의회',
        '학교폭력대책지역협의회',
        '보육정책위원회',
        '정보공개심의회',
        '청원심의회'
    ]

    committee_department = df_committees_status_1[
        df_committees_status_1['위원회명'].isin(target_committees)
    ][['위원회명', '부서명']].drop_duplicates()

    df_target = df_committees_status_2[
        df_committees_status_2['위원회명'].isin(target_committees)
        & df_committees_status_2['구분'].eq('위촉직')
    ].copy()

    if df_target.empty:
        return pd.DataFrame(columns=['부서명', '위원회명', '전체 위원', '청년 위원', '비율(%)', '상태'])

    df_target['생년월일'] = pd.to_datetime(df_target['생년월일'], errors='coerce')
    df_target = df_target.dropna(subset=['생년월일'])

    today = pd.Timestamp.now().normalize()
    df_target['나이'] = today.year - df_target['생년월일'].dt.year - (
        (today.month < df_target['생년월일'].dt.month) |
        (
            (today.month == df_target['생년월일'].dt.month) &
            (today.day < df_target['생년월일'].dt.day)
        )
    ).astype(int)

    total_members = df_target.groupby('위원회명').size()
    young_members = df_target[
        (df_target['나이'] >= 19) & (df_target['나이'] <= 39)
    ].groupby('위원회명').size()
    percentage = (young_members / total_members * 100).round(2)

    df_age_ratio = pd.DataFrame({'위원회명': target_committees})
    df_age_ratio = df_age_ratio.merge(committee_department, on='위원회명', how='left')
    df_age_ratio = df_age_ratio.merge(total_members.rename('전체 위원'), on='위원회명', how='left')
    df_age_ratio = df_age_ratio.merge(young_members.rename('청년 위원'), on='위원회명', how='left')
    df_age_ratio = df_age_ratio.merge(percentage.rename('비율(%)'), on='위원회명', how='left')
    df_age_ratio[['전체 위원', '청년 위원', '비율(%)']] = df_age_ratio[
        ['전체 위원', '청년 위원', '비율(%)']
    ].fillna(0)
    df_age_ratio['전체 위원'] = df_age_ratio['전체 위원'].astype(int)
    df_age_ratio['청년 위원'] = df_age_ratio['청년 위원'].astype(int)
    df_age_ratio['상태'] = df_age_ratio['비율(%)'].apply(lambda value: '유의' if value < 10 else '충족')
    return df_age_ratio[['부서명', '위원회명', '전체 위원', '청년 위원', '비율(%)', '상태']]


def 스타일_적용():
    st.markdown(
        """
        <style>
        .stApp {
            background:
                radial-gradient(circle at top left, rgba(20, 184, 166, 0.12), transparent 28%),
                radial-gradient(circle at top right, rgba(14, 116, 144, 0.10), transparent 24%),
                linear-gradient(180deg, #f7fbfb 0%, #edf5f4 100%);
        }
        [data-testid="block-container"] {
            padding-top: 1.75rem;
            padding-bottom: 3rem;
            max-width: 1400px;
        }
        h1, h2, h3, h4, h5, h6, p, span, div, label {
            font-family: "Segoe UI", "Apple SD Gothic Neo", "Noto Sans KR", sans-serif;
        }
        .stTabs [data-baseweb="tab-list"] {
            gap: 0.55rem;
            padding: 0.2rem 0 1rem 0;
        }
        .stTabs [data-baseweb="tab"] {
            height: 3rem;
            border-radius: 999px;
            padding: 0 1rem;
            background: rgba(255, 255, 255, 0.78);
            border: 1px solid rgba(15, 23, 42, 0.08);
            color: #0f172a;
        }
        .stTabs [aria-selected="true"] {
            background: linear-gradient(135deg, #0f766e 0%, #155e75 100%);
            color: white;
            box-shadow: 0 10px 24px rgba(21, 94, 117, 0.24);
        }
        .hero-panel {
            padding: 1.8rem 2rem;
            border-radius: 28px;
            background: linear-gradient(135deg, #0f172a 0%, #155e75 48%, #0f766e 100%);
            color: white;
            box-shadow: 0 24px 50px rgba(15, 23, 42, 0.18);
            margin-bottom: 1.15rem;
        }
        .hero-kicker {
            display: inline-block;
            padding: 0.38rem 0.72rem;
            border-radius: 999px;
            background: rgba(255, 255, 255, 0.14);
            font-size: 0.82rem;
            letter-spacing: 0.02em;
            margin-bottom: 0.9rem;
        }
        .hero-title {
            font-size: 2.2rem;
            font-weight: 700;
            margin: 0;
        }
        .hero-copy {
            max-width: 820px;
            margin: 0.8rem 0 0;
            color: rgba(255, 255, 255, 0.84);
            line-height: 1.7;
        }
        .hero-meta {
            display: flex;
            flex-wrap: wrap;
            gap: 0.55rem;
            margin-top: 1rem;
        }
        .hero-pill {
            padding: 0.42rem 0.78rem;
            border-radius: 999px;
            background: rgba(255, 255, 255, 0.12);
            font-size: 0.84rem;
        }
        .metric-card {
            padding: 1.15rem 1.2rem;
            border-radius: 22px;
            background: rgba(255, 255, 255, 0.9);
            border: 1px solid rgba(15, 23, 42, 0.07);
            box-shadow: 0 16px 35px rgba(15, 23, 42, 0.06);
            min-height: 132px;
        }
        .metric-label {
            font-size: 0.88rem;
            color: #475569;
            margin-bottom: 0.5rem;
        }
        .metric-value {
            font-size: 2rem;
            font-weight: 700;
            color: #0f172a;
            line-height: 1.1;
        }
        .metric-help {
            margin-top: 0.55rem;
            font-size: 0.88rem;
            color: #64748b;
            line-height: 1.5;
        }
        .section-title {
            margin-top: 1.5rem;
            margin-bottom: 0.15rem;
            font-size: 1.15rem;
            font-weight: 700;
            color: #0f172a;
        }
        .section-copy {
            color: #64748b;
            margin-bottom: 0.6rem;
        }
        .panel-title {
            margin: 0 0 0.45rem 0;
            font-size: 1.02rem;
            font-weight: 700;
            color: #0f172a;
        }
        .panel-copy {
            color: #64748b;
            font-size: 0.88rem;
            margin-bottom: 0.6rem;
        }
        .info-card {
            padding: 1.15rem 1.25rem;
            border-radius: 22px;
            background: rgba(255, 255, 255, 0.9);
            border: 1px solid rgba(15, 23, 42, 0.07);
            box-shadow: 0 16px 35px rgba(15, 23, 42, 0.05);
        }
        .quickview-info-card {
            min-height: 344px;
        }
        .info-card h4 {
            margin: 0 0 0.75rem 0;
            font-size: 1.05rem;
            color: #0f172a;
        }
        .info-grid {
            display: grid;
            grid-template-columns: repeat(2, minmax(0, 1fr));
            gap: 0.75rem;
            margin-top: 0.85rem;
        }
        .info-item {
            padding: 0.75rem 0.85rem;
            border-radius: 16px;
            background: #f8fafc;
            border: 1px solid rgba(148, 163, 184, 0.14);
        }
        .info-item-label {
            font-size: 0.8rem;
            color: #64748b;
            margin-bottom: 0.25rem;
        }
        .info-item-value {
            font-size: 1rem;
            font-weight: 600;
            color: #0f172a;
            word-break: keep-all;
        }
        .table-caption {
            color: #64748b;
            font-size: 0.9rem;
            margin-bottom: 0.35rem;
        }
        div[data-testid="stDataFrame"] {
            border-radius: 18px;
            overflow: visible;
            border: 1px solid rgba(15, 23, 42, 0.06);
            box-shadow: 0 10px 24px rgba(15, 23, 42, 0.04);
        }
        .scroll-table-card {
            border-radius: 18px;
            overflow: hidden;
            border: 1px solid rgba(15, 23, 42, 0.06);
            box-shadow: 0 10px 24px rgba(15, 23, 42, 0.04);
            background: rgba(255, 255, 255, 0.92);
        }
        .scroll-table-wrap {
            overflow-y: auto;
            overflow-x: hidden;
        }
        .scroll-table {
            width: 100%;
            border-collapse: collapse;
            table-layout: fixed;
            background: white;
        }
        .scroll-table thead th {
            position: sticky;
            top: 0;
            z-index: 1;
            background: #f8fafc;
            white-space: normal;
        }
        .scroll-table th,
        .scroll-table td {
            padding: 0.58rem 0.72rem;
            border-bottom: 1px solid rgba(148, 163, 184, 0.18);
            font-size: 0.9rem;
            color: #0f172a;
            text-align: left;
            vertical-align: top;
            line-height: 1.4;
            white-space: normal;
            word-break: break-word;
            overflow-wrap: anywhere;
        }
        .scroll-table th:nth-child(1),
        .scroll-table td:nth-child(1) {
            width: 62%;
        }
        .scroll-table th:nth-child(2),
        .scroll-table td:nth-child(2) {
            width: 20%;
        }
        .scroll-table th:nth-child(3),
        .scroll-table td:nth-child(3) {
            width: 18%;
        }
        .scroll-table tbody tr:last-child td {
            border-bottom: none;
        }
        div[data-testid="stVegaLiteChart"] {
            border-radius: 22px;
            overflow: hidden;
            border: 1px solid rgba(15, 23, 42, 0.07);
            box-shadow: 0 16px 35px rgba(15, 23, 42, 0.05);
            background: rgba(255, 255, 255, 0.92);
            padding: 0.7rem 0.85rem 0.35rem 0.85rem;
        }
        .summary-table-card {
            padding: 0.45rem;
            border-radius: 18px;
            background: rgba(255, 255, 255, 0.9);
            border: 1px solid rgba(15, 23, 42, 0.07);
            box-shadow: 0 10px 24px rgba(15, 23, 42, 0.05);
        }
        .summary-table {
            width: 100%;
            border-collapse: collapse;
            table-layout: fixed;
            background: white;
        }
        .summary-table th,
        .summary-table td {
            border: 1px solid #475569;
            padding: 0.5rem 0.35rem;
            text-align: center;
            font-size: 0.9rem;
            color: #0f172a;
            word-break: keep-all;
        }
        .summary-table thead .summary-table-title th {
            background: #e9c2c2;
            font-size: 0.98rem;
            font-weight: 700;
            padding: 0.42rem 0.35rem;
        }
        .summary-table thead .summary-table-header th {
            background: #f8f5f1;
            font-weight: 700;
        }
        .summary-table tbody th {
            background: #f8f5f1;
            font-weight: 700;
        }
        .summary-table-note {
            margin-top: 0.45rem;
            font-size: 0.8rem;
            color: #64748b;
        }
        </style>
        """,
        unsafe_allow_html=True
    )


def metric_card(title, value, description):
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-label">{html.escape(str(title))}</div>
            <div class="metric-value">{html.escape(str(value))}</div>
            <div class="metric-help">{html.escape(str(description))}</div>
        </div>
        """,
        unsafe_allow_html=True
    )


def info_panel(title, info_pairs, extra_class=""):
    items = "".join(
        f"""
        <div class="info-item">
            <div class="info-item-label">{html.escape(str(label))}</div>
            <div class="info-item-value">{html.escape(str(value or '-'))}</div>
        </div>
        """
        for label, value in info_pairs
    )
    class_name = "info-card"
    if extra_class:
        class_name = f"{class_name} {html.escape(str(extra_class))}"
    st.markdown(
        f"""
        <div class="{class_name}">
            <h4>{html.escape(str(title))}</h4>
            <div class="info-grid">{items}</div>
        </div>
        """,
        unsafe_allow_html=True
    )


def summary_table_panel(title, columns, rows, note=None):
    header_cells = "".join(f"<th>{html.escape(str(column))}</th>" for column in columns)
    body_rows = ""
    for row in rows:
        row_header = html.escape(str(row[0]))
        row_cells = "".join(f"<td>{html.escape(str(cell))}</td>" for cell in row[1:])
        body_rows += f"<tr><th scope='row'>{row_header}</th>{row_cells}</tr>"

    note_html = f"<div class='summary-table-note'>{html.escape(str(note))}</div>" if note else ""
    st.markdown(
        f"""
        <div class="summary-table-card">
            <table class="summary-table">
                <thead>
                    <tr class="summary-table-title">
                        <th colspan="{len(columns)}">{html.escape(str(title))}</th>
                    </tr>
                    <tr class="summary-table-header">
                        {header_cells}
                    </tr>
                </thead>
                <tbody>
                    {body_rows}
                </tbody>
            </table>
            {note_html}
        </div>
        """,
        unsafe_allow_html=True
    )


def scroll_table_panel(df, columns, height=260):
    header_cells = "".join(f"<th>{html.escape(str(column))}</th>" for column in columns)
    body_rows = ""
    for _, row in df[columns].iterrows():
        row_cells = "".join(
            f"<td>{html.escape('-' if pd.isna(value) else str(value))}</td>"
            for value in row
        )
        body_rows += f"<tr>{row_cells}</tr>"

    st.markdown(
        f"""
        <div class="scroll-table-card">
            <div class="scroll-table-wrap" style="max-height: {int(height)}px;">
                <table class="scroll-table">
                    <thead>
                        <tr>{header_cells}</tr>
                    </thead>
                    <tbody>
                        {body_rows}
                    </tbody>
                </table>
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


def 위원회별_위원_등록(df):
    try:
        today = datetime.datetime.now().strftime("%Y-%m-%d")
        df['생년월일'] = pd.to_datetime(df['생년월일'], errors='coerce')
        df['위촉일자'] = pd.to_datetime(df['위촉일자'], errors='coerce')
        df['만료일자'] = pd.to_datetime(df['만료일자'], errors='coerce')
        df['임기'] = pd.to_numeric(df['임기'], errors='coerce')

        df['생년월일'] = df['생년월일'].apply(lambda x: None if pd.isna(x) else x.strftime('%Y-%m-%d'))
        df['위촉일자'] = df['위촉일자'].apply(lambda x: None if pd.isna(x) else x.strftime('%Y-%m-%d'))
        df['만료일자'] = df['만료일자'].apply(lambda x: None if pd.isna(x) else x.strftime('%Y-%m-%d'))
        df['갱신일'] = today

        insert_sql = """
        INSERT INTO commissioners (순서, 위원회명, 구분, 소속, 직위, 성명, 생년월일, 위촉일자, 만료일자, 임기, 성별, 갱신일)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(위원회명, 생년월일, 성명) DO UPDATE SET
            순서 = excluded.순서,
            구분 = excluded.구분,
            소속 = excluded.소속,
            직위 = excluded.직위,
            성명 = excluded.성명,
            위촉일자 = excluded.위촉일자,
            만료일자 = excluded.만료일자,
            임기 = excluded.임기,
            성별 = excluded.성별,
            갱신일 = excluded.갱신일
        """

        for _, row in df.iterrows():
            c.execute(insert_sql, (
                row['순서'], row['위원회명'], row['구분'], row['소속'], row['직위'], row['성명'], row['생년월일'],
                row['위촉일자'], row['만료일자'], row['임기'], row['성별'], row['갱신일']
            ))

        conn.commit()
        return True
    except Exception as e:
        st.error(f"데이터베이스 갱신 오류: {e}")
        return False

def 위원회_삭제(위원회명):
    try:
        실제_위원회명 = 위원회명.split("_")[-1].strip()
        delete_sql = "DELETE FROM committees WHERE 위원회명 = ?"
        c.execute(delete_sql, (실제_위원회명,))
        conn.commit()
        return c.rowcount
    except Exception as e:
        st.error(f"위원회 삭제 중 오류: {e}")
        return 0


def 테이블_불러오기(query):
    try:
        df = pd.read_sql(query, conn)
        return df
    except Exception as e:
        st.error(f"테이블 불러오기 중 오류: {e}")
        return pd.DataFrame()


def 회의운영현황_시트작성(writer, sheet_name, df_wide, years):
    df_wide = df_wide.copy().reindex(columns=회의운영현황_와이드컬럼(years))
    df_wide.to_excel(writer, index=False, sheet_name=sheet_name, startrow=3, header=False)

    workbook = writer.book
    worksheet = writer.sheets[sheet_name]

    section_format = workbook.add_format({
        "bold": True,
        "align": "center",
        "valign": "vcenter",
        "border": 1,
        "bg_color": "#D9D9D9",
    })
    year_format = workbook.add_format({
        "bold": True,
        "align": "center",
        "valign": "vcenter",
        "border": 1,
        "bg_color": "#EDEDED",
    })
    header_format = workbook.add_format({
        "bold": True,
        "align": "center",
        "valign": "vcenter",
        "border": 1,
        "bg_color": "#F7F7F7",
    })
    text_format = workbook.add_format({"border": 1, "valign": "vcenter"})
    center_format = workbook.add_format({"border": 1, "align": "center", "valign": "vcenter"})
    number_format = workbook.add_format({"border": 1, "align": "center", "valign": "vcenter", "num_format": "#,##0"})

    worksheet.set_row(0, 26)
    worksheet.set_row(1, 24)
    worksheet.set_row(2, 24)

    worksheet.merge_range(0, 0, 2, 0, "연번", section_format)
    worksheet.merge_range(0, 1, 2, 1, "소관부서", section_format)
    worksheet.merge_range(0, 2, 2, 2, "위원회명", section_format)

    meeting_start_col = 3
    meeting_end_col = meeting_start_col + len(years) * 3 - 1
    budget_start_col = meeting_end_col + 1
    budget_end_col = budget_start_col + len(years) - 1

    if meeting_start_col <= meeting_end_col:
        worksheet.merge_range(0, meeting_start_col, 0, meeting_end_col, "회의 개최 횟수", section_format)
    if budget_start_col <= budget_end_col:
        worksheet.merge_range(0, budget_start_col, 0, budget_end_col, "운영경비(단위:천원)", section_format)

    current_col = meeting_start_col
    for year in years:
        worksheet.merge_range(1, current_col, 1, current_col + 2, f"{year}년", year_format)
        worksheet.write(2, current_col, "합계", header_format)
        worksheet.write(2, current_col + 1, "서면", header_format)
        worksheet.write(2, current_col + 2, "대면", header_format)
        current_col += 3

    for year in years:
        worksheet.write(1, current_col, f"{year}년", year_format)
        worksheet.write(2, current_col, "운영경비", header_format)
        current_col += 1

    data_row_start = 3
    data_row_end = data_row_start + len(df_wide) - 1
    if data_row_end >= data_row_start:
        worksheet.set_column("A:A", 8)
        worksheet.set_column("B:B", 18)
        worksheet.set_column("C:C", 28)
        if meeting_start_col <= meeting_end_col:
            worksheet.set_column(meeting_start_col, meeting_end_col, 10)
        if budget_start_col <= budget_end_col:
            worksheet.set_column(budget_start_col, budget_end_col, 14)

        for row_offset, (_, row) in enumerate(df_wide.iterrows()):
            row_idx = data_row_start + row_offset
            worksheet.set_row(row_idx, 22)
            for col_idx, value in enumerate(row.tolist()):
                if col_idx in (1, 2):
                    worksheet.write_string(row_idx, col_idx, "" if pd.isna(value) else str(value), text_format)
                elif pd.isna(value) or str(value).strip() == "":
                    worksheet.write_blank(row_idx, col_idx, None, center_format if col_idx != 1 and col_idx != 2 else text_format)
                else:
                    numeric_value = pd.to_numeric(pd.Series([value]), errors="coerce").iloc[0]
                    if pd.notna(numeric_value):
                        cell_format = number_format if col_idx >= budget_start_col else center_format
                        worksheet.write_number(row_idx, col_idx, float(numeric_value), cell_format)
                    else:
                        worksheet.write_string(row_idx, col_idx, str(value), text_format)

    worksheet.freeze_panes(3, 3)
    worksheet.autofilter(2, 0, max(2, data_row_end), budget_end_col)


def 회의운영현황_양식_파일():
    output = BytesIO()
    years = 회의운영현황_기준연도()
    blank_wide_df = 회의운영현황_와이드데이터()

    example_row = {
        "연번": 1,
        "소관부서": "기획예산담당관",
        "위원회명": "예시위원회",
    }
    for year in years:
        example_row[f"{year}년_합계"] = 3 if year == years[-1] else pd.NA
        example_row[f"{year}년_서면"] = 1 if year == years[-1] else pd.NA
        example_row[f"{year}년_대면"] = 2 if year == years[-1] else pd.NA
        example_row[f"{year}년_운영경비(천원)"] = 1250 if year == years[-1] else pd.NA
    example_wide_df = pd.DataFrame([example_row], columns=회의운영현황_와이드컬럼(years))

    guide_df = pd.DataFrame(
        {
            "항목": [
                "입력 기준",
                "회의 개최 횟수",
                "운영경비",
                "합계 작성",
                "업로드 기준",
            ],
            "설명": [
                "한 행이 한 위원회입니다. 위원회명은 수정하지 말고 해당 연도 칸에 값만 입력해 주세요.",
                "각 연도별로 합계, 서면, 대면 순서로 작성합니다.",
                "운영경비는 천원 단위 숫자로 입력합니다. 예: 1250",
                "합계는 확인용입니다. 실제 반영은 서면/대면 값 기준으로 처리됩니다.",
                "입력양식 또는 현재 데이터 다운로드 파일을 수정한 뒤 그대로 업로드하면 됩니다.",
            ],
        }
    )

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        회의운영현황_시트작성(writer, "입력양식", blank_wide_df, years)
        회의운영현황_시트작성(writer, "작성예시", example_wide_df, years)
        guide_df.to_excel(writer, index=False, sheet_name="입력안내")

        workbook = writer.book
        guide_header_format = workbook.add_format({"bold": True, "bg_color": "#E2F3F0", "border": 1})
        guide_text_format = workbook.add_format({"border": 1, "text_wrap": True, "valign": "top"})
        guide_sheet = writer.sheets["입력안내"]
        guide_sheet.set_row(0, None, guide_header_format)
        guide_sheet.set_column("A:A", 18, guide_text_format)
        guide_sheet.set_column("B:B", 72, guide_text_format)
        guide_sheet.freeze_panes(1, 0)

    return output.getvalue()


def 회의운영현황_다운로드_파일():
    df_long = 회의운영현황_데이터()
    years = 회의운영현황_기준연도(df_long)
    df_wide = 회의운영현황_와이드데이터(df_long)

    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        회의운영현황_시트작성(writer, "입력양식", df_wide, years)

    return output.getvalue()


def generate_excel_file(committee_name, 기준일=None, 기준라벨=None):
    committee_name = str(committee_name).strip()
    if not committee_name:
        return BytesIO().getvalue()

    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_committees = fetch_data(
            """
            SELECT 순서, 위원회명, 부서명, 설치일자, 설치근거, `법령상 의무설치 여부`, 상설여부, 근거법령, 담당자, 연락처, 갱신일
            FROM committees
            WHERE 위원회명 = ?
            """,
            (committee_name,),
            ["순서", "위원회명", "부서명", "설치일자", "설치근거", "법령상 의무설치 여부", "상설여부", "근거법령", "담당자", "연락처", "갱신일"]
        )

        if df_committees.empty:
            df_committees = pd.DataFrame(columns=["순서", "위원회명", "부서명", "설치일자", "설치근거", "법령상 의무설치 여부", "상설여부", "근거법령", "담당자", "연락처", "갱신일"])

        df_committees.to_excel(writer, index=False, sheet_name="부서별_위원회_현황")

        worksheet1 = writer.sheets["부서별_위원회_현황"]
        worksheet1.set_column("A:A", 10)
        worksheet1.set_column("B:B", 50)
        worksheet1.set_column("C:C", 25)
        worksheet1.set_column("D:D", 12)
        worksheet1.set_column("E:H", 16)
        worksheet1.set_column("I:J", 15)
        worksheet1.set_column("K:K", 20)

        workbook = writer.book
        border_format = workbook.add_format({'border': 0, 'align': 'center'})
        for row_num in range(1, len(df_committees) + 2):
            worksheet1.set_row(row_num, None, border_format)

        df_comm_status = 기준일_위원회_위촉현황_데이터(위원회명=committee_name, 기준일=기준일)
        if not df_comm_status.empty and "부서명" in df_comm_status.columns:
            df_comm_status = df_comm_status.drop(columns=["부서명"])

        if df_comm_status.empty:
            df_comm_status = pd.DataFrame(columns=["순서", "구분", "소속", "직위", "성명",
                                                   "생년월일", "위촉일자", "만료일자", "임기", "성별"])
        else:
            df_comm_status = df_comm_status[["순서", "구분", "소속", "직위", "성명", "생년월일", "위촉일자", "만료일자", "임기", "성별"]]
        df_comm_status.insert(1, "위원회명", committee_name)
        df_comm_status.sort_values("순서", ascending=True, inplace=True)

        if 기준일 is not None or 기준라벨 is not None:
            info_df = pd.DataFrame([{
                "기준구분": 기준라벨 or "기준일 지정",
                "기준일": pd.Timestamp(기준일 if 기준일 is not None else today).strftime("%Y-%m-%d"),
                "산정방식": "위촉일자 <= 기준일 <= 만료일자(만료일자 공란 포함)"
            }])
            info_df.to_excel(writer, index=False, sheet_name="기준정보")
            writer.sheets["기준정보"].set_column("A:C", 28)

        df_comm_status.to_excel(writer, index=False, index_label="순서", sheet_name="위원회_명단")

        worksheet2 = writer.sheets["위원회_명단"]
        worksheet2.set_column("A:A", 10)
        worksheet2.set_column("B:B", 30)
        worksheet2.set_column("C:C", 10)
        worksheet2.set_column("D:E", 20)
        worksheet2.set_column("G:I", 15)

        for row_num in range(1, len(df_comm_status) + 2):
            worksheet2.set_row(row_num, None, border_format)

    return output.getvalue()


def generate_committee_snapshot_excel_file(기준일, 기준라벨, 부서명=None):
    committee_query = """
        SELECT d.순서, c.부서명, c.위원회명, c.설치일자, c.설치근거, c.`법령상 의무설치 여부`, c.상설여부, c.근거법령, c.담당자, c.연락처, c.갱신일
        FROM committees AS c
        LEFT JOIN departments AS d ON c.부서명 = d.부서명
        WHERE 1 = 1
    """
    committee_params = []
    if 부서명:
        committee_query += " AND c.부서명 = ?"
        committee_params.append(부서명)
    committee_query += " ORDER BY d.순서, c.위원회명"

    df_committees = fetch_data(
        committee_query,
        params=committee_params,
        columns=["순서", "부서명", "위원회명", "설치일자", "설치근거", "법령상 의무설치 여부", "상설여부", "근거법령", "담당자", "연락처", "갱신일"]
    )

    df_snapshot = 기준일_위원회_위촉현황_데이터(부서명=부서명, 기준일=기준일)
    if df_snapshot.empty:
        df_snapshot = pd.DataFrame(columns=[
            "순서", "부서명", "위원회명", "구분", "소속", "직위", "성명",
            "생년월일", "위촉일자", "만료일자", "임기", "성별"
        ])
    else:
        df_snapshot = df_snapshot.sort_values(["부서명", "위원회명", "순서", "성명"], na_position="last")

    summary_columns = ["부서명", "위원회명", "위원수", "당연직", "위촉직", "남성", "여성"]
    if df_snapshot.empty:
        df_summary = pd.DataFrame(columns=summary_columns)
    else:
        df_summary = (
            df_snapshot.groupby(["부서명", "위원회명"], dropna=False)
            .agg(
                위원수=("구분", "count"),
                당연직=("구분", lambda s: int(s.eq("당연직").sum())),
                위촉직=("구분", lambda s: int(s.eq("위촉직").sum())),
                남성=("성별", lambda s: int(s.eq("남").sum())),
                여성=("성별", lambda s: int(s.eq("여").sum())),
            )
            .reset_index()
            .sort_values(["부서명", "위원회명"], na_position="last")
        )

    info_df = pd.DataFrame([{
        "기준구분": 기준라벨,
        "기준일": pd.Timestamp(기준일).strftime("%Y-%m-%d"),
        "대상범위": 부서명 if 부서명 else "전체 위원회",
        "산정방식": "위촉일자 <= 기준일 <= 만료일자(만료일자 공란 포함)"
    }])

    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        info_df.to_excel(writer, index=False, sheet_name="기준정보")
        df_committees.to_excel(writer, index=False, sheet_name="부서별_위원회_현황")
        df_summary.to_excel(writer, index=False, sheet_name="위원회별_요약")
        df_snapshot.to_excel(writer, index=False, sheet_name="기준일_위촉현황")

        writer.sheets["기준정보"].set_column("A:D", 24)
        writer.sheets["부서별_위원회_현황"].set_column("A:K", 16)
        writer.sheets["위원회별_요약"].set_column("A:G", 18)
        writer.sheets["기준일_위촉현황"].set_column("A:L", 16)

    return output.getvalue()


def 위원회별_zip_파일생성():
    zip_buffer = BytesIO()
    has_files = False

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        df_committees = fetch_data(
            """
            SELECT c.위원회명, d.순서, d.부서명
            FROM committees AS c
            LEFT JOIN departments AS d ON c.부서명 = d.부서명
            ORDER BY d.순서
            """,
            columns=["위원회명", "순서", "부서명"]
        )

        df_committees = df_committees.dropna(subset=["위원회명"])

        for _, row in df_committees.iterrows():
            comm_name = str(row["위원회명"]).strip()
            if not comm_name:
                continue

            department_name = row["부서명"] or "미등록부서"
            sequence = row["순서"] or 0
            excel_data = generate_excel_file(comm_name)
            file_name = f"{sequence}.{department_name}_{comm_name}.xlsx"
            zip_file.writestr(file_name, excel_data)
            has_files = True

    if not has_files:
        return None

    zip_buffer.seek(0)
    return zip_buffer.getvalue()


# def 위원회별_zip_파일생성():
#     zip_buffer = BytesIO()
#     with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
#         df_committees = fetch_data(
#             """
#             SELECT c.위원회명, d.순서, d.부서명
#             FROM committees AS c
#             LEFT JOIN departments AS d ON c.부서명 = d.부서명
#             ORDER BY d.순서
#             """,
#             columns=["위원회명", "순서", "부서명"]
#         )
#         for _, row in df_committees.iterrows():
#             comm_name = row["위원회명"]
#             department_name = row["부서명"]
#             sequence = row["순서"]
#             excel_data = generate_excel_file(comm_name)
#             file_name = f"{sequence}.{department_name}_{comm_name}.xlsx"
#             zip_file.writestr(file_name, excel_data)

#     if not zip_file.namelist():
#         return None

#     zip_buffer.seek(0)
#     return zip_buffer.getvalue()


today = datetime.datetime.now().strftime("%Y-%m-%d")
스타일_적용()

join_query_1 = """
SELECT d.순서, d.부서명, c.위원회명, m.갱신일
FROM committees AS c
LEFT JOIN departments AS d ON c.부서명 = d.부서명
LEFT JOIN commissioners AS m ON c.위원회명 = m.위원회명
ORDER BY d.순서
"""

join_query_2 = """
SELECT m.순서, d.부서명, c.위원회명, m.구분, m.소속, m.직위, m.성명, m.생년월일, m.위촉일자, m.만료일자, m.임기, m.성별
FROM committees AS c
LEFT JOIN departments AS d ON c.부서명 = d.부서명
LEFT JOIN commissioners AS m ON c.위원회명 = m.위원회명
ORDER BY d.순서
"""

df_committees_status_1 = 테이블_불러오기(join_query_1)
df_committees_status_2 = 테이블_불러오기(join_query_2)
df_committee_base = 위원회_기본정보_데이터()
df_composition = 위원회_구성_집계()
df_department_summary = 부서_대시보드_요약()
df_gender = 위원회_성별_비율()
df_duplicate = 중복_위촉위원_데이터()
df_skewed = 성별_편중_위원회_데이터()
df_expiring = 임기만료_예정_데이터()
df_youth_ratio = 청년위원회_비율_데이터(df_committees_status_1, df_committees_status_2)
df_meeting_yearly = 회의운영현황_데이터()
meeting_template_file = 회의운영현황_양식_파일()
meeting_current_file = 회의운영현황_다운로드_파일()

valid_committees = df_committees_status_1.dropna(subset=['부서명'])
부서별_위원회_개수 = valid_committees.groupby('부서명')['위원회명'].nunique().to_dict()

department_order = (
    df_committees_status_1[['부서명', '순서']]
    .drop_duplicates()
    .dropna(subset=['부서명'])
    .sort_values('순서')
)

부서명_options = [
    f"{부서} 운영 위원회({부서별_위원회_개수.get(부서, 0)}개)"
    for 부서 in department_order['부서명']
]

총부서수 = int(department_order['부서명'].nunique()) if not department_order.empty else 0
총위원회수 = int(df_committee_base['위원회명'].dropna().nunique()) if not df_committee_base.empty else 0
총위원수, 위촉직수, 위촉직_여성수 = 위원_핵심지표_조회()
위촉직_여성비율 = round(위촉직_여성수 * 100 / 위촉직수, 1) if 위촉직수 else 0.0
최근갱신일 = pd.to_datetime(df_committee_base['갱신일'], errors='coerce').max() if not df_committee_base.empty else pd.NaT
최근갱신일_표시 = 최근갱신일.strftime("%Y-%m-%d") if pd.notna(최근갱신일) else "미등록"

hero_left, hero_right = st.columns([6.6, 1.8], vertical_alignment="center")
with hero_left:
    st.markdown(
        f"""
        <div class="hero-panel">
            <div class="hero-kicker">Committee Dashboard</div>
            <h1 class="hero-title">양산시 위원회 관리 시스템</h1>
            <p class="hero-copy">
                위원회 운영 현황, 위촉직 구성, 설치 정보와 만료 예정 현황을 첫 화면에서 바로 확인할 수 있도록 정리한 대시보드입니다.
                기존 조회와 관리자 기능은 그대로 유지하면서, 핵심 지표와 위험 신호를 더 빠르게 읽을 수 있게 구성했습니다.
            </p>
            <div class="hero-meta">
                <span class="hero-pill">최근 갱신일 {최근갱신일_표시}</span>
                <span class="hero-pill">운영 부서 {총부서수:,}개</span>
                <span class="hero-pill">등록 위원회 {총위원회수:,}개</span>
                <span class="hero-pill">등록 위원 {총위원수:,}명</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )
with hero_right:
    st.link_button(
        "양산시 위원회 조례 바로가기",
        "https://www.elis.go.kr/allalr/selectAlrBdtOne?alrNo=48330102202009&histNo=006&menuNm=main",
        use_container_width=True
    )
    info_panel(
        "홈 요약",
        [
            ("90일 내 만료", f"{len(df_expiring):,}명"),
            ("특정성별 초과", f"{len(df_skewed):,}개"),
            ("중복 위촉", f"{len(df_duplicate):,}명"),
            ("청년 비율 유의", f"{int((df_youth_ratio['상태'] == '유의').sum()) if not df_youth_ratio.empty else 0:,}개"),
        ]
    )

metric_col1, metric_col2, metric_col3, metric_col4 = st.columns(4)
with metric_col1:
    metric_card("운영 부서", f"{총부서수:,}", "위원회를 관리하는 부서 기준")
with metric_col2:
    metric_card("등록 위원회", f"{총위원회수:,}", "현재 데이터베이스에 등록된 위원회")
with metric_col3:
    metric_card("전체 위원", f"{총위원수:,}", "당연직과 위촉직을 포함한 전체 인원")
with metric_col4:
    metric_card("위촉직 여성 비율", f"{위촉직_여성비율:.1f}%", "위촉직 구성의 성별 균형 지표")

tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs(["홈 대시보드", "위원회 검색", "시의원 위촉현황", "위원회 설치정보", "연도별 회의·경비", "위원회 업무 매뉴얼", "관리자페이지"])
with tab1:
    st.markdown('<div class="section-title">부서별 운영 현황</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-copy">부서별 위원회 수와 선택한 위원회의 상세 구성을 함께 확인할 수 있습니다.</div>', unsafe_allow_html=True)

    overview_col, detail_col = st.columns([1.4, 1.04], gap="large", vertical_alignment="top")
    with overview_col:
        if not df_department_summary.empty:
            department_summary_view = (
                department_order[['순서', '부서명']]
                .merge(df_department_summary, on='부서명', how='right')
                .sort_values(['순서', '부서명'], na_position='last')
                .reset_index(drop=True)
            )
            chart_data = department_summary_view.sort_values(
                ['위원회수', '위원수', '부서명'],
                ascending=[True, True, True]
            )
            chart_height = max(360, len(chart_data) * 34)
            base = alt.Chart(chart_data).encode(
                x=alt.X("위원회수:Q", title="위원회 수"),
                y=alt.Y("부서명:N", sort=None, title=None),
                tooltip=[
                    alt.Tooltip("부서명:N", title="부서"),
                    alt.Tooltip("위원회수:Q", title="위원회 수"),
                    alt.Tooltip("위원수:Q", title="전체 위원"),
                    alt.Tooltip("위촉직수:Q", title="위촉직"),
                    alt.Tooltip("위촉직_여성비율:Q", title="위촉직 여성 비율"),
                ]
            )
            bars = base.mark_bar(color="#0f766e", cornerRadiusTopRight=10, cornerRadiusBottomRight=10)
            labels = base.mark_text(align="left", baseline="middle", dx=6, color="#0f172a").encode(text="위원회수:Q")
            with st.container(height=420):
                st.altair_chart((bars + labels).properties(height=chart_height), use_container_width=True)
            st.dataframe(
                department_summary_view[['순서', '부서명', '위원회수', '위원수', '위촉직수', '위촉직_여성비율']],
                hide_index=True,
                height=360,
                column_config={
                    '순서': {'width': 60},
                    '부서명': {'width': 160},
                    '위원회수': {'width': 90},
                    '위원수': {'width': 90},
                    '위촉직수': {'width': 90},
                    '위촉직_여성비율': st.column_config.NumberColumn("위촉직 여성 비율(%)", format="%.1f"),
                },
                use_container_width=True
            )
        else:
            st.info("표시할 부서 요약 데이터가 없습니다.")

    with detail_col:
        st.markdown('<div style="height: 0.1rem;"></div>', unsafe_allow_html=True)
        home_committee_options = df_committee_base['위원회명'].dropna().unique().tolist()
        if home_committee_options:
            st.markdown('<div class="panel-title">위원회 빠른 보기</div>', unsafe_allow_html=True)
            st.markdown('<div class="panel-copy">선택한 위원회의 기본 정보와 위원 구성을 한 번에 확인할 수 있습니다.</div>', unsafe_allow_html=True)
            selected_committee_home = st.selectbox(
                "위원회 빠른 보기",
                home_committee_options,
                key="home_committee_select",
                label_visibility="collapsed"
            )
            selected_meta = df_committee_base[df_committee_base['위원회명'] == selected_committee_home].iloc[0]
            selected_summary_df = df_composition[df_composition['위원회명'] == selected_committee_home]
            selected_summary = selected_summary_df.iloc[0] if not selected_summary_df.empty else None
            selected_gender = df_gender[df_gender['위원회명'] == selected_committee_home]
            설치일자 = pd.to_datetime(selected_meta['설치일자'], errors='coerce')
            설치일자_표시 = 설치일자.strftime("%Y-%m-%d") if pd.notna(설치일자) else "미등록"

            info_panel(
                selected_committee_home,
                [
                    ("소관 부서", selected_meta['부서명']),
                    ("설치일자", 설치일자_표시),
                    ("담당자", selected_meta['담당자']),
                    ("연락처", selected_meta['연락처']),
                    ("총 위원", f"{int(selected_summary['총계']):,}명" if selected_summary is not None else "-"),
                    ("위촉직", f"{int(selected_summary['위촉직']):,}명" if selected_summary is not None else "-"),
                ],
                extra_class="quickview-info-card"
            )

            if not selected_gender.empty:
                st.markdown('<div style="height: 0.85rem;"></div>', unsafe_allow_html=True)
                gender_chart_col, gender_info_col = st.columns([1.18, 0.82], gap="small")
                gender_summary_df = (
                    pd.DataFrame({"성별": ["남", "여"]})
                    .merge(selected_gender[["성별", "인원"]], on="성별", how="left")
                    .fillna({"인원": 0})
                )
                gender_summary_df["인원"] = gender_summary_df["인원"].astype(int)
                total_gender_count = int(gender_summary_df["인원"].sum())

                donut_chart = (
                    alt.Chart(gender_summary_df)
                    .mark_arc(innerRadius=56, outerRadius=86)
                    .encode(
                        theta=alt.Theta(field="인원", type="quantitative"),
                        color=alt.Color(
                            field="성별",
                            type="nominal",
                            scale=alt.Scale(domain=["남", "여"], range=["#0f766e", "#f59e0b"]),
                            legend=None
                        ),
                        tooltip=[
                            alt.Tooltip("성별:N", title="성별"),
                            alt.Tooltip("인원:Q", title="인원")
                        ]
                    )
                    .properties(
                        height=260,
                        title=f"{selected_committee_home} 위촉직 성별 구성",
                        padding={"left": 10, "right": 10, "top": 8, "bottom": 8}
                    )
                )

                with gender_chart_col:
                    st.altair_chart(donut_chart, use_container_width=True)

                with gender_info_col:
                    gender_summary_map = {
                        row["성별"]: f"{int(row['인원']):,}명 ({(int(row['인원']) / total_gender_count * 100 if total_gender_count else 0):.1f}%)"
                        for _, row in gender_summary_df.iterrows()
                    }
                    info_panel(
                        "성별 요약",
                        [
                            ("남", gender_summary_map.get("남", "0명 (0.0%)")),
                            ("여", gender_summary_map.get("여", "0명 (0.0%)")),
                        ]
                    )
            else:
                st.markdown('<div style="height: 0.85rem;"></div>', unsafe_allow_html=True)
                st.info("선택한 위원회의 위촉직 성별 정보가 없습니다.")
        else:
            st.info("표시할 위원회 데이터가 없습니다.")

    risk_col1, risk_col2, risk_col3, risk_col4 = st.columns(4)
    with risk_col1:
        st.markdown('<div class="section-title">중복 위촉 유의</div>', unsafe_allow_html=True)
        st.markdown('<div class="table-caption">4개 이상 위원회에 참여한 위촉직 위원입니다.</div>', unsafe_allow_html=True)
        if df_duplicate.empty:
            st.info("중복 위촉 유의 대상이 없습니다.")
        else:
            st.dataframe(
                df_duplicate[['성명', '참여위원회수', '소관부서']],
                hide_index=True,
                height=260,
                column_config={
                    '성명': {'width': 90},
                    '참여위원회수': {'width': 90},
                    '소관부서': {'width': 230},
                }
            )

    with risk_col2:
        st.markdown('<div class="section-title">특정 성별 초과</div>', unsafe_allow_html=True)
        st.markdown('<div class="table-caption">위촉직 특정 성별 비율이 60%를 초과한 위원회입니다.</div>', unsafe_allow_html=True)
        if df_skewed.empty:
            st.info("성별 편중 유의 대상이 없습니다.")
        else:
            st.dataframe(
                df_skewed[['부서명', '위원회명', '남성비율', '여성비율']],
                hide_index=True,
                height=260,
                column_config={
                    '부서명': {'width': 110},
                    '위원회명': {'width': 190},
                    '남성비율': st.column_config.NumberColumn('남성(%)', format="%.1f"),
                    '여성비율': st.column_config.NumberColumn('여성(%)', format="%.1f"),
                }
            )

    with risk_col3:
        st.markdown('<div class="section-title">3개월 내 만료</div>', unsafe_allow_html=True)
        st.markdown('<div class="table-caption">향후 90일 안에 임기가 끝나는 위촉직 위원입니다.</div>', unsafe_allow_html=True)
        if df_expiring.empty:
            st.info("90일 내 만료 예정 위원이 없습니다.")
        else:
            st.dataframe(
                df_expiring[['부서명', '위원회명', '성명', '만료일자']],
                hide_index=True,
                height=260,
                column_config={
                    '부서명': {'width': 110},
                    '위원회명': {'width': 160},
                    '성명': {'width': 90},
                    '만료일자': {'width': 100},
                }
            )

    with risk_col4:
        st.markdown('<div class="section-title">청년 비율 점검</div>', unsafe_allow_html=True)
        st.markdown('<div class="table-caption">청년 참여 비율 10% 기준을 함께 확인할 수 있습니다.</div>', unsafe_allow_html=True)
        if df_youth_ratio.empty:
            st.info("청년 비율 점검 대상 위원회가 없습니다.")
        else:
            youth_view = (
                df_youth_ratio.assign(정렬값=df_youth_ratio['상태'].map({'유의': 0, '충족': 1}).fillna(2))
                .sort_values(['정렬값', '비율(%)', '위원회명'])
                .drop(columns=['정렬값'])
            )
            st.dataframe(
                youth_view[['위원회명', '비율(%)', '상태']],
                hide_index=True,
                height=260,
                use_container_width=True,
                column_config={
                    '위원회명': st.column_config.TextColumn('위원회명', width='medium'),
                    '비율(%)': st.column_config.NumberColumn('비율(%)', format="%.2f"),
                    '상태': st.column_config.TextColumn('상태', width='small'),
                }
            )

    home_composition_excel = to_excel(df_composition, sheet_name="위원회_구성_전체보기")
    composition_header_col1, composition_header_col2 = st.columns([3.6, 1.2], vertical_alignment="bottom")
    with composition_header_col1:
        st.markdown('<div class="section-title">위원회 구성 전체 보기</div>', unsafe_allow_html=True)
        st.markdown('<div class="section-copy">부서 또는 위원회명으로 빠르게 좁혀서 전체 구성을 볼 수 있습니다.</div>', unsafe_allow_html=True)
    with composition_header_col2:
        st.download_button(
            label="전체 자료 엑셀 다운로드",
            data=home_composition_excel,
            file_name=f"홈대시보드_위원회구성전체보기_{today}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="home_composition_download"
        )

    filter_col1, filter_col2 = st.columns([1, 1.4])
    with filter_col1:
        home_department_filter = st.selectbox(
            "부서 필터",
            ["전체"] + department_order['부서명'].tolist(),
            key="home_department_filter"
        )
    with filter_col2:
        home_committee_keyword = st.text_input(
            "위원회명 검색",
            placeholder="예: 청년정책위원회",
            key="home_committee_keyword"
        )

    filtered_composition = df_composition.copy()
    if home_department_filter != "전체":
        filtered_composition = filtered_composition[filtered_composition['부서명'] == home_department_filter]
    if home_committee_keyword:
        filtered_composition = filtered_composition[
            filtered_composition['위원회명'].str.contains(home_committee_keyword, case=False, na=False)
        ]

    if filtered_composition.empty:
        st.info("조건에 맞는 위원회 구성 정보가 없습니다.")
    else:
        st.dataframe(
            filtered_composition,
            hide_index=True,
            height=420,
            column_config={
                "부서명": {"width": 130},
                "위원회명": {"width": 230},
                "총계": {"width": 70},
                "당연직": {"width": 70},
                "위촉직": {"width": 70},
                "위촉직_남성": {"width": 90, "label": "위촉직(남)"},
                "위촉직_여성": {"width": 90, "label": "위촉직(여)"},
                "위촉직_남성_비율": st.column_config.NumberColumn("위촉직 남성(%)", format="%.1f"),
                "위촉직_여성_비율": st.column_config.NumberColumn("위촉직 여성(%)", format="%.1f"),
            },
            use_container_width=True
        )

with tab2:
    st.subheader("가. 부서별 위원회 현황")
    선택된_부서명 = None
    col1, col2 = st.columns([2.5, 7.5])

    with col1:
        if 부서명_options:
            선택된_부서명_with_count = st.selectbox('부서를 선택하세요', 부서명_options, label_visibility="collapsed", key='unique_key_for_this_selectbox')
            선택된_부서명 = 선택된_부서명_with_count.split(" ")[0] if 선택된_부서명_with_count else None

            if 선택된_부서명:
                선택된_부서명_df = df_committees_status_1[df_committees_status_1['부서명'] == 선택된_부서명].drop_duplicates(subset='위원회명')
                if not 선택된_부서명_df.empty:
                    st.dataframe(선택된_부서명_df, hide_index=True, column_config={
                        '순서': {'width': 50}, '부서명': {'width': 120}, '위원회명': {'width': 180}, '갱신일': {'width': 100}},height=300)
                else:
                    st.info("등록된 위원회 및 위원 정보가 없습니다.")
            else:
                st.error("부서를 선택해 주세요.")
        else:   
            st.info("부서 정보를 등록하세요.")

    with col2:
        선택된_부서명_df = df_committees_status_2[df_committees_status_2['부서명'] == 선택된_부서명]

        if not 선택된_부서명_df.empty:
            위원회명_options = 선택된_부서명_df['위원회명'].unique()
            선택된_위원회명 = st.selectbox('위원회를 선택하세요', 위원회명_options, label_visibility="collapsed", key='unique_key_for_this_selectbox2')
            선택된_위원회명_df = 선택된_부서명_df[선택된_부서명_df['위원회명'] == 선택된_위원회명].sort_values(by='순서')
            st.dataframe(선택된_위원회명_df, hide_index=True, height=300, column_config={
                '순서': {'width': 50},
                '부서명': {'width': 120},
                '위원회명': {'width': 200},
                '구분': {'width': 80},
                '소속': {'width': 150},
                '직위': {'width': 150},
                '성명': {'width': 100},
                '생년월일': {'width': 110},
                '위촉일자': {'width': 110},
                '만료일자': {'width': 110},
                '임기': {'width': 80},
                '성별': {'width': 80}
            })
        else:
            st.info("선택한 부서에 등록된 위원회 및 위원 정보가 없습니다.")


    st.subheader("나. 양산시 위원 조회")

    col1, col2 = st.columns([2.5, 7.5])    

    with col1:
        위원 = st.text_input("위원검색",placeholder="양산시 위원을 이름으로 검색하세요",label_visibility="collapsed")

    with col2:
        if 위원:

            c.execute("""
            SELECT 위원회명, 구분, 소속, 직위, 성명, 생년월일, 위촉일자, 만료일자, 임기, 성별
            FROM commissioners
            WHERE 성명 LIKE ?
            """, ('%' + 위원 + '%',))


            # 검색 결과를 데이터프레임으로 변환
            결과 = c.fetchall()
            
            열_이름 = ['위원회명','구분','소속','직위','성명', '생년월일', '위촉일자', '만료일자', '임기','성별']
            결과_df = pd.DataFrame(결과, columns=열_이름)

            # 검색 결과가 존재하는 경우
            if not 결과_df.empty:
                st.dataframe(결과_df,hide_index=True, height=300, column_config={
                '위원회명': {'width': 370},
                '구분': {'width': 80},
                '소속': {'width': 150},
                '직위': {'width': 150},
                '성명': {'width': 100},
                '생년월일': {'width': 110},
                '위촉일자': {'width': 110},
                '만료일자': {'width': 110},
                '임기': {'width': 80},
                '성별': {'width': 80}
            })
            else:
                #st.write("검색결과가 없습니다.")
                st.success("검색결과가 없습니다.", icon="✅")
            
            # 데이터베이스 변경 사항 커밋
            conn.commit()


with tab3:
    st.subheader("시의원 위원 위촉 현황")
    st.warning('시의원은 중복 위촉(3회 초과) 예외', icon="⚠️")
    st.caption(f"기준일별 현황은 현재 DB에 저장된 위촉일자와 만료일자를 기준으로 계산하며, 기준일은 {기준일_고정값.strftime('%Y-%m-%d')}입니다.")

    기준일_목록 = 기준일_옵션_목록()
    council_tabs = st.tabs([label for label, _, _ in 기준일_목록])

    for council_tab, (label, 기준일, key_suffix) in zip(council_tabs, 기준일_목록):
        with council_tab:
            기준일_표시 = 기준일.strftime("%Y-%m-%d")
            df_council = 시의원_위촉현황_데이터(기준일)
            summary = 시의원_위촉현황_요약(df_council)
            council_excel = 시의원_위촉현황_다운로드_파일(df_council, summary, 기준일, label)

            council_header_col1, council_header_col2 = st.columns([2.2, 1])
            with council_header_col1:
                st.caption(f"기준일: {기준일_표시}")
            with council_header_col2:
                st.download_button(
                    label=f"{label} 엑셀 다운로드",
                    data=council_excel,
                    file_name=f"시의원_위촉현황_{key_suffix}_{기준일_표시}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key=f"council_download_{key_suffix}"
                )

            if df_council.empty:
                st.info(f"{기준일_표시} 기준으로 유효한 시의원 위촉위원 정보가 없습니다.")
                continue

            st.markdown("### 1) 시의원 위원 개별 명단")
            st.dataframe(
                df_council,
                hide_index=True,
                column_config={
                    "부서명": {"width": 120},
                    "위원회명": {"width": 200},
                    "성명": {"width": 90},
                    "구분": {"width": 80},
                    "직위": {"width": 80},
                    "생년월일": {"width": 110},
                    "위촉일자": {"width": 110},
                    "만료일자": {"width": 110},
                    "임기": {"width": 80},
                    "성별": {"width": 60},
                    "소속": {"width": 150},
                },
                height=350
            )

            st.markdown("### 2) 위원회별 시의원 위촉 현황")
            st.dataframe(
                summary,
                hide_index=True,
                column_config={
                    "부서명": {"width": 120},
                    "위원회명": {"width": 220},
                    "시의원수": {"width": 90},
                    "명단": {"width": 250},
                },
                height=250
            )

with tab4:
    install_query = """
    SELECT d.순서, c.부서명, c.위원회명, c.설치일자, c.설치근거, c.`법령상 의무설치 여부`, c.상설여부, c.근거법령, c.갱신일
    FROM committees AS c
    LEFT JOIN departments AS d ON c.부서명 = d.부서명
    ORDER BY d.순서, c.위원회명
    """
    columns = ["순서", "부서명", "위원회명", "설치일자", "설치근거", "법령상 의무설치 여부", "상설여부", "근거법령", "갱신일"]
    df_install_info = fetch_data(install_query, columns=columns)

    if not df_install_info.empty:
        df_install_info["설치일자"] = pd.to_datetime(df_install_info["설치일자"], errors="coerce").dt.date.astype("string")
        df_install_info["갱신일"] = pd.to_datetime(df_install_info["갱신일"], errors="coerce").dt.date.astype("string")

    install_subtab1, install_subtab2 = st.tabs(["가. 위원회 설치정보", "나. 한눈에 보기"])

    with install_subtab1:
        st.caption("설치일자, 설치근거, 의무설치 여부, 상설여부, 조례 정보를 포함한 전체 설치정보입니다.")
        if df_install_info.empty:
            st.info("등록된 설치정보가 없습니다.")
        else:
            install_full_height = min(640, max(320, 35 * (len(df_install_info) + 1)))
            st.dataframe(
                df_install_info,
                hide_index=True,
                column_config={
                    "순서": st.column_config.NumberColumn("순서", width="small"),
                    "부서명": st.column_config.TextColumn("부서명", width="small"),
                    "위원회명": st.column_config.TextColumn("위원회명", width="medium"),
                    "설치일자": st.column_config.TextColumn("설치일자", width="small"),
                    "설치근거": st.column_config.TextColumn("설치근거", width="small"),
                    "법령상 의무설치 여부": st.column_config.TextColumn("법령상 의무설치 여부", width="small"),
                    "상설여부": st.column_config.TextColumn("상설여부", width="small"),
                    "근거법령": st.column_config.TextColumn("근거법령", width="medium"),
                    "갱신일": st.column_config.TextColumn("갱신일", width="small"),
                },
                height=install_full_height,
                use_container_width=True,
            )

        install_excel = to_excel(df_install_info, sheet_name="위원회_설치정보")
        st.download_button(
            label="설치정보 엑셀 다운로드",
            data=install_excel,
            file_name=f"위원회_설치정보_{today}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

        st.caption("설치정보 엑셀 업로드시 부서명, 위원회명, 설치일자, 설치근거, 근거법령 컬럼이 필요하며, 법령상 의무설치 여부와 상설여부 컬럼은 추가 입력할 수 있습니다.")
        install_upload = st.file_uploader("설치정보 엑셀을 업로드하세요", type=["xlsx"], key="install_info_upload", label_visibility="collapsed")

        if install_upload:
            df_install_upload = pd.read_excel(install_upload)
            if "강행/임의" in df_install_upload.columns and "법령상 의무설치 여부" not in df_install_upload.columns:
                df_install_upload = df_install_upload.rename(columns={"강행/임의": "법령상 의무설치 여부"})
            st.write("업로드된 설치정보")
            st.dataframe(df_install_upload, hide_index=True)

            required_columns = {"부서명", "위원회명", "설치일자", "설치근거", "근거법령"}

            if required_columns.issubset(df_install_upload.columns):
                if st.button("설치정보 반영하기", use_container_width=True, key="install_update"):
                    updated = 부서별_위원회_갱신(df_install_upload)
                    refresh_db_connection()
                    st.success(f"설치정보가 {updated}건 반영되었습니다.")
                    time.sleep(1)
                    st.rerun()
            else:
                st.error("엑셀 파일에는 [부서명, 위원회명, 설치일자, 설치근거, 근거법령] 컬럼이 모두 포함되어야 하며, [법령상 의무설치 여부], [상설여부] 컬럼은 선택 입력할 수 있습니다.")

    with install_subtab2:
        st.caption("설치근거와 상설여부를 위원회 수와 비율로 한눈에 볼 수 있도록 정리했습니다.")
        if df_install_info.empty:
            st.info("등록된 설치정보가 없습니다.")
        else:
            install_summary_source = df_install_info.copy().fillna("")
            install_basis = install_summary_source["설치근거"].astype(str).str.strip()
            install_mandatory = install_summary_source["법령상 의무설치 여부"].astype(str).str.strip()
            install_standing = install_summary_source["상설여부"].astype(str).str.strip()
            total_count = len(install_summary_source)

            def percent_value(count):
                return int(round((count / total_count) * 100)) if total_count else 0

            law_mandatory_count = int(((install_basis == "법령") & (install_mandatory == "강행")).sum())
            law_optional_count = int(((install_basis == "법령") & (install_mandatory == "임의")).sum())
            ordinance_count = int((install_basis == "조례").sum())
            other_basis_count = int(total_count - law_mandatory_count - law_optional_count - ordinance_count)

            standing_count = int((install_standing == "상설").sum())
            nonstanding_count = int((install_standing == "비상설").sum())
            other_standing_count = int(total_count - standing_count - nonstanding_count)

            df_install_basis_summary = pd.DataFrame(
                [
                    ["위원회 수(개)", total_count, law_mandatory_count, law_optional_count, ordinance_count, other_basis_count],
                    ["비율(%)", 100 if total_count else 0, percent_value(law_mandatory_count), percent_value(law_optional_count), percent_value(ordinance_count), percent_value(other_basis_count)],
                ],
                columns=["구분", "계", "법령상 강행", "법령상 임의", "조례상", "기타"]
            )

            df_standing_summary = pd.DataFrame(
                [
                    ["위원회 수(개)", total_count, standing_count, nonstanding_count, other_standing_count],
                    ["비율(%)", 100 if total_count else 0, percent_value(standing_count), percent_value(nonstanding_count), percent_value(other_standing_count)],
                ],
                columns=["구분", "계", "상설", "비상설", "기타"]
            )

            summary_col1, summary_col2 = st.columns([1.45, 1.0], gap="medium")

            with summary_col1:
                summary_table_panel(
                    title="설치 근거 및 비율(%)",
                    columns=["구분", "계", "법령상 강행", "법령상 임의", "조례상", "기타"],
                    rows=df_install_basis_summary.values.tolist(),
                    note="기타는 규칙, 지침 등과 미입력 항목을 포함합니다.",
                )

            with summary_col2:
                summary_table_panel(
                    title="상설 여부 및 비율(%)",
                    columns=["구분", "계", "상설", "비상설", "기타"],
                    rows=df_standing_summary.values.tolist(),
                    note="기타는 상설 여부가 입력되지 않은 항목입니다.",
                )


with tab5:
    st.markdown('<div class="section-title">연도별 회의 개최 및 운영경비</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="section-copy">양식은 회의 개최 횟수(연도별 합계·서면·대면), 운영경비 형식으로 구성되어 있으며, 위원회별 상세 내역까지 함께 확인할 수 있습니다.</div>',
        unsafe_allow_html=True
    )

    meeting_download_col1, meeting_download_col2 = st.columns(2)
    with meeting_download_col1:
        st.download_button(
            label="입력양식 다운로드",
            data=meeting_template_file,
            file_name=f"연도별_회의운영경비_입력양식_{today}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="meeting_template_download"
        )
    with meeting_download_col2:
        st.download_button(
            label="현재 데이터 다운로드",
            data=meeting_current_file,
            file_name=f"연도별_회의운영경비_{today}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="meeting_current_download"
        )

    if df_meeting_yearly.empty:
        st.info("등록된 연도별 회의·운영경비 데이터가 없습니다. 입력양식을 내려받아 관리자페이지에서 업로드해 주세요.")
    else:
        meeting_filter_col1, meeting_filter_col2, meeting_filter_col3 = st.columns([1, 1, 1.2])
        meeting_year_options = ["전체"] + [str(year) for year in sorted(df_meeting_yearly["연도"].unique(), reverse=True)]
        meeting_department_options = ["전체"] + sorted(df_meeting_yearly["부서명"].fillna("미등록").unique().tolist())

        with meeting_filter_col1:
            selected_meeting_year = st.selectbox("연도", meeting_year_options, key="meeting_year_filter")
        with meeting_filter_col2:
            selected_meeting_department = st.selectbox("부서", meeting_department_options, key="meeting_department_filter")
        with meeting_filter_col3:
            meeting_keyword = st.text_input(
                "위원회명 검색",
                placeholder="예: 청년정책위원회",
                key="meeting_keyword_filter"
            )

        filtered_meeting_df = df_meeting_yearly.copy()
        if selected_meeting_year != "전체":
            filtered_meeting_df = filtered_meeting_df[filtered_meeting_df["연도"] == int(selected_meeting_year)]
        if selected_meeting_department != "전체":
            filtered_meeting_df = filtered_meeting_df[filtered_meeting_df["부서명"] == selected_meeting_department]
        if meeting_keyword:
            filtered_meeting_df = filtered_meeting_df[
                filtered_meeting_df["위원회명"].str.contains(meeting_keyword, case=False, na=False)
            ]

        if filtered_meeting_df.empty:
            st.info("선택한 조건에 해당하는 연도별 회의·운영경비 데이터가 없습니다.")
        else:
            meeting_summary = 연도별_회의운영_요약(filtered_meeting_df)
            summary_face_to_face = int(filtered_meeting_df["대면회의"].sum())
            summary_document = int(filtered_meeting_df["서면회의"].sum())
            summary_budget = int(filtered_meeting_df["운영경비(천원)"].sum())
            summary_committees = int(filtered_meeting_df["위원회명"].nunique())

            metric_col1, metric_col2, metric_col3, metric_col4 = st.columns(4)
            with metric_col1:
                metric_card("대상 위원회", f"{summary_committees:,}개", "현재 필터 기준 위원회 수")
            with metric_col2:
                metric_card("대면 회의", f"{summary_face_to_face:,}회", "필터 조건의 연도별 합계")
            with metric_col3:
                metric_card("서면 회의", f"{summary_document:,}회", "필터 조건의 연도별 합계")
            with metric_col4:
                metric_card("운영경비", f"{summary_budget:,}천원", "필터 조건의 연도별 합계")

            st.markdown('<div class="section-title">연도별 요약</div>', unsafe_allow_html=True)
            st.dataframe(
                meeting_summary,
                hide_index=True,
                column_config={
                    "연도": {"width": 80},
                    "위원회수": {"width": 90},
                    "대면회의": {"width": 90},
                    "서면회의": {"width": 90},
                    "총회의": {"width": 90},
                    "운영경비(천원)": {"width": 120},
                },
                use_container_width=True
            )

            st.markdown('<div class="section-title">위원회별 상세 내역</div>', unsafe_allow_html=True)
            detailed_meeting_df = filtered_meeting_df.sort_values(
                ["연도", "순서", "부서명", "위원회명"],
                ascending=[False, True, True, True],
                na_position="last"
            )
            st.dataframe(
                detailed_meeting_df[["연도", "부서명", "위원회명", "대면회의", "서면회의", "총회의", "운영경비(천원)", "갱신일"]],
                hide_index=True,
                column_config={
                    "연도": {"width": 80},
                    "부서명": {"width": 130},
                    "위원회명": {"width": 230},
                    "대면회의": {"width": 90},
                    "서면회의": {"width": 90},
                    "총회의": {"width": 90},
                    "운영경비(천원)": {"width": 120},
                    "갱신일": {"width": 100},
                },
                height=420,
                use_container_width=True
            )


# 위원회 업무 매뉴얼
with tab6:
    st.subheader("위원회 업무 매뉴얼")
    st.caption("위원회별 매뉴얼이 아니라, 전반적인 절차와 흐름을 위한 공통 한글(.hwp, .hwpx) 파일 1건을 등록하고 다운로드할 수 있습니다.")

    common_manual = 공통_업무매뉴얼_파일()
    manual_left, manual_right = st.columns([1.05, 1], vertical_alignment="top")

    with manual_left:
        if common_manual:
            uploaded_at = pd.to_datetime(common_manual["업로드일"], errors="coerce")
            uploaded_at_text = uploaded_at.strftime("%Y-%m-%d %H:%M") if pd.notna(uploaded_at) else "-"
            info_panel(
                "현재 등록 파일",
                [
                    ("구분", "공통 업무 절차 매뉴얼"),
                    ("파일명", common_manual["파일명"]),
                    ("파일크기", 파일크기_표시(common_manual["파일크기"])),
                    ("업로드일", uploaded_at_text),
                ]
            )
            st.download_button(
                label="공통 업무 매뉴얼 다운로드",
                data=common_manual["파일데이터"],
                file_name=common_manual["파일명"],
                mime=common_manual["파일형식"],
                use_container_width=True,
                key="manual_download_button"
            )
        else:
            st.info("등록된 공통 업무 매뉴얼이 없습니다.")

    with manual_right:
        st.markdown("#### 공통 업무 매뉴얼 등록")
        st.caption("위원회 공통으로 참고하는 절차, 진행 순서, 업무 흐름 등을 정리한 한글 파일을 업로드하세요.")

        manual_upload = st.file_uploader(
            "공통 업무 매뉴얼 한글 파일을 업로드하세요",
            type=["hwp", "hwpx"],
            key="committee_manual_upload"
        )

        if manual_upload is not None and st.button("업무 매뉴얼 저장하기", use_container_width=True, key="manual_upload_button"):
            if 공통_업무매뉴얼_저장(manual_upload):
                st.success("공통 업무 매뉴얼이 저장되었습니다.")
                time.sleep(1)
                st.rerun()


# 관리자 페이지
with tab7:
    password = st.text_input("관리자 암호 입력", type="password")
    if password == "1234":
        st.title("관리자 기능")

        col1, col2, col3 = st.columns([2.5,3,5])

        with col1:
            st.subheader("1. 부서현황 관리")

            st.expander("1. 위원회 부서현황을 현행화 할 수 있습니다.")
            st.caption("1. 등록된 부서를 확인할 수 있습니다.")
            df_departments = 테이블_불러오기("SELECT * FROM departments")

            if not df_departments.empty:
                st.dataframe(df_departments, height=350, hide_index=True, column_config={
                    '순서': {'width': 100},
                    '부서명': {'width': 150},
                    '갱신일': {'width': 150}
                })
            else:
                st.info("등록된 부서 정보가 없습니다.") 

            df_dept_db = fetch_data("SELECT 순서, 부서명 FROM departments", columns=["순서","부서명"])
            dept_excel = to_excel(df_dept_db, sheet_name="부서현황")

            st.caption("2. 부서정보를 엑셀 파일으로 다운로드할 수 있습니다.")
            st.download_button(
                label="부서정보 다운로드",
                data=dept_excel,
                file_name=f"부서현황_{today}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

            st.caption("3. 부서정보를 일괄 변경할 수 있습니다.")
            dept_upload = st.file_uploader("부서정보를 등록 및 갱신하세요.", type=["xlsx"], key="dept", label_visibility="collapsed")

            if dept_upload:
                df_dept = pd.read_excel(dept_upload)
                if "부서명" in df_dept.columns:
                    st.write("업로드된 부서 정보")
                    st.dataframe(df_dept, hide_index=True, height=250, column_config={
                        '순번': {'width': 100},
                        '부서명': {'width': 150}
                    })  # 데이터프레임을 화면에 출력

                    if st.button("부서정보 반영하기", use_container_width=True):
                        success = 부서현황_갱신("departments", df_dept, "부서명")
                        if success:
                            st.success("부서 정보가 갱신되었습니다!")
                            time.sleep(1)
                            st.rerun()
                        else:
                            st.error("부서 정보 갱신에 실패했습니다.")
                else:
                    st.error("엑셀 파일에 [부서명] 컬럼이 포함되어 있어야 합니다.")


            st.caption("4. 등록된 개별 부서정보를 삭제할 수 있습니다.")
            df_departments = 테이블_불러오기("SELECT * FROM departments")

            # 부서명을 선택할 수 있는 셀렉트 박스를 생성합니다.
            부서명_리스트 = [f"{row['순서']}. {row['부서명']}" for index, row in df_departments.iterrows()]
            선택된_부서명 = st.selectbox("삭제할 부서를 선택하세요:", 부서명_리스트, label_visibility="collapsed")

            # 삭제 버튼 동작 구현
            if st.button("선택 부서 삭제하기", use_container_width=True):
                if 선택된_부서명:
                    삭제된_행_수 = 부서_삭제(선택된_부서명)
                    if 삭제된_행_수:
                        st.success(f"{선택된_부서명} 삭제 완료.")
                        time.sleep(1)
                        st.rerun()
                    else:
                        st.error("부서 삭제에 실패했습니다.")
                else:
                    st.warning("부서명을 선택하세요.")

        with col2:
            st.subheader("2. 부서별 위원회 현황 관리")
            st.expander("부서별 위원회를 등록하거나 수정할 수 있습니다.")
            st.caption("1. 등록된 위원회 현황을 확인할 수 있습니다.")
            join_query = """
            SELECT c.순서, d.부서명, c.위원회명, c.담당자, c.연락처, c.갱신일
            FROM departments AS d
            LEFT JOIN committees AS c ON d.부서명 = c.부서명
            ORDER BY d.순서
            """
            df_join = 테이블_불러오기(join_query)
            if not df_join.empty:
                st.dataframe(df_join, height=350, hide_index=True, column_config={
                    '순서': {'width': 30},
                    '부서명': {'width': 70},
                    '위원회명': {'width': 80},
                    '설치일자': {'width': 80},
                    '설치근거': {'width': 50},
                    '근거법령(조례)': {'width': 80},
                    '담당자': {'width': 50},
                    '연락처': {'width': 50},
                    '갱신일': {'width': 80}
                })
            else:
                st.info("등록된 위원회 정보가 없습니다.")

            # 모든 컬럼을 가져와 정확히 사용합니다.
            df_comm_db = fetch_data("SELECT 순서, 부서명, 위원회명, 담당자, 연락처 FROM committees", columns=["순서", "부서명", "위원회명", "담당자", "연락처"])

            # 엑셀Writer 객체를 만들고, 가져온 데이터프레임을 특정 시트에 저장
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_comm_db.to_excel(writer, sheet_name="부서별_위원회_현황", index=False)

                # 워크시트와 형식 설정
                workbook = writer.book
                worksheet = writer.sheets["부서별_위원회_현황"]
                worksheet.set_column("A:A", 10)
                worksheet.set_column("B:B", 20)
                worksheet.set_column("C:C", 50)
                worksheet.set_column("D:E", 10)

                # 셀 테두리 형식 정의
                border_format = workbook.add_format({'border': 1, 'align': 'center'})

                # 데이터를 담고 있는 셀에 테두리 적용
                for row_num in range(1, len(df_comm_db) + 2):
                    worksheet.set_row(row_num, None, border_format)

            # 데이터를 바이너리로 변환
            output.seek(0)

            st.caption("2. 관리부서 및 담당자 현황을 다운로드할 수 있습니다.")
            st.download_button(
                label="양산시 위원회현황 다운로드",
                data=output,
                file_name=f"양산시_위원회현황_{today}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

            st.caption("3. 위원회를 기준으로 관리부서 및 담당자를 갱신할 수 있습니다.")
            committee_upload = st.file_uploader("위원회 엑셀 파일을 선택하세요", type=["xlsx"], key="committee", label_visibility="collapsed")

            if committee_upload:
                df_comm = pd.read_excel(committee_upload, sheet_name="부서별_위원회_현황")
                st.write("업로드된 데이터")
                st.dataframe(df_comm, hide_index=True, column_config={'순서': {'width': 100}, '부서명': {'width': 150}, '위원회명': {'width': 200}, '담당자': {'width': 150}, '연락처': {'width': 100}})  # 데이터프레임을 화면에 출력

                required_columns = {"부서명", "위원회명"}

                if required_columns.issubset(df_comm.columns):
                    # 등록된 부서명을 데이터베이스에서 가져옴
                    dept_names = pd.DataFrame(fetch_data("SELECT 부서명 FROM departments", columns=["부서명"]))

                    # 엑셀에 있는 부서 중 등록되지 않은 부서명 확인
                    invalid_entries = df_comm[~df_comm["부서명"].isin(dept_names["부서명"])]
                    if not invalid_entries.empty:
                        st.sidebar.warning(
                            "다음 부서명은 등록되어 있지 않습니다:\n" + "\n".join(invalid_entries["부서명"].unique())
                        )

                    # 갱신 버튼 추가
                    if st.button("위원회자료 반영하기", use_container_width=True):
                        # 올바른 부서명이 있는 행들만 필터링
                        valid_entries = df_comm[df_comm["부서명"].isin(dept_names["부서명"])]
                        if not valid_entries.empty:
                            updated = 부서별_위원회_갱신(valid_entries)
                            st.success(f"위원회 데이터가 {updated}건 갱신되었습니다!")
                            time.sleep(1)
                            st.rerun()
                        else:
                            st.sidebar.error("유효한 부서명이 없습니다. 부서명을 확인해 주세요.")
                else:
                    st.sidebar.error("엑셀 파일에는 [부서명, 위원회명] 컬럼이 모두 포함되어 있어야 합니다.")

            st.caption("4. 등록된 개별 위원회정보를 삭제할 수 있습니다.")
            join_query = """
                        SELECT d.순서, d.부서명, c.위원회명, c.담당자, c.연락처
                        FROM departments AS d
                        LEFT JOIN committees AS c ON d.부서명 = c.부서명
                        ORDER BY d.순서
                        """
            df_committees = 테이블_불러오기(join_query)

            # 순서, 부서명, 위원회명을 포함하여 리스트 구성
            위원회명_리스트 = [f"{row['순서']}. {row['부서명']}_{row['위원회명']}" for _, row in df_committees.iterrows()]

            # 위원회명을 선택할 수 있는 셀렉트 박스를 생성합니다.
            선택된_위원회명 = st.selectbox("삭제할 위원회를 선택하세요:", 위원회명_리스트, label_visibility="collapsed")

            if st.button("선택한 위원회 삭제하기", use_container_width=True):
                if not 선택된_위원회명:
                    st.warning("위원회명을 선택하세요.")
                else:
                    삭제된_행_수 = 위원회_삭제(선택된_위원회명)
                    if 삭제된_행_수:
                        st.success(f"{선택된_위원회명} 삭제 완료.")
                        time.sleep(1)
                        st.rerun()
                    else:
                        st.error("위원회 삭제에 실패했습니다.")
            
            # if st.button("선택한 위원회 삭제하기", use_container_width=True):
            #     if 선택된_위원회명:
            #         삭제된_행_수 = 위원회_삭제(선택된_위원회명)
            #         if 삭제된_행_수:
            #             st.success(f"{선택된_위원회명} 삭제 완료.")
            #             time.sleep(1)
            #             st.rerun()
            #         else:
            #             st.error("위원회 삭제에 실패했습니다.")
            #     else:
            #         st.warning("위원회명을 선택하세요.")

            st.markdown("---")
            st.subheader("2-1. 연도별 회의·운영경비 관리")
            st.caption("입력양식은 6번 회의 개최 횟수(연도별 합계·서면·대면), 7번 운영경비 형식입니다.")
            st.caption("같은 연도와 같은 위원회명으로 업로드하면 기존 데이터가 덮어써지며, 합계는 확인용이고 실제 반영은 서면/대면 값 기준입니다.")

            admin_meeting_download_col1, admin_meeting_download_col2 = st.columns(2)
            with admin_meeting_download_col1:
                st.download_button(
                    label="입력양식 다운로드",
                    data=meeting_template_file,
                    file_name=f"연도별_회의운영경비_입력양식_{today}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="admin_meeting_template_download"
                )
            with admin_meeting_download_col2:
                st.download_button(
                    label="현재 데이터 다운로드",
                    data=meeting_current_file,
                    file_name=f"연도별_회의운영경비_{today}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="admin_meeting_current_download"
                )

            meeting_upload = st.file_uploader(
                "연도별 회의·운영경비 엑셀 파일을 업로드하세요",
                type=["xlsx"],
                key="meeting_yearly_upload",
                label_visibility="collapsed"
            )

            if meeting_upload:
                df_meeting_upload = 회의운영현황_엑셀읽기(meeting_upload)
                st.write("업로드된 회의·운영경비 데이터")
                st.dataframe(df_meeting_upload, hide_index=True, use_container_width=True)

                required_columns = {"연도", "위원회명", "대면회의", "서면회의", "운영경비(천원)"}
                if required_columns.issubset(df_meeting_upload.columns):
                    registered_committees = fetch_data(
                        "SELECT 위원회명 FROM committees",
                        columns=["위원회명"]
                    )
                    registered_committee_names = (
                        registered_committees["위원회명"].astype("string").fillna("").str.strip().tolist()
                        if not registered_committees.empty else []
                    )

                    df_meeting_upload = df_meeting_upload.copy()
                    df_meeting_upload["위원회명"] = df_meeting_upload["위원회명"].astype("string").fillna("").str.strip()
                    invalid_meeting_rows = df_meeting_upload[
                        ~df_meeting_upload["위원회명"].isin(registered_committee_names)
                    ]

                    if not invalid_meeting_rows.empty:
                        invalid_committee_list = ", ".join(invalid_meeting_rows["위원회명"].dropna().unique().tolist())
                        st.warning(f"등록되지 않은 위원회명은 반영되지 않습니다: {invalid_committee_list}")

                    valid_meeting_rows = df_meeting_upload[
                        df_meeting_upload["위원회명"].isin(registered_committee_names)
                    ].copy()

                    if st.button("회의·운영경비 반영하기", use_container_width=True, key="meeting_yearly_update"):
                        updated = 연도별_회의운영_갱신(valid_meeting_rows)
                        if updated:
                            st.success(f"연도별 회의·운영경비 데이터가 {updated}건 반영되었습니다.")
                            time.sleep(1)
                            st.rerun()
                        else:
                            st.error("반영할 수 있는 유효한 데이터가 없습니다.")
                else:
                    st.error("엑셀 파일에는 [연도, 위원회명, 대면회의, 서면회의, 운영경비(천원)] 컬럼이 모두 포함되어야 합니다.")

        with col3:

            st.subheader("3. 위원회 현황 관리")
            st.caption("1. 등록된 개별 위원회정보를 삭제할 수 있습니다.")
            st.expander("위원회별 현황을 등록하거나 수정할 수 있습니다.")
            join_query_1 = """
                            SELECT d.순서,d.부서명,c.위원회명, m.갱신일
                            FROM committees AS c
                            LEFT JOIN departments AS d ON c.부서명 = d.부서명
                            LEFT JOIN commissioners AS m ON c.위원회명 = m.위원회명
                            ORDER BY d.순서
                            """
                
            join_query_2 = """
                            SELECT m.순서,d.부서명,c.위원회명, m.구분, m.소속, m.직위, m.성명, m.생년월일, m.위촉일자, m.만료일자, m.임기, m.성별
                            FROM committees AS c
                            LEFT JOIN departments AS d ON c.부서명 = d.부서명
                            LEFT JOIN commissioners AS m ON c.위원회명 = m.위원회명
                            ORDER BY d.순서
                            """

            df_committees_status_1 = 테이블_불러오기(join_query_1)
            df_committees_status_2 = 테이블_불러오기(join_query_2)

            ###부서 추가시 오류 발생으로 수정된 코드###
            valid_committees = df_committees_status_1.dropna(subset=['부서명'])
            부서별_위원회_개수 = valid_committees.groupby('부서명')['위원회명'].nunique().to_dict()

            department_order = (
                df_committees_status_1[['부서명', '순서']]
                .drop_duplicates()
                .dropna(subset=['부서명'])
                .sort_values('순서')
            )

            부서명_options = [
                f"{부서} 운영 위원회({부서별_위원회_개수.get(부서, 0)}개)"
                for 부서 in department_order['부서명']
            ]


            # 부서별_위원회_개수 = df_committees_status_1.groupby('부서명')['위원회명'].nunique()

            # # 부서명과 순서를 사용하여 정렬된 옵션 생성
            # department_order = df_committees_status_1[['부서명', '순서']].drop_duplicates().sort_values('순서')
        
            # # 부서명과 개수를 함께 표시할 리스트 준비
            # 부서명_options = [f"{부서} 운영 위원회({부서별_위원회_개수[부서]}개)" for 부서 in department_order['부서명']]

            col1, col2 = st.columns(2)
            선택된_위원회명 = None

            with col1:
                # 부서명을 선택할 수 있는 셀렉트 박스를 생성
                if 부서명_options:
                    선택된_부서명_with_count = st.selectbox('부서를 선택하세요', 부서명_options, label_visibility="collapsed")
                    
                    # 선택된 부서명의 실제 이름 추출
                    선택된_부서명 = 선택된_부서명_with_count.split(" ")[0]

                if 선택된_부서명:
                    선택된_부서명_df = df_committees_status_1[df_committees_status_1['부서명'] == 선택된_부서명].drop_duplicates(subset='위원회명')
                    if not 선택된_부서명_df.empty:
                        st.dataframe(선택된_부서명_df, hide_index=True, column_config={
                            '순서': {'width': 50}, '부서명': {'width': 120}, '위원회명': {'width': 180}, '갱신일': {'width': 100}},height=290)
                    else:
                        st.info("등록된 위원회 및 위원 정보가 없습니다.")
                else:
                    st.error("부서를 선택해 주세요.")

                #     # 선택된 부서의 데이터 필터링 및 표시
                #     if 선택된_부서명:
                #         선택된_부서명_df = df_committees_status_1[df_committees_status_1['부서명'] == 선택된_부서명]
                #         # 위원회명에 대해 중복 제거
                #         선택된_부서명_df = 선택된_부서명_df.drop_duplicates(subset='위원회명')
                #     else:
                #         st.error("부서를 선택해 주세요.")
                # else:
                #     st.info("부서 정보를 등록하세요.")

                선택된_부서명_df = df_committees_status_2[df_committees_status_2['부서명'] == 선택된_부서명]

                선택된_위원회명_df = 선택된_부서명_df[선택된_부서명_df['위원회명'] == 선택된_위원회명]

                # 선택된 부서에 속한 위원회 목록 추출
                위원회명_options = 선택된_부서명_df['위원회명'].unique()
                # print(위원회명_options)

            with col2:
                if not 선택된_부서명_df.empty:
                    # 셀렉트 박스에서 위원회를 선택
                    선택된_위원회명 = st.selectbox('위원회를 선택하세요', 위원회명_options, label_visibility="collapsed")

                    # 선택된 위원회의 데이터 필터링
                    선택된_위원회명_df = 선택된_부서명_df[선택된_부서명_df['위원회명'] == 선택된_위원회명]

                    # # '부서명'과 '위원회명' 열을 데이터프레임에서 제거
                    # 선택된_위원회명_df = 선택된_위원회명_df.drop(columns=['부서명', '위원회명'])

                    # '순서' 컬럼을 기준으로 데이터프레임 정렬
                    선택된_위원회명_df = 선택된_위원회명_df.sort_values(by='순서')
                    # 필터링된 데이터를 데이터프레임으로 표시
                    st.dataframe(선택된_위원회명_df, hide_index=True, height=290, column_config={
                        '순서': {'width': 40},
                        '구분': {'width': 50},
                        '소속': {'width': 60},
                        '직위': {'width': 110},
                        '성명': {'width': 70},
                        '생년월일': {'width': 100},
                        '위촉일자': {'width': 100},
                        '만료일자': {'width': 100},
                        '임기': {'width': 50},
                        '성별': {'width': 50}
                    })
                else:
                    st.info("선택한 부서에 등록된 위원회 및 위원 정보가 없습니다.")

            st.caption("2. 위원회 서식 엑셀 파일로 다운로드할 수 있습니다.")

            query = """
            SELECT 순서, 위원회명, 구분, 소속, 직위, 성명, 생년월일, 위촉일자, 만료일자, 임기, 성별
            FROM commissioners
            """
            columns = ["순서","위원회명", "구분", "소속", "직위", "성명", "생년월일", "위촉일자", "만료일자", "임기", "성별"]
            df_committee_db = fetch_data(query, columns=columns)

            # 엑셀 파일 변환 및 준비
            committee_excel = to_excel(df_committee_db, sheet_name="위원회_명단")


            st.download_button(
                label="개별 위원회 현황 다운로드",
                data=generate_excel_file(선택된_위원회명),
                file_name=f"{선택된_부서명}({선택된_위원회명})_{today}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

            st.caption("3. 개별 위원회 엑셀 파일을 업로드하세요.")
            committee_upload = st.file_uploader("개별 위원회 엑셀 파일을 업로드하세요", type=["xlsx"], key="committees", label_visibility="collapsed")

            if committee_upload:

                df_uploaded = pd.read_excel(committee_upload, sheet_name="위원회_명단")
    
                st.write("업로드된 데이터:")
                st.dataframe(df_uploaded, hide_index=True, height=300)

                required_columns = {"순서", "위원회명", "구분", "소속", "직위", "성명", "위촉일자", "만료일자", "임기", "성별"}

                if required_columns.issubset(df_uploaded.columns):
                    # 유일한 위원회명을 배열로 가져오기
                    입력한_위원회명 = df_uploaded["위원회명"].unique()
                    선택된_위원회명_arr = 선택된_위원회명_df["위원회명"].unique()

                    # 버튼을 눌렀을 때의 동작 정의
                    if st.button("위원회 갱신자료 반영", use_container_width=True):
                        # 입력 데이터와 선택 데이터의 첫 번째 요소를 비교하여 같으면 진행
                        if len(입력한_위원회명) > 0 and len(선택된_위원회명_arr) > 0 and 입력한_위원회명[0] == 선택된_위원회명_arr[0]:

                            # 데이터베이스에서 기존 데이터 삭제
                            delete_query = "DELETE FROM commissioners WHERE 위원회명 = ?"
                            c.execute(delete_query, (선택된_위원회명_arr[0],))
                            conn.commit()

                            # 삭제 완료 메시지 표시
                            st.success(f"{선택된_위원회명_arr[0]} 기존 자료를 삭제합니다.")
                            time.sleep(1)

                            # 새로운 위원회 데이터 등록
                            success = 위원회별_위원_등록(df_uploaded)

                            if success:
                                refresh_db_connection()
                                # 등록 성공 메시지 및 페이지 새로고침
                                st.success(f"{입력한_위원회명[0]} 성공적으로 갱신되었습니다!")
                                time.sleep(1)
                                st.rerun()
                            else:
                                # 등록 실패 메시지
                                st.error("데이터베이스 업데이트 중 오류가 발생했습니다.")
                        else:
                            # 조건 불일치 에러 메시지
                            st.error("위원회명이 일치하지 않습니다. 다시 확인해 주세요.")

                else:
                    st.error(f"엑셀 파일에는 {required_columns} 컬럼이 모두 포함되어 있어야 합니다.")

            st.caption("4. 개별 위원회 현황을 삭제할 수 있습니다.")
            
            if st.button("개별 위원회 현황 삭제하기", use_container_width=True):
                if 선택된_위원회명 is None:
                    st.warning("위원회 정보를 등록하세요.")  # 위원회명이 None일 경우 경고 메시지
                
                else:
                        # SQL DELETE 쿼리 실행
                        delete_query = "DELETE FROM commissioners WHERE 위원회명 = ?"
                        # 데이터베이스에 쿼리를 실행하는 함수를 호출합니다.
                        c.execute(delete_query, (선택된_위원회명,))
                        conn.commit()
                        refresh_db_connection()
                        st.success(f"{선택된_위원회명} 위원회가 성공적으로 삭제되었습니다.")
                        time.sleep(1)
                        st.rerun()

            st.caption("5. 부서별 위원회 갱신 파일을 다운로드할 수 있습니다.")                
                    
            if st.download_button(
                label="부서별 위원회 갱신 파일 다운로드",
                data=위원회별_zip_파일생성(),
                file_name=f"부서별_위원회_위원회현황_{today}.zip",
                mime="application/zip",
                use_container_width=True
                ):
                st.success("다운로드가 완료되었습니다.")

            st.caption("6. 기준일별 전체 위원회 위촉 현황을 다운로드할 수 있습니다.")
            st.caption(f"현재 DB에 저장된 위촉일자와 만료일자를 기준으로 계산하며, 기준일은 {기준일_고정값.strftime('%Y-%m-%d')}입니다.")

            기준일_탭목록 = 기준일_옵션_목록()
            history_tabs = st.tabs([label for label, _, _ in 기준일_탭목록])

            for history_tab, (label, 기준일, key_suffix) in zip(history_tabs, 기준일_탭목록):
                with history_tab:
                    기준일_표시 = 기준일.strftime("%Y-%m-%d")
                    기준일_전체_df = 기준일_위원회_위촉현황_데이터(기준일=기준일).sort_values(
                        by=['부서명', '위원회명', '순서'],
                        na_position='last'
                    )

                    download_col1, download_col2 = st.columns([2.1, 1])
                    with download_col1:
                        st.caption(f"기준일: {기준일_표시}")
                    with download_col2:
                        st.download_button(
                            label=f"{label} 전체 위촉 현황 다운로드",
                            data=generate_committee_snapshot_excel_file(기준일=기준일, 기준라벨=label),
                            file_name=f"전체위원회_위촉현황_{key_suffix}_{기준일_표시}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                            key=f"admin_committee_history_download_all_{key_suffix}"
                        )

                    if 기준일_전체_df.empty:
                        st.info(f"{기준일_표시} 기준으로 유효한 위촉 현황이 없습니다.")
                    else:
                        st.dataframe(기준일_전체_df, hide_index=True, height=300, column_config={
                            '순서': {'width': 40},
                            '부서명': {'width': 120},
                            '위원회명': {'width': 180},
                            '구분': {'width': 50},
                            '소속': {'width': 90},
                            '직위': {'width': 110},
                            '성명': {'width': 70},
                            '생년월일': {'width': 100},
                            '위촉일자': {'width': 100},
                            '만료일자': {'width': 100},
                            '임기': {'width': 50},
                            '성별': {'width': 50}
                        })
            # else:
            #     st.warning("부서 및 위원회 데이터를 등록하시기 바랍니다.")

    else:
        st.info("관리자 암호를 입력하면 관리자 기능을 사용할 수 있습니다.")


# def get_committee_installation_status():
#     installation_query = """
#         SELECT
#             c.위원회명,
#             c.설치일자,
#             SUM(CASE WHEN cm.구분 = '당연직' THEN 1 ELSE 0 END) AS 당연직_위원,
#             SUM(CASE WHEN cm.구분 = '위촉직' THEN 1 ELSE 0 END) AS 위촉직_위원
#         FROM committees AS c
#         LEFT JOIN commissioners AS cm ON c.위원회명 = cm.위원회명
#         GROUP BY c.위원회명, c.설치일자
#         ORDER BY c.위원회명
#     """
#     columns = ["위원회명", "설치일자", "당연직_위원", "위촉직_위원"]
#     df = fetch_data(installation_query, columns=columns)
#     if not df.empty:
#         df["전체 위원"] = df["당연직_위원"] + df["위촉직_위원"]
#         df["설치일자"] = df["설치일자"].fillna("미등록")
#     return df

# with tab1:
#     # 기존 섹션 위에 추가하거나 아래쪽에 위치
#     st.markdown("## 위원회 설치 현황")
#     df_installation = get_committee_installation_status()
#     if not df_installation.empty:
#         st.dataframe(
#             df_installation,
#             hide_index=True,
#             column_config={
#                 "위원회명": {"width": 250},
#                 "설치일자": {"width": 110},
#                 "당연직_위원": {"width": 120},
#                 "위촉직_위원": {"width": 120},
#                 "전체 위원": {"width": 120},
#             },
#             use_container_width=True
#         )
#     else:
#         st.info("위원회 설치 정보가 없습니다.")
