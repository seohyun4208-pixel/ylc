import streamlit as st
import re
import pandas as pd
from datetime import datetime, timedelta
import io

st.title("경남지부 출석 자동 체크 프로그램")

# ---------- 사용자 입력 ----------
session_date = st.date_input("강연 날짜")
start_notice = st.time_input("시작 출석 공지 시각")
end_notice = st.time_input("종료 출석 공지 시각")
valid_min = st.number_input("유효 시간(분)", value=10, min_value=1)

zoom_file = st.file_uploader("줌 채팅 txt 파일 업로드", type="txt")
excel_file = st.file_uploader("기존 명단 엑셀 업로드 (경남지부 시트 포함)", type="xlsx")

if zoom_file and excel_file:
    raw_text = zoom_file.read().decode("utf-8")
    raw_text = raw_text.replace("\r\n", "\n").replace("\r", "\n")

    pattern = re.compile(r"(\d{4}-\d{2}-\d{2}) (\d{2}:\d{2}:\d{2}) 시작 .*? 수신자 모든 사람:\n\t(.+)")
    raw_entries = []
    for date_str, time_str, content in pattern.findall(raw_text):
        raw_entries.append({"시간": time_str, "원문": content.strip()})

    def parse_content(row):
        c = row["원문"]
        if not c.startswith("경남지부/"):
            return None
        parts = c.split("/")
        if len(parts) != 2:
            return {"이름": None, "상태": "형식오류", "시간": row["시간"], "원문": c}
        return {"이름": parts[1].strip(), "상태": "정상", "시간": row["시간"]}

    parsed = [parse_content(r) for r in raw_entries]
    parsed = [p for p in parsed if p is not None]
    df_chat = pd.DataFrame(parsed)

    format_errors = df_chat[df_chat["상태"] == "형식오류"] if not df_chat.empty else pd.DataFrame()
    df_chat = df_chat[df_chat["상태"] == "정상"] if not df_chat.empty else df_chat

    xls = pd.ExcelFile(excel_file)
    if "경남지부" not in xls.sheet_names:
        st.error("경남지부 시트를 찾을 수 없습니다.")
        st.stop()
    df_members = pd.read_excel(xls, sheet_name="경남지부")
    df_members.columns = df_members.columns.str.strip()
    df_members["이름"] = df_members["이름"].astype(str).str.strip()

    member_names = set(df_members["이름"])
    unknown_names = set(df_chat["이름"]) - member_names if not df_chat.empty else set()

    start_dt = datetime.combine(session_date, start_notice)
    end_dt = datetime.combine(session_date, end_notice)
    midpoint = start_dt + (end_dt - start_dt) / 2

    def classify(name, notice_time, chat_df, lower_bound=None, upper_bound=None):
        notice_dt = datetime.combine(session_date, notice_time)
        window_end = notice_dt + timedelta(minutes=valid_min)
        rows = chat_df[chat_df["이름"] == name] if not chat_df.empty else pd.DataFrame()
        if rows.empty:
            return "불참"

        times = [datetime.combine(session_date, datetime.strptime(t, "%H:%M:%S").time()) for t in rows["시간"]]

        if lower_bound is not None:
            times = [t for t in times if t >= lower_bound]
        if upper_bound is not None:
            times = [t for t in times if t < upper_bound]

        if not times:
            return "불참"

        earliest = min(times)
        if earliest < notice_dt:
            return "오류(공지전송신)"
        elif earliest <= window_end:
            return "참여"
        else:
            late_min = int((earliest - window_end).total_seconds() // 60) + 1
            return f"지각({late_min}분)"

    def base_status(status):
        if status.startswith("지각"):
            return "지각"
        if status.startswith("오류"):
            return "오류"
        return status

    results = []
    manual_review = []

    for name in df_members["이름"]:
        s_status = classify(name, start_notice, df_chat, upper_bound=midpoint)
        e_status = classify(name, end_notice, df_chat, lower_bound=midpoint)
        s_base, e_base = base_status(s_status), base_status(e_status)

        if s_base == "오류" or e_base == "오류":
            result, point = "수동확인", None
            manual_review.append({
                "이름": name,
                "시작출석": s_status,
                "종료출석": e_status,
                "사유": "공지 시각 이전 전송 감지"
            })
        elif s_base == "불참" or e_base == "불참":
            result, point = "X", 0
        elif s_base == "지각" or e_base == "지각":
            result, point = "세모", 2.5
        else:
            result, point = "O", 5

        results.append({
            "이름": name,
            "시작출석": s_status,
            "종료출석": e_status,
            "결과": result,
            "AP point": point
        })

    df_result = pd.DataFrame(results)
    df_manual = pd.DataFrame(manual_review)

    date_label = session_date.strftime("%m/%d")
    df_final = df_members.merge(df_result, on="이름", how="left")
    df_final = df_final.rename(columns={
        "시작출석": f"{date_label} 시작출석",
        "종료출석": f"{date_label} 종료출석",
        "결과": f"{date_label} 결과",
        "AP point": f"{date_label} AP"
    })

    st.subheader("출석 결과 미리보기")
    st.dataframe(df_final)

    if not df_manual.empty:
        st.error("⚠️ 아래 인원은 '공지 시각 이전 전송'으로 감지되어 수동 확인이 필요합니다:")
        st.dataframe(df_manual)

    if not format_errors.empty:
        st.warning("형식 오류가 있는 댓글이 있습니다 (수동 처리 필요):")
        st.dataframe(format_errors)

    if unknown_names:
        st.warning(f"명단에 없는 이름이 댓글에 있습니다 (오탈자 의심): {', '.join(unknown_names)}")

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet in xls.sheet_names:
            if sheet == "경남지부":
                df_final.to_excel(writer, sheet_name=sheet, index=False)
            else:
                pd.read_excel(xls, sheet_name=sheet).to_excel(writer, sheet_name=sheet, index=False)

    st.download_button(
        "결과 엑셀 다운로드",
        data=output.getvalue(),
        file_name=f"경남지부_출석부_{date_label}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
