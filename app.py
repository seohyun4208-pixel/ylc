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

    # ---------- 1. 채팅 파싱 ----------
    pattern = re.compile(
        r"(\d{4}-\d{2}-\d{2}) (\d{2}:\d{2}:\d{2}) 시작 .*? 수신자 모든 사람:\n\t(.+)"
    )
    raw_entries = []
    for date_str, time_str, content in pattern.findall(
