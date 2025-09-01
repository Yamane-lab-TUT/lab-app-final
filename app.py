# --------------------------------------------------------------------------
# Yamane Lab Convenience Tool - Streamlit Application (v7.0 - Final Version)
#
# v7.0:
# - Fixes all previous SyntaxErrors and NameErrors.
# - Implements @st.cache_data for efficient Google Sheets API calls.
# - Addresses the Google Drive storage quota error by using Shared Drives.
# - Integrates all requested features into a robust, single-file structure.
# --------------------------------------------------------------------------

import streamlit as st
import gspread
import pandas as pd
import os
import io
import re
import json
import base64
import MimeText
from datetime import datetime, time, timedelta
from urllib.parse import quote as url_quote, urlencode
import numpy as np
import matplotlib.pyplot as plt

# Google API client libraries
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from googleapiclient.errors import HttpError

# --- Global Configuration & Setup ---
st.set_page_config(page_title="山根研 便利屋さん", layout="wide")

try:
    plt.rcParams['font.family'] = 'Meiryo'
except:
    try:
        plt.rcParams['font.family'] = 'Yu Gothic'
    except:
        plt.rcParams['font.family'] = 'sans-serif'
plt.rcParams['font.weight'] = 'bold'; plt.rcParams['axes.labelweight'] = 'bold'
plt.rcParams['axes.linewidth'] = 1.5; plt.rcParams['xtick.major.width'] = 1.5
plt.rcParams['ytick.major.width'] = 1.5; plt.rcParams['font.size'] = 14
plt.rcParams['axes.unicode_minus'] = False

# Google Cloud related settings
# IMPORTANT: Use your actual credentials and folder IDs here.
SERVICE_ACCOUNT_FILE = 'research-lab-app-42f3c0b5d5b1.json'
SPREADSHEET_NAME = 'エピノート'
FOLDER_IDS = {
    'EP_D1': '1KQEeEsHChqtrAIvP91ILnf6oS4fTVi1p',
    'EP_D2': '1inmARuM_SgiYHi4PR7rcWRH0jERKZVJy',
    'MT': '1YllkIwYuV3IqY4_i0YoyY43SAB-U8-0i',
    'MINUTES': '1g7qiEFuEchsFFBKFJwxN2D2PjShuDtzM',
    'HANDOVER': '1Mr70YjsgCzMboD7UZStm7bE8LQs1mwFu',
    'QA': '1cil7cMFmQlgfzqOD-8QOm4KqVB4Emy79'
}
DEFAULT_CALENDAR_ID = 'yamane.lab.6747@gmail.com'
INQUIRY_RECIPIENT_EMAIL = 'kyuno.yamato.ns@tut.ac.jp'

# --- Initialize Google Services ---
@st.cache_resource(show_spinner="Googleサービスに接続中...")
def initialize_google_services():
    try:
        scopes = [
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive',
            'https://www.googleapis.com/auth/calendar'
        ]
        
        creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=scopes)
        gc = gspread.service_account(filename=SERVICE_ACCOUNT_FILE)

        drive_service = build('drive', 'v3', credentials=creds)
        calendar_service = build('calendar', 'v3', credentials=creds)
        
        return gc, drive_service, calendar_service
    except Exception as e:
        st.error(f"❌ 致命的なエラー: サービスの初期化に失敗しました: {e}"); st.stop()

gc, drive_service, calendar_service = initialize_google_services()

# --- Utility Functions ---

@st.cache_data(ttl=300, show_spinner="シート「{sheet_name}」を読み込み中...")
def get_sheet_as_df(_gc, spreadsheet_name, sheet_name):
    try:
        spreadsheet = _gc.open(spreadsheet_name)
        worksheet = spreadsheet.worksheet(sheet_name)
        data = worksheet.get_all_values()
        if not data or len(data) < 1: return pd.DataFrame()
        headers = data[0]
        df = pd.DataFrame(data[1:], columns=headers)
        return df
    except gspread.exceptions.WorksheetNotFound:
        st.error(f"スプレッドシート内にシート名「{sheet_name}」が見つかりません。"); return pd.DataFrame()
    except Exception as e:
        st.warning(f"シート「{sheet_name}」の読込中にエラー: {e}"); return pd.DataFrame()

def upload_file_to_drive(service, file_uploader_obj, folder_id, memo_content=""):
    if not file_uploader_obj: return "", ""
    try:
        with st.spinner(f"'{file_uploader_obj.name}'をアップロード中..."):
            timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
            file_extension = os.path.splitext(file_uploader_obj.name)[1]
            sanitized_memo = re.sub(r'[\\/:*?"<>|\r\n]+', '', memo_content)[:50] if memo_content else "無題"
            new_filename = f"{sanitized_memo} ({timestamp}){file_extension}"
            file_metadata = {'name': new_filename, 'parents': [folder_id]}
            media = MediaIoBaseUpload(io.BytesIO(file_uploader_obj.getvalue()), mimetype=file_uploader_obj.type, resumable=True)
            file = service.files().create(body=file_metadata, media_body=media, fields='id, webViewLink').execute()
        st.success(f"📄 ファイル '{new_filename}' をアップロードしました。"); return new_filename, file.get('webViewLink')
    except Exception as e:
        st.error(f"ファイルアップロード中にエラー: {e}"); return "アップロード失敗", ""

def generate_gmail_link(recipient, subject, body):
    base_url = "https://mail.google.com/mail/?view=cm&fs=1"
    params = {"to": recipient, "su": subject, "body": body}
    return f"{base_url}&{urlencode(params)}"

# --- UI Page Functions (modularized for clarity and re-run safety) ---

def page_note_recording():
    st.header("📝 エピノート・メンテノートの記録")
    note_type = st.radio("どちらを登録しますか？", ("エピノート", "メンテノート"), horizontal=True)
    if note_type == "エピノート":
        with st.form("ep_note_form", clear_on_submit=True):
            ep_category = st.radio("カテゴリ", ("D1", "D2"), horizontal=True)
            ep_memo = st.text_area("メモ内容(番号など)")
            uploaded_file = st.file_uploader("エピノートの写真（必須）", type=["jpg", "jpeg", "png"])
            submitted = st.form_submit_button("エピノートを保存")
            if submitted:
                if uploaded_file:
                    folder_id = FOLDER_IDS['EP_D1'] if ep_category == "D1" else FOLDER_IDS['EP_D2']
                    filename, url = upload_file_to_drive(drive_service, uploaded_file, folder_id, ep_memo)
                    if url:
                        row_data = [datetime.now().strftime("%Y%m%d_%H%M%S"), "エピノート", ep_category, ep_memo, filename, url]
                        spreadsheet = gc.open(SPREADSHEET_NAME); spreadsheet.worksheet('エピノート_データ').append_row(row_data)
                        st.success("エピノートを保存しました！"); st.cache_data.clear(); st.rerun()
                else: st.error("写真をアップロードしてください。")
    elif note_type == "メンテノート":
        with st.form("mt_note_form", clear_on_submit=True):
            mt_memo = st.text_area("メモ内容（日付など）")
            uploaded_file = st.file_uploader("関連写真", type=["jpg", "jpeg", "png"])
            submitted = st.form_submit_button("メンテノートを保存")
            if submitted:
                if not mt_memo: st.error("メモ内容を入力してください。")
                else:
                    filename, url = upload_file_to_drive(drive_service, uploaded_file, FOLDER_IDS['MT'], mt_memo)
                    row_data = [datetime.now().strftime("%Y%m%d_%H%M%S"), "メンテノート", mt_memo, filename, url]
                    spreadsheet = gc.open(SPREADSHEET_NAME); spreadsheet.worksheet('メンテノート_データ').append_row(row_data)
                    st.success("メンテノートを保存しました！"); st.cache_data.clear(); st.rerun()

def page_note_list():
    st.header("📓 登録済みのノート一覧")
    note_display_type = st.radio("表示するノート", ("エピノート", "メンテノート"), horizontal=True, key="note_display_type")
    
    if note_display_type == "エピノート":
        st.markdown("#### エピノート一覧")
        df_ep = get_sheet_as_df(gc, SPREADSHEET_NAME, 'エピノート_データ')
        required_cols = ['タイムスタンプ', '種類', 'カテゴリ', 'メモ', '写真ファイル名', '写真URL']
        if df_ep.empty or not all(col in df_ep.columns for col in required_cols):
            st.warning(f"エピノートシートのデータがありません、またはヘッダー形式が正しくありません。"); return
        
        ep_category_filter = st.selectbox("カテゴリで絞り込み", ["すべて"] + list(df_ep['カテゴリ'].unique()))
        
        filtered_df = df_ep.sort_values(by='タイムスタンプ', ascending=False)
        if ep_category_filter != "すべて":
            filtered_df = filtered_df[filtered_df['カテゴリ'] == ep_category_filter]
        
        if filtered_df.empty:
            st.info(f"検索条件に一致するノートはありません。"); return

        selected_row_idx = st.selectbox(
            "ノートを選択",
            options=["---"] + filtered_df.index.tolist(),
            format_func=lambda idx: "---" if idx == "---" else f"{filtered_df.loc[idx, 'メモ'][:40]}" + ("..." if len(filtered_df.loc[idx, 'メモ']) > 40 else ""),
            key="select_ep_note_view"
        )
        
        if selected_row_idx != "---":
            selected_row = filtered_df.loc[selected_row_idx]
            st.subheader(f"詳細: {selected_row['タイムスタンプ']}")
            st.write(f"**カテゴリ:** {selected_row['カテゴリ']}")
            st.write(f"**メモ:**"); st.text(selected_row['メモ'])
            if selected_row['写真URL']:
                st.markdown(f"**写真:** [ファイルを開く]({selected_row['写真URL']})", unsafe_allow_html=True)

    elif note_display_type == "メンテノート":
        st.markdown("#### メンテノート一覧")
        df_mt = get_sheet_as_df(gc, SPREADSHEET_NAME, 'メンテノート_データ')
        required_cols = ['タイムスタンプ', '種類', 'メモ', '写真ファイル名', '写真URL']
        if df_mt.empty or not all(col in df_mt.columns for col in required_cols):
            st.warning(f"メンテノートシートのデータがありません、またはヘッダー形式が正しくありません。"); return

        if df_mt.empty: st.info("まだメンテノートは登録されていません。"); return
        
        filtered_df = df_mt.sort_values(by='タイムスタンプ', ascending=False)
        
        if filtered_df.empty: st.info(f"検索条件に一致するノートはありません。"); return

        selected_row_idx = st.selectbox(
            "ノートを選択",
            options=["---"] + filtered_df.index.tolist(),
            format_func=lambda idx: "---" if idx == "---" else f"{filtered_df.loc[idx, 'メモ'][:40]}" + ("..." if len(filtered_df.loc[idx, 'メモ']) > 40 else ""),
            key="select_mt_note_view"
        )

        if selected_row_idx != "---":
            selected_row = filtered_df.loc[selected_row_idx]
            st.subheader(f"詳細: {selected_row['タイムスタンプ']}")
            st.write(f"**メモ:**"); st.text(selected_row['メモ'])
            if selected_row['写真URL']:
                st.markdown(f"**写真:** [ファイルを開く]({selected_row['写真URL']})", unsafe_allow_html=True)


def page_calendar():
    st.header("📅 Googleカレンダーの管理")
    tab1, tab2 = st.tabs(["予定の確認", "新しい予定の追加"])
    with tab1:
        st.subheader("期間を指定して予定を表示")
        calendar_url = f"https://calendar.google.com/calendar/u/0/r?cid={DEFAULT_CALENDAR_ID}"
        st.markdown(f"**[Googleカレンダーで直接開く]({calendar_url})**", unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        start_date = col1.date_input("開始日", datetime.today().date())
        end_date = col2.date_input("終了日", datetime.today().date() + timedelta(days=7))
        if st.button("予定を読み込む"):
            if start_date > end_date: st.error("終了日は開始日以降に設定してください。")
            else: display_calendar_events(calendar_service, DEFAULT_CALENDAR_ID, start_date, end_date)
    with tab2:
        st.subheader("新しい予定を追加")
        with st.form("add_event_form", clear_on_submit=True):
            group_types = ["輻射G", "Ge-family", "中性子G"]
            selected_group_type = st.selectbox("グループ名", group_types)
            event_types = ["エピ", "XRD", "フォトリソ", "PL", "AFM", "蒸着", "アニール", "その他"]
            selected_event_type = st.selectbox("予定の種類", event_types)
            event_summary_base = selected_event_type if selected_event_type != "その他" else st.text_input("予定のタイトル (その他)", key="other_event_title")
            event_summary = f"{selected_group_type}_{event_summary_base}"
            col1, col2 = st.columns(2)
            event_date = col1.date_input("日付 *", datetime.today().date())
            is_allday = col2.checkbox("終日", value=False)
            if not is_allday:
                col3, col4 = st.columns(2)
                start_time, end_time = col3.time_input("開始時刻 *", time(9, 0)), col4.time_input("終了時刻 *", time(10, 0))
            event_location = st.text_input("場所"); event_description = st.text_area("説明")
            submitted = st.form_submit_button("カレンダーに追加")
            if submitted:
                if not event_summary_base: st.error("件名は必須です。")
                else:
                    if is_allday: start, end = {'date': event_date.isoformat()}, {'date': (event_date + timedelta(days=1)).isoformat()}
                    else:
                        tz = "Asia/Tokyo"; start = {'dateTime': datetime.combine(event_date, start_time).isoformat(), 'timeZone': tz}; end = {'dateTime': datetime.combine(event_date, end_time).isoformat(), 'timeZone': tz}
                    event_body = {'summary': event_summary, 'location': event_location, 'description': event_description, 'start': start, 'end': end}
                    try:
                        created_event = calendar_service.events().insert(calendarId=DEFAULT_CALENDAR_ID, body=event_body).execute()
                        st.success(f"予定「{created_event.get('summary')}」を追加しました。"); st.markdown(f"[カレンダーで確認]({created_event.get('htmlLink')})")
                    except HttpError as e: st.error(f"予定の追加に失敗しました: {e}")

def page_minutes():
    st.header("🎙️ 会議の議事録の管理")
    minutes_sheet_name = '議事録_データ'
    tab1, tab2 = st.tabs(["議事録の確認", "新しい議事録の登録"])
    with tab1:
        st.subheader("登録済みの議事録")
        df_minutes = get_sheet_as_df(gc, SPREADSHEET_NAME, minutes_sheet_name)
        required_cols = ['タイムスタンプ', '会議タイトル', '音声ファイル名', '音声ファイルURL', '議事録内容']
        if df_minutes.empty or not all(col in df_minutes.columns for col in required_cols):
            st.warning(f"議事録シートのデータがありません、またはヘッダー形式が正しくありません。"); return
        
        options = {f"{row['タイムスタンプ']} - {row['会議タイトル']}": idx for idx, row in df_minutes.iterrows()}
        selected_key = st.selectbox("議事録を選択", ["---"] + list(options.keys()))
        if selected_key != "---":
            selected_row = df_minutes.loc[options[selected_key]]
            st.subheader(selected_row['会議タイトル'])
            st.caption(f"登録日時: {selected_row['タイムスタンプ']}")
            if selected_row['音声ファイルURL']:
                st.markdown(f"**[音声ファイルを開く]({selected_row['音声ファイルURL']})** ({selected_row['音声ファイル名']})")
            st.markdown("---")
            st.markdown(selected_row['議事録内容'])
    with tab2:
        st.subheader("新しい議事録を登録")
        with st.form("minutes_form", clear_on_submit=True):
            title = st.text_input("会議のタイトル *")
            audio_file = st.file_uploader("音声ファイル (任意)", type=["mp3", "wav", "m4a", "flac"])
            content = st.text_area("議事録内容", height=300)
            submitted = st.form_submit_button("議事録を保存")
            if submitted:
                if not title: st.error("タイトルは必須です。")
                else:
                    filename, url = upload_file_to_drive(drive_service, audio_file, FOLDER_IDS['MINUTES'], title)
                    row_data = [datetime.now().strftime("%Y%m%d_%H%M%S"), title, filename, url, content]
                    spreadsheet = gc.open(SPREADSHEET_NAME); spreadsheet.worksheet(minutes_sheet_name).append_row(row_data)
                    st.success("議事録を保存しました。"); st.cache_data.clear(); st.rerun()

def page_qa():
    st.header("💡 山根研 知恵袋")
    qa_sheet_name = '知恵袋_データ'; answers_sheet_name = '知恵袋_解答'
    
    # NEW: Simple filtering via selectbox instead of tabs
    qa_status_filter = st.selectbox("表示する質問のステータス", ["すべての質問", "未解決のみ", "解決済みのみ"])

    st.subheader("質問と回答を見る")
    df_qa = get_sheet_as_df(gc, SPREADSHEET_NAME, qa_sheet_name)
    df_answers = get_sheet_as_df(gc, SPREADSHEET_NAME, answers_sheet_name)

    required_qa_cols = ['タイムスタンプ', '質問タイトル', '質問内容', '連絡先メールアドレス', '添付ファイル名', '添付ファイルURL', 'ステータス']
    if df_qa.empty or not all(col in df_qa.columns for col in required_qa_cols):
        st.warning(f"知恵袋_データ シートのデータがないか、ヘッダー形式が正しくありません。"); return
    
    df_qa['タイムスタンプ_dt'] = pd.to_datetime(df_qa['タイムスタンプ'], format="%Y%m%d_%H%M%S")
    df_qa = df_qa.sort_values(by='タイムスタンプ_dt', ascending=False)
    
    filtered_df_qa = df_qa.copy()
    if qa_status_filter == "未解決のみ":
        filtered_df_qa = filtered_df_qa[filtered_df_qa['ステータス'] == '未解決']
    elif qa_status_filter == "解決済みのみ":
        filtered_df_qa = filtered_df_qa[filtered_df_qa['ステータス'] == '解決済み']
        
    if filtered_df_qa.empty:
        st.info("条件に一致する質問はありません。")
        return
        
    options = {f"[{row['ステータス']}] {row['質問タイトル']} ({row['タイムスタンプ_dt'].strftime('%Y/%m/%d %H:%M:%S')})": row['タイムスタンプ'] for _, row in filtered_df_qa.iterrows()}
    selected_ts = st.selectbox("質問を選択", ["---"] + list(options.keys()))
    
    if selected_ts != "---":
        question_id = options[selected_ts]
        question = df_qa[df_qa['タイムスタンプ'] == question_id].iloc[0]
        with st.container(border=True):
            st.subheader(f"Q: {question['質問タイトル']}")
            st.caption(f"投稿日時: {question['タイムスタンプ']} | ステータス: {question['ステータス']}")
            st.markdown(question['質問内容'])
            if question['添付ファイルURL']:
                st.markdown(f"**添付ファイル:** [リンクを開く]({question['添付ファイルURL']})", unsafe_allow_html=True)
            if question['ステータス'] == '未解決':
                if st.button("この質問を解決済みにする", key=f"resolve_{question_id}"):
                    try:
                        spreadsheet = gc.open(SPREADSHEET_NAME)
                        qa_sheet_obj = spreadsheet.worksheet(qa_sheet_name)
                        cell = qa_sheet_obj.find(question_id)
                        status_col_index = qa_sheet_obj.row_values(1).index("ステータス") + 1
                        qa_sheet_obj.update_cell(cell.row, status_col_index, "解決済み")
                        st.success("ステータスを「解決済み」に更新しました。"); st.cache_data.clear(); st.rerun()
                    except Exception as e: st.error(f"更新に失敗しました: {e}")

        st.markdown("---")
        st.subheader("回答")
        answers = df_answers[df_answers['質問タイムスタンプ (質問ID)'] == question_id] if not df_answers.empty else pd.DataFrame()
        if answers.empty:
            st.info("まだ回答はありません。")
        else:
            for _, answer in answers.iterrows():
                with st.container(border=True):
                    st.markdown(f"**A:** {answer['解答内容']}"); st.caption(f"回答者: {answer['解答者 (任意)'] or '匿名'} | 日時: {answer['タイムスタンプ']}")
                    if answer['添付ファイルURL']:
                        st.markdown(f"**添付ファイル:** [リンクを開く]({answer['添付ファイルURL']})", unsafe_allow_html=True)
        
        with st.expander("回答を投稿する"):
            with st.form(f"answer_form_{question_id}", clear_on_submit=True):
                answer_content = st.text_area("回答内容 *"); answerer_name = st.text_input("回答者名（任意）"); answer_file = st.file_uploader("参考ファイル（任意）")
                submitted = st.form_submit_button("回答を投稿する")
                if submitted:
                    if not answer_content: st.warning("回答内容を入力してください。")
                    else:
                        filename, url = upload_file_to_drive(drive_service, answer_file, FOLDER_IDS['QA'], question['質問タイトル'])
                        row_data = [datetime.now().strftime("%Y%m%d_%H%M%S"), question['質問タイトル'], question_id, answer_content, answerer_name, "", filename, url]
                        spreadsheet = gc.open(SPREADSHEET_NAME); spreadsheet.worksheet(answers_sheet_name).append_row(row_data)
                        st.success("回答を投稿しました！"); st.cache_data.clear(); st.rerun()

    st.markdown("---")
    st.subheader("新しい質問を投稿する")
    with st.form("new_question_form", clear_on_submit=True):
        q_title = st.text_input("質問タイトル *"); q_content = st.text_area("質問内容 *", height=150)
        q_email = st.text_input("連絡先メールアドレス（任意）"); q_file = st.file_uploader("参考ファイル（画像など）", type=["jpg", "jpeg", "png", "pdf"])
        submitted = st.form_submit_button("質問を投稿")
        if submitted:
            if not q_title or not q_content: st.error("タイトルと内容は必須です。")
            else:
                filename, url = upload_file_to_drive(drive_service, q_file, FOLDER_IDS['QA'], q_title)
                row_data = [datetime.now().strftime("%Y%m%d_%H%M%S"), q_title, q_content, q_email, filename, url, "未解決"]
                spreadsheet = gc.open(SPREADSHEET_NAME); spreadsheet.worksheet(qa_sheet_name).append_row(row_data)
                st.success("質問を投稿しました。「質問と回答を見る」タブで確認してください。"); st.cache_data.clear(); st.rerun()


def page_handover():
    st.header("🔑 引き継ぎ情報の管理")
    handover_sheet_name = '引き継ぎ_データ'
    tab1, tab2 = st.tabs(["情報の確認", "新しい情報の登録"])
    with tab1:
        st.subheader("登録済みの引き継ぎ情報")
        df = get_sheet_as_df(gc, SPREADSHEET_NAME, handover_sheet_name)
        required_handover_cols = ['タイムスタンプ', '種類', 'タイトル', '内容1', '内容2', '内容3', 'メモ']
        if df.empty or not all(col in df.columns for col in required_handover_cols):
            st.warning(f"引き継ぎシートのデータがないか、ヘッダー形式が正しくありません。"); return
        col1, col2 = st.columns(2)
        with col1:
            unique_types = ["すべて"] + df['種類'].unique().tolist() if not df.empty else ["すべて"]
            selected_type = st.selectbox("情報の種類で絞り込み", unique_types)
        if selected_type == "すべて": filtered_df = df
        else: filtered_df = df[df['種類'] == selected_type]
        with col2:
            search_term = st.text_input("タイトルで検索")
        if search_term:
            filtered_df = filtered_df[filtered_df['タイトル'].str.contains(search_term, case=False, na=False)]
        if filtered_df.empty: st.info(f"検索条件に一致する情報はありません。"); return
        options = {f"[{row['種類']}] {row['タイトル']}": idx for idx, row in filtered_df.iterrows()}
        selected_key = st.selectbox("情報を選択", ["---"] + list(options.keys()))
        if selected_key != "---":
            selected_row = filtered_df.loc[options[selected_key]]
            st.subheader(f"{selected_row['タイトル']} の詳細")
            st.write(f"**種類:** {selected_row['種類']}")
            if selected_row['種類'] == "マニュアル":
                if selected_row['内容1']: st.markdown(f"**ファイル/URL:** [リンクを開く]({selected_row['内容1']})")
                st.write("**メモ:**"); st.text(selected_row['メモ'])
            elif selected_row['種類'] == "連絡先": st.write(f"**電話番号:** {selected_row['内容1']}"); st.write(f"**メール:** {selected_row['内容2']}"); st.write("**メモ:**"); st.text(selected_row['メモ'])
            elif selected_row['種類'] == "パスワード": st.write(f"**サービス名/場所:** {selected_row['タイトル']}"); st.write(f"**ユーザー名:** {selected_row['内容1']}"); st.write(f"**パスワード:** {selected_row['内容2']}"); st.write("**メモ:**"); st.text(selected_row['メモ'])
            else: st.write(f"**内容:**"); st.markdown(selected_row['内容1']); st.write("**メモ:**"); st.text(selected_row['メモ'])
    with tab2:
        st.subheader("新しい引き継ぎ情報を登録")
        handover_type = st.selectbox("情報の種類", ["マニュアル", "連絡先", "パスワード", "その他"])
        with st.form("handover_form", clear_on_submit=True):
            title = st.text_input("タイトル / サービス名 / 氏名 *")
            content1, content2, file = "", "", None
            if handover_type == "マニュアル": content1 = st.text_input("マニュアルのURL"); file = st.file_uploader("またはファイルをアップロード")
            elif handover_type == "連絡先": content1 = st.text_input("電話番号"); content2 = st.text_input("メールアドレス")
            elif handover_type == "パスワード": st.warning("パスワードの直接保存は非推奨です。"); content1 = st.text_input("ユーザー名"); content2 = st.text_input("パスワード", type="password")
            else: content1 = st.text_area("内容")
            memo = st.text_area("メモ（任意）")
            submitted = st.form_submit_button("保存")
            if submitted:
                if not title: st.error("タイトルは必須です。")
                else:
                    file_url = ""
                    if handover_type == "マニュアル" and file: _, file_url = upload_file_to_drive(drive_service, file, FOLDER_IDS['HANDOVER'], title)
                    final_c1 = file_url or content1
                    row_data = [datetime.now().strftime("%Y%m%d_%H%M%S"), handover_type, title, final_c1, content2, "", memo]
                    spreadsheet = gc.open(SPREADSHEET_NAME); spreadsheet.worksheet('引き継ぎ_データ').append_row(row_data); st.success("引き継ぎ情報を保存しました。"); st.cache_data.clear(); st.rerun()

def page_inquiry():
    st.header("✉️ お問い合わせフォーム")
    inquiry_sheet_name = 'お問い合わせ_データ'
    st.info("このアプリに関するご意見、ご要望、バグ報告などはこちらからお送りください。")
    with st.form("inquiry_form", clear_on_submit=True):
        category = st.selectbox("お問い合わせの種類", ["バグ報告", "機能改善要望", "その他"]); content = st.text_area("詳細内容 *", height=150); contact = st.text_input("連絡先（メールアドレスなど、返信が必要な場合）")
        submitted = st.form_submit_button("送信")
        if submitted:
            if not content: st.error("詳細内容を入力してください。")
            else:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S"); row_data = [timestamp, category, content, contact]
                spreadsheet = gc.open(SPREADSHEET_NAME); spreadsheet.worksheet(inquiry_sheet_name).append_row(row_data)
                subject = f"【研究室便利屋さん】お問い合わせ: {category}"; body = f"新しいお問い合わせが届きました。\n\n種類: {category}\n内容:\n{content}\n\n連絡先: {contact or 'なし'}\nタイムスタンプ: {timestamp}"
                gmail_link = generate_gmail_link(INQUIRY_RECIPIENT_EMAIL, subject, body)
                st.success("お問い合わせ内容を記録しました。ご協力ありがとうございます！"); st.info("管理者にすぐに伝えたい場合は以下のリンクをクリックして、Gmailで内容を送信してください。")
                st.markdown(f"**[Gmailを起動して管理者にメールを送信する]({gmail_link})**", unsafe_allow_html=True); st.cache_data.clear(); st.rerun()

def main():
    gc, drive_service, calendar_service = initialize_google_services()
    st.sidebar.header("メニュー")
    menu_options = ["ノート記録", "ノート一覧", "カレンダー", "議事録管理", "山根研知恵袋", "引き継ぎ情報", "お問い合わせフォーム"]
    selected_menu = st.sidebar.radio("機能を選択", menu_options)
    
    if selected_menu == "ノート記録": page_note_recording(); st.stop()
    elif selected_menu == "ノート一覧": page_note_list(); st.stop()
    elif selected_menu == "カレンダー": page_calendar(); st.stop()
    elif selected_menu == "議事録管理": page_minutes(); st.stop()
    elif selected_menu == "山根研知恵袋": page_qa(); st.stop()
    elif selected_menu == "引き継ぎ情報": page_handover(); st.stop()
    elif selected_menu == "お問い合わせフォーム": page_inquiry(); st.stop()

if __name__ == "__main__":
    main()
