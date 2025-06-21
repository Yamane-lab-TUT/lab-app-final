# --------------------------------------------------------------------------
# Yamane Lab Convenience Tool - Streamlit Application (v6.0 - 最終完成版)
#
# v6.0:
# - 不正なインデント文字を全て修正。
# - ローカル実行(ファイル読込)とデプロイ(Secrets読込)の両方に対応する
#   正しい認証処理を実装。
# - 全機能を統合した最終版。
# --------------------------------------------------------------------------

import streamlit as st
import gspread
import pandas as pd
import os
import io
import re
import json
from datetime import datetime, time, timedelta
from urllib.parse import quote as url_quote, urlencode
from io import BytesIO
import numpy as np
import matplotlib.pyplot as plt

# Google API クライアントライブラリ
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from googleapiclient.errors import HttpError
from google.oauth2 import service_account

# --- 1. グローバル設定 ---
st.set_page_config(page_title="山根研 便利屋さん", layout="wide")

# Matplotlibの日本語・スタイル設定
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

# --- Google Cloud 関連設定 ---
SERVICE_ACCOUNT_FILE = 'research-lab-app-42f3c0b5d5b1.json' # ローカル実行時のみ使用
SPREADSHEET_NAME = 'エピノート'
FOLDER_IDS = { 'EP_D1': '1KQEeEsHChqtrAIvP91ILnf6oS4fTVi1p', 'EP_D2': '1inmARuM_SgiYHi4PR7rcWRH0jERKZVJy', 'MT': '1YllkIwYuV3IqY4_i0YoyY43SAB-U8-0i', 'MINUTES': '1g7qiEFuEchsFFBKFJwxN2D2PjShuDtzM', 'HANDOVER': '1Mr70YjsgCzMboD7UZStm7bE8LQs1mwFu', 'QA': '1cil7cMFmQlgfzqOD-8QOm4KqVB4Emy79' }
DEFAULT_CALENDAR_ID = 'yamane.lab.6747@gmail.com'
INQUIRY_RECIPIENT_EMAIL = 'kyuno.yamato.ns@tut.ac.jp'


# --- 2. Googleサービス初期化 ---
@st.cache_resource(show_spinner="Googleサービスに接続中...")
def initialize_google_services():
    """
    Streamlit CloudのSecretsとローカルのJSONファイルの両方に対応した認証処理。
    """
    try:
        scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/calendar']
        
        # Streamlit CloudのSecretsに情報があるかチェック
        if "gcs_credentials" in st.secrets:
            # Secretsから認証情報を読み込む
            creds_dict = json.loads(st.secrets["gcs"]["gcs_credentials"])
            creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
            credentials = service_account.Credentials.from_service_account_info(creds)
            gc = gspread.service_account_from_dict(creds_dict)
        else:
            # ローカルで実行する場合（ファイルから読み込む）
            if os.path.exists(SERVICE_ACCOUNT_FILE):
                creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=scopes)
                gc = gspread.service_account(filename=SERVICE_ACCOUNT_FILE)
            else:
                st.error(f"認証ファイルが見つかりません: {SERVICE_ACCOUNT_FILE}")
                st.info("ローカルで実行する場合、app.pyと同じフォルダに認証用のJSONファイルを置いてください。")
                st.stop()

        drive_service = build('drive', 'v3', credentials=creds)
        calendar_service = build('calendar', 'v3', credentials=creds)
        
        return gc, drive_service, calendar_service
    except Exception as e:
        st.error(f"❌ 致命的なエラー: サービスの初期化に失敗しました: {e}"); st.stop()

# --- 3. ヘルパー関数 ---
@st.cache_data(ttl=300, show_spinner="シート「{sheet_name}」を読み込み中...")
def get_sheet_as_df(_gc, spreadsheet_name, sheet_name):
    try:
        spreadsheet = _gc.open(spreadsheet_name)
        worksheet = spreadsheet.worksheet(sheet_name)
        data = worksheet.get_all_values()
        if len(data) < 1: return pd.DataFrame()
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

def load_pl_data(uploaded_file):
    if uploaded_file is None: return None
    try:
        string_data = uploaded_file.getvalue().decode('utf-8')
        data = pd.read_csv(io.StringIO(string_data), comment='#', header=None, names=['pixel', 'intensity'])
        if data.isnull().values.any():
            st.error(f"ファイル「{uploaded_file.name}」内に空の行や不正なデータが含まれています。"); return None
        return data
    except Exception as e:
        st.error(f"ファイル「{uploaded_file.name}」の読み込みに失敗しました。エラー: {e}"); return None

# --- 4. UIページ関数 ---

def page_note_recording(drive_service, gc):
    st.header("📝 エピノート・メンテノートの記録")
    note_type = st.radio("どちらを登録しますか？", ("エピノート", "メンテノート"), horizontal=True)
    if note_type == "エピノート":
        with st.form("ep_note_form", clear_on_submit=True):
            ep_category = st.radio("カテゴリ", ("D1", "D2"), horizontal=True); ep_memo = st.text_area("メモ内容")
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
            mt_memo = st.text_area("メモ内容"); uploaded_file = st.file_uploader("関連写真（任意）", type=["jpg", "jpeg", "png"])
            submitted = st.form_submit_button("メンテノートを保存")
            if submitted:
                if not mt_memo: st.error("メモ内容を入力してください。")
                else:
                    filename, url = upload_file_to_drive(drive_service, uploaded_file, FOLDER_IDS['MT'], mt_memo)
                    row_data = [datetime.now().strftime("%Y%m%d_%H%M%S"), "メンテノート", mt_memo, filename, url]
                    spreadsheet = gc.open(SPREADSHEET_NAME); spreadsheet.worksheet('メンテノート_データ').append_row(row_data)
                    st.success("メンテノートを保存しました！"); st.cache_data.clear(); st.rerun()

def page_note_list(gc):
    st.header("📓 登録済みのノート一覧")
    note_display_type = st.radio("表示するノート", ("エピノート", "メンテノート"), horizontal=True, key="note_display_type")
    
    if note_display_type == "エピノート":
        df = get_sheet_as_df(gc, SPREADSHEET_NAME, 'エピノート_データ')
        required_cols = ['タイムスタンプ', 'ノート種別', 'カテゴリ', 'メモ', '写真ファイル名', '写真URL']
        if not df.empty and not all(col in df.columns for col in required_cols): st.warning(f"エピノートシートのヘッダー形式が正しくありません。"); return
        if df.empty: st.info("まだエピノートは登録されていません。"); return
        
        col1, col2 = st.columns(2)
        with col1:
            ep_category_filter = st.selectbox("カテゴリで絞り込み", ["すべて"] + list(df['カテゴリ'].unique()))
        with col2:
            search_term = st.text_input("メモの内容で検索")

        filtered_df = df.sort_values(by='タイムスタンプ', ascending=False)
        if ep_category_filter != "すべて":
            filtered_df = filtered_df[filtered_df['カテゴリ'] == ep_category_filter]
        if search_term:
            filtered_df = filtered_df[filtered_df['メモ'].str.contains(search_term, case=False, na=False)]
        
        if filtered_df.empty: st.info(f"検索条件に一致するノートはありません。"); return

        options_indices = ["---"] + filtered_df.index.tolist()
        selected_index = st.selectbox(
            "ノートを選択",
            options=options_indices,
            format_func=lambda idx: "---" if idx == "---" else f"{filtered_df.loc[idx, 'メモ'][:40]}" + ("..." if len(filtered_df.loc[idx, 'メモ']) > 40 else "")
        )
        
        if selected_index != "---":
            selected_row = filtered_df.loc[selected_index]
            st.subheader(f"詳細: {selected_row['タイムスタンプ']}")
            st.write(f"**カテゴリ:** {selected_row['カテゴリ']}"); st.write(f"**メモ:**"); st.text(selected_row['メモ'])
            if selected_row['写真URL']:
                st.markdown(f"**写真:** [ファイルを開く]({selected_row['写真URL']})", unsafe_allow_html=True)

    elif note_display_type == "メンテノート":
        df = get_sheet_as_df(gc, SPREADSHEET_NAME, 'メンテノート_データ')
        required_cols = ['タイムスタンプ', 'ノート種別', 'メモ', '写真ファイル名', '写真URL']
        if not df.empty and not all(col in df.columns for col in required_cols): st.warning(f"メンテノートシートのヘッダー形式が正しくありません。"); return
        if df.empty: st.info("まだメンテノートは登録されていません。"); return
        
        search_term = st.text_input("メモの内容で検索")
        filtered_df = df.sort_values(by='タイムスタンプ', ascending=False)
        if search_term:
            filtered_df = filtered_df[filtered_df['メモ'].str.contains(search_term, case=False, na=False)]
        
        if filtered_df.empty: st.info(f"検索条件に一致するノートはありません。"); return

        options_indices = ["---"] + filtered_df.index.tolist()
        selected_index = st.selectbox(
            "ノートを選択",
            options=options_indices,
            format_func=lambda idx: "---" if idx == "---" else f"{filtered_df.loc[idx, 'メモ'][:40]}" + ("..." if len(filtered_df.loc[idx, 'メモ']) > 40 else "")
        )

        if selected_index != "---":
            selected_row = filtered_df.loc[selected_index]
            st.subheader(f"詳細: {selected_row['タイムスタンプ']}")
            st.write(f"**メモ:**"); st.text(selected_row['メモ'])
            if selected_row['写真URL']:
                st.markdown(f"**写真:** [ファイルを開く]({selected_row['写真URL']})", unsafe_allow_html=True)

def page_calendar(calendar_service):
    st.header("📅 Googleカレンダーの管理")
    tab1, tab2 = st.tabs(["予定の確認", "新しい予定の追加"])
    with tab1:
        st.subheader("期間を指定して予定を表示"); calendar_url = f"https://calendar.google.com/calendar/u/0/r?cid={DEFAULT_CALENDAR_ID}"; st.markdown(f"**[Googleカレンダーで直接開く]({calendar_url})**", unsafe_allow_html=True)
        col1, col2 = st.columns(2); start_date = col1.date_input("開始日", datetime.today()); end_date = col2.date_input("終了日", datetime.today() + timedelta(days=7))
        if start_date > end_date: st.error("終了日は開始日以降に設定してください。")
        else:
            try:
                timeMin = datetime.combine(start_date, time.min).isoformat() + 'Z'; timeMax = datetime.combine(end_date, time.max).isoformat() + 'Z'
                events_result = calendar_service.events().list(calendarId=DEFAULT_CALENDAR_ID, timeMin=timeMin, timeMax=timeMax, singleEvents=True, orderBy='startTime').execute()
                events = events_result.get('items', [])
                if not events: st.info("指定された期間に予定はありません。")
                else:
                    event_data = []
                    for event in events:
                        start = event['start'].get('dateTime', event['start'].get('date'))
                        if 'T' in start: dt = datetime.fromisoformat(start); date_str, time_str = dt.strftime("%Y/%m/%d (%a)"), dt.strftime("%H:%M")
                        else: date_str, time_str = datetime.strptime(start, "%Y-%m-%d").strftime("%Y/%m/%d (%a)"), "終日"
                        event_data.append({"日付": date_str, "時刻": time_str, "件名": event['summary'], "場所": event.get('location', '')})
                    st.dataframe(pd.DataFrame(event_data), use_container_width=True)
            except HttpError as e: st.error(f"カレンダーの読み込みに失敗しました: {e}")
    with tab2:
        st.subheader("新しい予定を追加")
        with st.form("add_event_form", clear_on_submit=False):
            event_summary = st.text_input("件名 *")
            col1, col2 = st.columns(2); event_date = col1.date_input("日付 *", datetime.today()); is_allday = col2.checkbox("終日", value=False)
            if not is_allday:
                col3, col4 = st.columns(2); start_time, end_time = col3.time_input("開始時刻 *", time(9, 0)), col4.time_input("終了時刻 *", time(10, 0))
            event_location = st.text_input("場所"); event_description = st.text_area("説明")
            submitted = st.form_submit_button("カレンダーに追加")
            if submitted:
                if not event_summary: st.error("件名は必須です。")
                else:
                    if is_allday: start, end = {'date': event_date.isoformat()}, {'date': (event_date + timedelta(days=1)).isoformat()}
                    else:
                        tz = "Asia/Tokyo"; start = {'dateTime': datetime.combine(event_date, start_time).isoformat(), 'timeZone': tz}; end = {'dateTime': datetime.combine(event_date, end_time).isoformat(), 'timeZone': tz}
                    event_body = {'summary': event_summary, 'location': event_location, 'description': event_description, 'start': start, 'end': end}
                    try:
                        created_event = calendar_service.events().insert(calendarId=DEFAULT_CALENDAR_ID, body=event_body).execute()
                        st.success(f"予定「{created_event.get('summary')}」を追加しました。"); st.markdown(f"[カレンダーで確認]({created_event.get('htmlLink')})")
                    except HttpError as e: st.error(f"予定の追加に失敗しました: {e}")

def page_minutes(drive_service, gc):
    st.header("🎙️ 会議の議事録の管理"); minutes_sheet_name = '議事録_データ'
    tab1, tab2 = st.tabs(["議事録の確認", "新しい議事録の登録"])
    with tab1:
        st.subheader("登録済みの議事録")
        df = get_sheet_as_df(gc, SPREADSHEET_NAME, minutes_sheet_name); required_cols = ['タイムスタンプ', '会議タイトル', '音声ファイルURL', '音声ファイル名', '議事録内容']
        if not df.empty and not all(col in df.columns for col in required_cols): st.warning(f"議事録シートのヘッダー形式が正しくありません。"); return
        if df.empty: st.info("まだ議事録は登録されていません。"); return
        options = {f"{row['タイムスタンプ']} - {row['会議タイトル']}": idx for idx, row in df.iterrows()}
        selected_key = st.selectbox("議事録を選択", ["---"] + list(options.keys()))
        if selected_key != "---":
            selected_row = df.loc[options[selected_key]]
            st.subheader(selected_row['会議タイトル']); st.caption(f"登録日時: {selected_row['タイムスタンプ']}")
            if selected_row['音声ファイルURL']: st.markdown(f"**[音声ファイルを開く]({selected_row['音声ファイルURL']})** ({selected_row['音声ファイル名']})")
            st.markdown("---"); st.markdown(selected_row['議事録内容'])
    with tab2:
        st.subheader("新しい議事録を登録")
        with st.form("minutes_form", clear_on_submit=True):
            title = st.text_input("会議のタイトル *"); audio_file = st.file_uploader("音声ファイル (任意)", type=["mp3", "wav", "m4a", "flac"]); content = st.text_area("議事録内容", height=300)
            submitted = st.form_submit_button("議事録を保存")
            if submitted:
                if not title: st.error("タイトルは必須です。")
                else:
                    filename, url = upload_file_to_drive(drive_service, audio_file, FOLDER_IDS['MINUTES'], title)
                    row_data = [datetime.now().strftime("%Y%m%d_%H%M%S"), title, filename, url, content]
                    spreadsheet = gc.open(SPREADSHEET_NAME); spreadsheet.worksheet(minutes_sheet_name).append_row(row_data)
                    st.success("議事録を保存しました。"); st.cache_data.clear(); st.rerun()

def page_qa(drive_service, gc):
    st.header("💡 山根研 知恵袋"); qa_sheet_name = '知恵袋_データ'; answers_sheet_name = '知恵袋_解答'
    tab1, tab2 = st.tabs(["質問と回答を見る", "新しい質問を投稿する"])
    with tab1:
        df_qa = get_sheet_as_df(gc, SPREADSHEET_NAME, qa_sheet_name); df_answers = get_sheet_as_df(gc, SPREADSHEET_NAME, answers_sheet_name)
        required_qa_cols = ['タイムスタンプ', '質問タイトル', '質問内容', '連絡先メールアドレス', '添付ファイル名', '添付ファイルURL', 'ステータス']
        if not df_qa.empty and not all(col in df_qa.columns for col in required_qa_cols):
            st.error(f"`{qa_sheet_name}`シートのヘッダーが正しくありません。"); st.warning(f"シートの1行目が、以下のようになっているか確認してください。"); st.code(required_qa_cols, language='python'); return
        if df_qa.empty:
            st.info("まだ質問はありません。「新しい質問を投稿する」タブから投稿してください。"); return
        df_qa['タイムスタンプ_dt'] = pd.to_datetime(df_qa['タイムスタンプ'], format="%Y%m%d_%H%M%S"); df_qa = df_qa.sort_values(by='タイムスタンプ_dt', ascending=False)
        options = {f"[{row['ステータス']}] {row['質問タイトル']} ({row['タイムスタンプ']})": row['タイムスタンプ'] for _, row in df_qa.iterrows()}
        selected_ts = st.selectbox("質問を選択", ["---"] + list(options.keys()))
        if selected_ts != "---":
            question_id = options[selected_ts]; question = df_qa[df_qa['タイムスタンプ'] == question_id].iloc[0]
            with st.container(border=True):
                st.subheader(f"Q: {question['質問タイトル']}"); st.caption(f"投稿日時: {question['タイムスタンプ']} | ステータス: {question['ステータス']}"); st.markdown(question['質問内容'])
                if question['添付ファイルURL']:
                    st.markdown(f"**添付ファイル:** [リンクを開く]({question['添付ファイルURL']})", unsafe_allow_html=True)
                if question['ステータス'] == '未解決' and st.button("この質問を解決済みにする", key=f"resolve_{question_id}"):
                    try:
                        spreadsheet = gc.open(SPREADSHEET_NAME); qa_sheet_obj = spreadsheet.worksheet(qa_sheet_name)
                        cell = qa_sheet_obj.find(question_id)
                        status_col = qa_sheet_obj.get_all_values()[0].index("ステータス") + 1
                        qa_sheet_obj.update_cell(cell.row, status_col, "解決済み")
                        st.success("ステータスを「解決済み」に更新しました。"); st.cache_data.clear(); st.rerun()
                    except Exception as e: st.error(f"更新に失敗しました: {e}")
            st.markdown("---"); st.subheader("Answers")
            answers = df_answers[df_answers['質問タイムスタンプ (質問ID)'] == question_id] if not df_answers.empty else pd.DataFrame()
            if answers.empty: st.info("この質問にはまだ回答がありません。")
            else:
                for _, answer in answers.iterrows():
                    with st.container(border=True):
                        st.markdown(f"**A:** {answer['解答内容']}"); st.caption(f"回答者: {answer['解答者 (任意)'] or '匿名'} | 日時: {answer['タイムスタンプ']}")
                        if answer['添付ファイルURL']:
                            st.markdown(f"**添付ファイル:** [リンクを開く]({answer['添付ファイルURL']})", unsafe_allow_html=True)
            with st.expander("この質問に回答する"):
                with st.form(f"answer_form_{question_id}", clear_on_submit=True):
                    answer_content = st.text_area("回答内容 *"); answerer_name = st.text_input("回答者名（任意）"); answer_file = st.file_uploader("参考ファイル（任意）")
                    submitted = st.form_submit_button("回答を投稿する")
                    if submitted:
                        if not answer_content: st.warning("回答内容を入力してください。")
                        else:
                            filename, url = upload_file_to_drive(drive_service, answer_file, FOLDER_IDS['QA'], question['質問タイトル'])
                            row_data = [datetime.now().strftime("%Y%m%d_%H%M%S"), question['質問タイトル'], question_id, answer_content, answerer_name, "", filename, url]
                            spreadsheet = gc.open(SPREADSHEET_NAME); spreadsheet.worksheet(answers_sheet_name).append_row(row_data); st.success("回答を投稿しました！"); st.cache_data.clear(); st.rerun()
    with tab2:
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

def page_handover(drive_service, gc):
    st.header("🔑 引き継ぎ情報の管理"); handover_sheet_name = '引き継ぎ_データ'
    tab1, tab2 = st.tabs(["情報の確認", "新しい情報の登録"])
    with tab1:
        st.subheader("登録済みの引き継ぎ情報")
        df = get_sheet_as_df(gc, SPREADSHEET_NAME, handover_sheet_name); required_handover_cols = ['タイムスタンプ', '種類', 'タイトル', '内容1', '内容2', '内容3', 'メモ']
        if not df.empty and not all(col in df.columns for col in required_handover_cols):
            st.error(f"`{handover_sheet_name}`シートのヘッダーが正しくありません。"); st.warning(f"必要な列: {', '.join(required_handover_cols)}"); return
        if df.empty:
            st.info("まだ引き継ぎ情報はありません。"); return
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
            st.subheader(f"{selected_row['タイトル']} の詳細"); st.write(f"**種類:** {selected_row['種類']}")
            if selected_row['種類'] == "マニュアル":
                if selected_row['内容1']: st.markdown(f"**ファイル/URL:** [リンクを開く]({selected_row['内容1']})")
                st.write("**メモ:**"); st.text(selected_row['メモ'])
            elif selected_row['種類'] == "連絡先": st.write(f"**電話番号:** {selected_row['内容1']}"); st.write(f"**メール:** {selected_row['内容2']}"); st.write("**メモ:**"); st.text(selected_row['メモ'])
            elif selected_row['種類'] == "パスワード": st.write(f"**サービス名/場所:** {selected_row['タイトル']}"); st.write(f"**ユーザー名:** {selected_row['内容1']}"); st.write(f"**パスワード:** {selected_row['内容2']}"); st.write("**メモ:**"); st.text(selected_row['メモ'])
            elif selected_row['種類'] == "その他": st.write(f"**内容:**"); st.markdown(selected_row['内容1']); st.write("**メモ:**"); st.text(selected_row['メモ'])
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

def page_inquiry(gc):
    st.header("✉️ お問い合わせフォーム"); inquiry_sheet_name = 'お問い合わせ_データ'
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
                st.success("お問い合わせ内容を記録しました。ご協力ありがとうございます！"); st.info("以下のリンクをクリックして、Gmailで内容を送信してください。")
                st.markdown(f"**[Gmailを起動して管理者にメールを送信する]({gmail_link})**", unsafe_allow_html=True); st.cache_data.clear()

def page_pl_analysis():
    st.header("🔬 PLデータ解析")
    with st.expander("ステップ1：波長校正", expanded=True):
        st.write("2つの基準波長の反射光データをアップロードして、分光器の傾き（nm/pixel）を校正します。")
        col1, col2 = st.columns(2)
        with col1:
            cal1_wavelength = st.number_input("基準波長1 (nm)", value=1500)
            cal1_file = st.file_uploader(f"{cal1_wavelength}nm の校正ファイル (.txt)", type=['txt'], key="cal1")
        with col2:
            cal2_wavelength = st.number_input("基準波長2 (nm)", value=1570)
            cal2_file = st.file_uploader(f"{cal2_wavelength}nm の校正ファイル (.txt)", type=['txt'], key="cal2")
        if st.button("校正を実行", key="run_calibration"):
            if cal1_file and cal2_file:
                df1 = load_pl_data(cal1_file)
                df2 = load_pl_data(cal2_file)
                if df1 is not None and df2 is not None:
                    peak_pixel1 = df1.loc[df1['intensity'].idxmax()]['pixel']
                    peak_pixel2 = df2.loc[df2['intensity'].idxmax()]['pixel']
                    st.write("---"); st.subheader("校正結果")
                    col_res1, col_res2, col_res3 = st.columns(3)
                    col_res1.metric(f"{cal1_wavelength}nmのピーク位置", f"{int(peak_pixel1)} pixel")
                    col_res2.metric(f"{cal2_wavelength}nmのピーク位置", f"{int(peak_pixel2)} pixel")
                    try:
                        delta_wave = float(cal2_wavelength - cal1_wavelength)
                        delta_pixel = float(peak_pixel2 - peak_pixel1)
                        if delta_pixel == 0:
                            st.error("2つのピーク位置が同じです。異なる校正ファイルを選択するか、データを確認してください。")
                        else:
                            slope = delta_wave / delta_pixel
                            col_res3.metric("校正係数 (nm/pixel)", f"{slope:.4f}")
                            st.session_state['pl_calibrated'] = True
                            st.session_state['pl_slope'] = slope
                            st.success("校正係数を保存しました。ステップ2に進んでください。")
                    except Exception as e:
                        st.error(f"校正パラメータの計算中にエラーが発生しました: {e}")
            else:
                st.warning("両方の校正ファイルをアップロードしてください。")

    st.write("---")
    st.subheader("ステップ2：測定データ解析")
    if 'pl_calibrated' not in st.session_state or not st.session_state['pl_calibrated']:
        st.info("まず、ステップ1の波長校正を完了させてください。")
    else:
        st.success(f"波長校正済みです。（校正係数: {st.session_state['pl_slope']:.4f} nm/pixel）")
        with st.container(border=True):
            center_wavelength_input = st.number_input(
                "測定時の中心波長 (nm)", min_value=0, value=1650, step=10,
                help="この測定で装置に設定した中心波長を入力してください。凡例の自動整形にも使われます。"
            )
            uploaded_files = st.file_uploader("測定データファイル（複数選択可）をアップロード", type=['txt'], accept_multiple_files=True)
            if uploaded_files:
                st.subheader("解析結果")
                fig, ax = plt.subplots(figsize=(10, 6))
                all_dfs, filenames = [], []
                center_wl_str = str(int(center_wavelength_input))
                legend_labels = []
                for f in uploaded_files:
                    base_name = os.path.splitext(f.name)[0]
                    cleaned_label = base_name.replace(center_wl_str, "").strip(' _-')
                    legend_labels.append(cleaned_label if cleaned_label else base_name)
                for uploaded_file, label in zip(uploaded_files, legend_labels):
                    df = load_pl_data(uploaded_file)
                    if df is not None:
                        slope = st.session_state['pl_slope']
                        center_pixel = 256.5 
                        df['wavelength_nm'] = (df['pixel'] - center_pixel) * slope + center_wavelength_input
                        ax.plot(df['wavelength_nm'], df['intensity'], label=label, linewidth=2.5)
                        all_dfs.append(df); filenames.append(uploaded_file.name)
                if all_dfs:
                    ax.set_title(f"PLスペクトル (中心波長: {center_wavelength_input} nm)")
                    ax.set_xlabel("wavelength [nm]"); ax.set_ylabel("PL intensity")
                    ax.legend(loc='upper left', frameon=False, fontsize=14)
                    ax.grid(axis='y', linestyle='-', color='lightgray', zorder=0)
                    ax.tick_params(direction='in', top=True, right=True, which='both')
                    combined_df = pd.concat(all_dfs)
                    min_wl = combined_df['wavelength_nm'].min()
                    max_wl = combined_df['wavelength_nm'].max()
                    padding = (max_wl - min_wl) * 0.05
                    ax.set_xlim(min_wl - padding, max_wl + padding)
                    st.pyplot(fig)
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        for df, uploaded_file, label in zip(all_dfs, uploaded_files, legend_labels):
                            sheet_name = label[:31]
                            export_df = df[['wavelength_nm', 'intensity']].copy()
                            export_df.rename(columns={'intensity': os.path.splitext(uploaded_file.name)[0]}, inplace=True)
                            export_df.to_excel(writer, index=False, sheet_name=sheet_name)
                    processed_data = output.getvalue()
                    st.download_button(label="📈 Excelデータとしてダウンロード", data=processed_data, file_name=f"pl_analysis_{center_wavelength_input}nm.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# --- 5. メイン実行ブロック ---
def main():
    st.title("🛠️ 山根研 便利屋さん")
    gc, drive_service, calendar_service = initialize_google_services()
    st.sidebar.header("メニュー")
    menu_options = ["ノート記録", "ノート一覧", "PLデータ解析", "カレンダー", "議事録管理", "山根研知恵袋", "引き継ぎ情報", "お問い合わせフォーム"]
    selected_menu = st.sidebar.radio("機能を選択", menu_options)
    
    if selected_menu == "ノート記録":
        page_note_recording(drive_service, gc)
    elif selected_menu == "ノート一覧":
        page_note_list(gc)
    elif selected_menu == "PLデータ解析":
        page_pl_analysis()
    elif selected_menu == "カレンダー":
        page_calendar(calendar_service)
    elif selected_menu == "議事録管理":
        page_minutes(drive_service, gc)
    elif selected_menu == "山根研知恵袋":
        page_qa(drive_service, gc)
    elif selected_menu == "引き継ぎ情報":
        page_handover(drive_service, gc)
    elif selected_menu == "お問い合わせフォーム":
        page_inquiry(gc)

if __name__ == "__main__":
    main()
