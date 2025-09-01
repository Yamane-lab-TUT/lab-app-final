# --------------------------------------------------------------------------
# Yamane Lab Convenience Tool - Streamlit Application (v9.1 - Final)
#
# v9.1:
# -
# -
# -
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

# Google API client libraries
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from googleapiclient.errors import HttpError

# --- Global Configuration & Setup ---
st.set_page_config(page_title="山根研 便利屋さん", layout="wide")

# --- Initialize Google Services (Authentication Fix) ---
@st.cache_resource(show_spinner="Googleサービスに接続中...")
def initialize_google_services():
    try:
        scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/calendar']
        
        if "gcs_credentials" not in st.secrets:
            st.error("❌ 致命的なエラー: Streamlit CloudのSecretsに `gcs_credentials` が見つかりません。")
            st.stop()
        
        # SecretsからJSON "文字列" を取得
        creds_string = st.secrets["gcs_credentials"]

        # ★★★ 重要: コピー＆ペースト時に混入する可能性がある不正な空白文字(U+00A0)を自動的に削除 ★★★
        creds_string_cleaned = creds_string.replace('\u00A0', '')

        # 文字列を辞書(dictionary)に変換
        creds_dict = json.loads(creds_string_cleaned)
        
        # 辞書を使って各サービスを認証
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        gc = gspread.service_account_from_dict(creds_dict)
        drive_service = build('drive', 'v3', credentials=creds)
        calendar_service = build('calendar', 'v3', credentials=creds)
        
        return gc, drive_service, calendar_service

    except json.JSONDecodeError:
        st.error("❌ 致命的なエラー: SecretsのJSON文字列のフォーマットが正しくありません。")
        st.error("Secretsの内容を再度確認してください。特に、不要な文字が混入していないかご確認ください。")
        st.stop()
    except Exception as e:
        st.error(f"❌ 致命的なエラー: サービスの初期化に失敗しました。")
        st.exception(e)
        st.stop()

gc, drive_service, calendar_service = initialize_google_services()

# --- Utility Functions ---
@st.cache_data(ttl=300, show_spinner="シート「{sheet_name}」を読み込み中...")
def get_sheet_as_df(_gc, spreadsheet_name, sheet_name):
    try:
        spreadsheet = _gc.open(spreadsheet_name)
        worksheet = spreadsheet.worksheet(sheet_name)
        data = worksheet.get_all_values()
        if not data: return pd.DataFrame()
        
        headers = data[0]
        # Handle cases where there is only a header row
        if len(data) == 1:
            return pd.DataFrame(columns=headers)
            
        df = pd.DataFrame(data[1:], columns=headers)
        return df
    except gspread.exceptions.WorksheetNotFound:
        st.error(f"スプレッドシート内にシート名「{sheet_name}」が見つかりません。"); return pd.DataFrame()
    except Exception as e:
        st.warning(f"シート「{sheet_name}」の読込中にエラーが発生しました。シートが空か、ヘッダーのみの可能性があります。"); return pd.DataFrame()

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

# --- UI Page Functions ---

# (お客様のv9.0コードのUI関数がここに入ります。内容は変更しません)
# 以下はv9.0コードのUI関数群をそのまま貼り付けたものです。

SPREADSHEET_NAME = 'エピノート'
FOLDER_IDS = {
    'EP_D1': '1KQEeEsHChqtrAIvP91ILnf6oS4fTVi1p', 'EP_D2': '1inmARuM_SgiYHi4PR7rcWRH0jERKZVJy',
    'MT': '1YllkIwYuV3IqY4_i0YoyY43SAB-U8-0i', 'MINUTES': '1g7qiEFuEchsFFBKFJwxN2D2PjShuDtzM',
    'HANDOVER': '1Mr70YjsgCzMboD7UZStm7bE8LQs1mwFu', 'QA': '1cil7cMFmQlgfzqOD-8QOm4KqVB4Emy79'
}
DEFAULT_CALENDAR_ID = 'yamane.lab.6747@gmail.com'
INQUIRY_RECIPIENT_EMAIL = 'kyuno.yamato.ns@tut.ac.jp'

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
                        spreadsheet = gc.open(SPREADSHEET_NAME)
                        spreadsheet.worksheet('エピノート_データ').append_row(row_data)
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
                    spreadsheet = gc.open(SPREADSHEET_NAME)
                    spreadsheet.worksheet('メンテノート_データ').append_row(row_data)
                    st.success("メンテノートを保存しました！"); st.cache_data.clear(); st.rerun()

def page_note_list():
    st.header("📓 登録済みのノート一覧")
    note_display_type = st.radio("表示するノート", ("エピノート", "メンテノート"), horizontal=True, key="note_display_type")
    
    if note_display_type == "エピノート":
        df_ep = get_sheet_as_df(gc, SPREADSHEET_NAME, 'エピノート_データ')
        if df_ep.empty:
            st.info("まだエピノートは登録されていません。"); return
        
        ep_category_filter = st.selectbox("カテゴリで絞り込み", ["すべて"] + list(df_ep['カテゴリ'].unique()))
        
        filtered_df = df_ep.sort_values(by='タイムスタンプ', ascending=False)
        if ep_category_filter != "すべて":
            filtered_df = filtered_df[filtered_df['カテゴリ'] == ep_category_filter]
        
        if filtered_df.empty:
            st.info(f"検索条件に一致するノートはありません。"); return

        options_indices = ["---"] + filtered_df.index.tolist()
        selected_index = st.selectbox(
            "ノートを選択", options=options_indices,
            format_func=lambda idx: "---" if idx == "---" else f"{filtered_df.loc[idx, 'メモ'][:40]}" + ("..." if len(filtered_df.loc[idx, 'メモ']) > 40 else "")
        )
        
        if selected_index != "---":
            selected_row = filtered_df.loc[selected_index]
            st.subheader(f"詳細: {selected_row['タイムスタンプ']}")
            st.write(f"**カテゴリ:** {selected_row['カテゴリ']}")
            st.write(f"**メモ:**"); st.text(selected_row['メモ'])
            if selected_row['写真URL']:
                st.markdown(f"**写真:** [ファイルを開く]({selected_row['写真URL']})", unsafe_allow_html=True)

    elif note_display_type == "メンテノート":
        df_mt = get_sheet_as_df(gc, SPREADSHEET_NAME, 'メンテノート_データ')
        if df_mt.empty:
            st.info("まだメンテノートは登録されていません。"); return
        
        filtered_df = df_mt.sort_values(by='タイムスタンプ', ascending=False)
        
        options_indices = ["---"] + filtered_df.index.tolist()
        selected_index = st.selectbox(
            "ノートを選択", options=options_indices,
            format_func=lambda idx: "---" if idx == "---" else f"{filtered_df.loc[idx, 'メモ'][:40]}" + ("..." if len(filtered_df.loc[idx, 'メモ']) > 40 else "")
        )

        if selected_index != "---":
            selected_row = filtered_df.loc[selected_index]
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
            else:
                try:
                    timeMin = datetime.combine(start_date, time.min).isoformat() + 'Z'
                    timeMax = datetime.combine(end_date, time.max).isoformat() + 'Z'
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

# ... (他のpage関数も同様にここに続く) ...
def page_minutes():
    st.header("🎙️ 会議の議事録の管理"); minutes_sheet_name = '議事録_データ'
    tab1, tab2 = st.tabs(["議事録の確認", "新しい議事録の登録"])
    with tab1:
        df = get_sheet_as_df(gc, SPREADSHEET_NAME, minutes_sheet_name)
        if df.empty:
            st.info("まだ議事録は登録されていません。"); return
        options = {f"{row['タイムスタンプ']} - {row['会議タイトル']}": idx for idx, row in df.iterrows()}
        selected_key = st.selectbox("議事録を選択", ["---"] + list(options.keys()))
        if selected_key != "---":
            selected_row = df.loc[options[selected_key]]
            st.subheader(selected_row['会議タイトル']); st.caption(f"登録日時: {selected_row['タイムスタンプ']}")
            if selected_row.get('音声ファイルURL'): st.markdown(f"**[音声ファイルを開く]({selected_row['音声ファイルURL']})** ({selected_row.get('音声ファイル名', '')})")
            st.markdown("---"); st.markdown(selected_row['議事録内容'])
    with tab2:
        with st.form("minutes_form", clear_on_submit=True):
            title = st.text_input("会議のタイトル *"); audio_file = st.file_uploader("音声ファイル (任意)", type=["mp3", "wav", "m4a"]); content = st.text_area("議事録内容", height=300)
            submitted = st.form_submit_button("議事録を保存")
            if submitted:
                if not title: st.error("タイトルは必須です。")
                else:
                    filename, url = upload_file_to_drive(drive_service, audio_file, FOLDER_IDS['MINUTES'], title)
                    row_data = [datetime.now().strftime("%Y%m%d_%H%M%S"), title, filename, url, content]
                    gc.open(SPREADSHEET_NAME).worksheet(minutes_sheet_name).append_row(row_data)
                    st.success("議事録を保存しました。"); st.cache_data.clear(); st.rerun()

def page_qa():
    st.header("💡 山根研 知恵袋"); qa_sheet_name, answers_sheet_name = '知恵袋_データ', '知恵袋_解答'
    
    qa_status_filter = st.selectbox("表示する質問のステータス", ["すべての質問", "未解決のみ", "解決済みのみ"])

    df_qa = get_sheet_as_df(gc, SPREADSHEET_NAME, qa_sheet_name)
    if df_qa.empty:
        st.info("まだ質問はありません。"); 
    else:
        df_qa['タイムスタンプ_dt'] = pd.to_datetime(df_qa['タイムスタンプ'], format="%Y%m%d_%H%M%S")
        df_qa = df_qa.sort_values(by='タイムスタンプ_dt', ascending=False)
        
        filtered_df_qa = df_qa
        if qa_status_filter == "未解決のみ": filtered_df_qa = df_qa[df_qa['ステータス'] == '未解決']
        elif qa_status_filter == "解決済みのみ": filtered_df_qa = df_qa[df_qa['ステータス'] == '解決済み']
        
        if filtered_df_qa.empty:
            st.info("条件に一致する質問はありません。")
        else:
            options = {f"[{row['ステータス']}] {row['質問タイトル']}": row['タイムスタンプ'] for _, row in filtered_df_qa.iterrows()}
            selected_ts_key = st.selectbox("質問を選択", ["---"] + list(options.keys()))

            if selected_ts_key != "---":
                question_id = options[selected_ts_key]
                question = df_qa[df_qa['タイムスタンプ'] == question_id].iloc[0]
                with st.container(border=True):
                    st.subheader(f"Q: {question['質問タイトル']}")
                    st.caption(f"投稿日時: {question['タイムスタンプ']} | ステータス: {question['ステータス']}")
                    st.markdown(question['質問内容'])
                    if question['添付ファイルURL']: st.markdown(f"**添付ファイル:** [リンクを開く]({question['添付ファイルURL']})")
                    if question['ステータス'] == '未解決' and st.button("解決済みにする", key=f"resolve_{question_id}"):
                        cell = gc.open(SPREADSHEET_NAME).worksheet(qa_sheet_name).find(question_id)
                        gc.open(SPREADSHEET_NAME).worksheet(qa_sheet_name).update_cell(cell.row, 7, "解決済み")
                        st.success("ステータスを更新しました。"); st.cache_data.clear(); st.rerun()

                st.subheader("回答")
                df_answers = get_sheet_as_df(gc, SPREADSHEET_NAME, answers_sheet_name)
                answers = df_answers[df_answers['質問タイムスタンプ (質問ID)'] == question_id] if not df_answers.empty else pd.DataFrame()
                if answers.empty: st.info("まだ回答はありません。")
                else:
                    for _, answer in answers.iterrows():
                        with st.container(border=True):
                            st.markdown(f"**A:** {answer['解答内容']}")
                            st.caption(f"回答者: {answer.get('解答者 (任意)') or '匿名'} | 日時: {answer['タイムスタンプ']}")
                
                with st.expander("回答を投稿する"):
                    with st.form(f"answer_form_{question_id}", clear_on_submit=True):
                        answer_content = st.text_area("回答内容 *"); answerer_name = st.text_input("回答者名（任意）")
                        submitted = st.form_submit_button("回答を投稿")
                        if submitted and answer_content:
                            row_data = [datetime.now().strftime("%Y%m%d_%H%M%S"), question['質問タイトル'], question_id, answer_content, answerer_name, "", "", ""]
                            gc.open(SPREADSHEET_NAME).worksheet(answers_sheet_name).append_row(row_data)
                            st.success("回答を投稿しました。"); st.cache_data.clear(); st.rerun()

    with st.expander("新しい質問を投稿する", expanded=False):
        with st.form("new_question_form", clear_on_submit=True):
            q_title = st.text_input("質問タイトル *"); q_content = st.text_area("質問内容 *", height=150)
            q_file = st.file_uploader("参考ファイル"); q_email = st.text_input("連絡先メールアドレス（任意）")
            if st.form_submit_button("質問を投稿"):
                if q_title and q_content:
                    fname, furl = upload_file_to_drive(drive_service, q_file, FOLDER_IDS['QA'], q_title)
                    row_data = [datetime.now().strftime("%Y%m%d_%H%M%S"), q_title, q_content, q_email, fname, furl, "未解決"]
                    gc.open(SPREADSHEET_NAME).worksheet(qa_sheet_name).append_row(row_data)
                    st.success("質問を投稿しました。"); st.cache_data.clear(); st.rerun()
                else: st.error("タイトルと内容は必須です。")
                
def page_handover():
    st.header("🔑 引き継ぎ情報の管理"); handover_sheet_name = '引き継ぎ_データ'
    tab1, tab2 = st.tabs(["情報の確認", "新しい情報の登録"])
    with tab1:
        df = get_sheet_as_df(gc, SPREADSHEET_NAME, handover_sheet_name)
        if df.empty:
            st.info("まだ引き継ぎ情報はありません。"); return
        
        selected_type = st.selectbox("情報の種類で絞り込み", ["すべて"] + df['種類'].unique().tolist())
        filtered_df = df if selected_type == "すべて" else df[df['種類'] == selected_type]
        
        if filtered_df.empty: st.info(f"検索条件に一致する情報はありません。"); return
        
        options = {f"[{row['種類']}] {row['タイトル']}": idx for idx, row in filtered_df.iterrows()}
        selected_key = st.selectbox("情報を選択", ["---"] + list(options.keys()))
        if selected_key != "---":
            row = filtered_df.loc[options[selected_key]]
            st.subheader(f"{row['タイトル']} の詳細"); st.write(f"**種類:** {row['種類']}")
            if row['種類'] == "パスワード":
                st.write(f"**ユーザー名:** {row['内容1']}"); st.write(f"**パスワード:** {row['内容2']}")
            else: # マニュアル, 連絡先, その他
                st.write(f"**内容1:** {row['内容1']}"); st.write(f"**内容2:** {row['内容2']}")
            st.write("**メモ:**"); st.text(row['メモ'])
            
    with tab2:
        with st.form("handover_form", clear_on_submit=True):
            handover_type = st.selectbox("情報の種類", ["マニュアル", "連絡先", "パスワード", "その他"])
            title = st.text_input("タイトル / サービス名 / 氏名 *")
            c1, c2, file = "", "", None
            if handover_type == "パスワード": c1, c2 = st.text_input("ユーザー名"), st.text_input("パスワード", type="password")
            else: c1, c2 = st.text_area("内容1"), st.text_area("内容2")
            memo = st.text_area("メモ")
            if st.form_submit_button("保存"):
                if title:
                    row_data = [datetime.now().strftime("%Y%m%d_%H%M%S"), handover_type, title, c1, c2, "", memo]
                    gc.open(SPREADSHEET_NAME).worksheet(handover_sheet_name).append_row(row_data)
                    st.success("情報を保存しました。"); st.cache_data.clear(); st.rerun()
                else: st.error("タイトルは必須です。")

def page_inquiry():
    st.header("✉️ お問い合わせフォーム")
    with st.form("inquiry_form", clear_on_submit=True):
        category = st.selectbox("お問い合わせの種類", ["バグ報告", "機能改善要望", "その他"])
        content = st.text_area("詳細内容 *", height=150); contact = st.text_input("連絡先（任意）")
        if st.form_submit_button("送信"):
            if content:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                row_data = [timestamp, category, content, contact]
                gc.open(SPREADSHEET_NAME).worksheet('お問い合わせ_データ').append_row(row_data)
                subject = f"【研究室便利屋さん】お問い合わせ: {category}"
                body = f"種類: {category}\n内容:\n{content}\n連絡先: {contact or 'なし'}"
                gmail_link = generate_gmail_link(INQUIRY_RECIPIENT_EMAIL, subject, body)
                st.success("お問い合わせを記録しました。"); st.markdown(f"**[Gmailで管理者に通知する]({gmail_link})**", unsafe_allow_html=True)
                st.cache_data.clear()
            else: st.error("詳細内容を入力してください。")

# --- Main App Logic ---
def main():
    st.title("🛠️ 山根研 便利屋さん")
    st.sidebar.header("メニュー")
    menu = ["ノート記録", "ノート一覧", "カレンダー", "議事録管理", "山根研知恵袋", "引き継ぎ情報", "お問い合わせフォーム"]
    selected_page = st.sidebar.radio("機能を選択", menu)

    page_map = {
        "ノート記録": page_note_recording,
        "ノート一覧": page_note_list,
        "カレンダー": page_calendar,
        "議事録管理": page_minutes,
        "山根研知恵袋": page_qa,
        "引き継ぎ情報": page_handover,
        "お問い合わせフォーム": page_inquiry
    }
    page_map[selected_page]()

if __name__ == "__main__":
    main()
