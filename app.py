import streamlit as st
import gspread
import pandas as pd
from google.oauth2.service_account import Credentials

# --- 1. Googleサービス初期化（お客様のコードから流用） ---
@st.cache_resource(show_spinner="Googleサービスに接続中...")
def initialize_google_services():
    try:
        scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
        
        if "gcs_credentials" not in st.secrets:
            st.error("❌ 致命的なエラー: Streamlit CloudのSecretsに [gcs_credentials] が見つかりません。")
            st.stop()
            
        creds_info = st.secrets["gcs_credentials"]
        creds = Credentials.from_service_account_info(creds_info, scopes=scopes)
        gc = gspread.service_account_from_dict(creds_info)
        return gc
    except Exception as e:
        st.error(f"❌ 致命的なエラー: サービスの初期化に失敗しました。")
        st.info("Secretsの設定が正しいTOML形式になっているか、再度ご確認ください。")
        st.exception(e)
        st.stop()

# --- 2. データ読み込み診断関数 ---
def debug_get_sheet_as_df(gc, spreadsheet_name, sheet_name):
    """診断用に、より詳細なエラー情報を表示する関数"""
    st.markdown(f"--- \n ### 診断中: シート「`{sheet_name}`」")
    try:
        spreadsheet = gc.open(spreadsheet_name)
        worksheet = spreadsheet.worksheet(sheet_name)
        data = worksheet.get_all_values()
        
        if not data:
            st.warning("⚠️ このシートは完全に空です。")
            return pd.DataFrame()
            
        st.success(f"✅ シートの全データ取得に成功。({len(data)}行)")

        headers = data[0]
        st.write(f"**ヘッダー行:** `{headers}`")

        df = pd.DataFrame(data[1:], columns=headers)
        st.success(f"✅ DataFrameへの変換に成功。")
        st.dataframe(df.head(3)) # 最初の3行を表示
        return df

    except gspread.exceptions.WorksheetNotFound:
        st.error(f"❌ **失敗**: スプレッドシート内にこの名前のシートが見つかりません。")
        return None
    except Exception as e:
        st.error(f"❌ **失敗**: このシートの読み込み中にエラーが発生しました。`Short substrate on input`はここで発生している可能性が高いです。")
        st.exception(e) # エラーの詳細を画面に表示
        return None

# --- 3. メインの診断処理 ---
st.title("🔬 Google Sheets 連携 診断ツール")

# Googleサービスに接続
gc = initialize_google_services()
st.success("✅ **ステップ1:** Googleサービスへの認証に成功しました。")

# スプレッドシートを開く
SPREADSHEET_NAME = 'エピノート'
st.header(f"**ステップ2:** スプレッドシート「{SPREADSHEET_NAME}」の各シートを診断します。")

# お客様から頂いたCSVファイル名からシート名をリスト化
# このリストにあるシートを順番にチェックします。
sheets_to_check = [
    "エピノート_データ",
    "メンテノート_データ",
    "議事録_データ",
    "引き継ぎ_データ",
    "知恵袋_データ",
    "知恵袋_解答",
    "お問い合わせ_データ"
]

# 各シートを診断
all_success = True
for sheet_name in sheets_to_check:
    df = debug_get_sheet_as_df(gc, SPREADSHEET_NAME, sheet_name)
    if df is None:
        all_success = False

st.markdown("---")
st.header("最終診断結果")
if all_success:
    st.success("🎉 **すべてのシートの読み込みに成功しました！**")
    st.info("データ読み込み自体には問題がないようです。元のコードの別の部分に問題がある可能性があります。")
else:
    st.error("💔 **いくつかのシートで読み込みエラーが検出されました。**")
    st.warning("上記で「❌ 失敗」と表示されているシートが、アプリ全体のエラーの原因です。特にエラー詳細が表示されている箇所をご確認ください。")
