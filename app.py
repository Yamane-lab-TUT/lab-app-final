import streamlit as st
from google.oauth2.service_account import Credentials
import gspread

st.set_page_config(layout="wide")
st.title("🔬 Google 認証テスト")

try:
    st.info("ステップ1: Streamlit CloudのSecretsから認証情報を読み込みます...")
    
    # この`creds_info`は辞書(dictionary)として自動的に読み込まれます
    creds_info = st.secrets["gcs_credentials"]
    
    st.success("✅ Secretsの読み込みに成功しました。")
    st.info("ステップ2: 読み込んだ情報を使ってGoogleのサーバーに接続します...")
    
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_info(creds_info, scopes=scopes)
    gc = gspread.authorize(creds)
    
    st.success("🎉 **認証に成功しました！**")
    st.balloons()
    st.info("この画面が表示されれば、Secretsの設定は完璧です。元のapp.pyに戻してアプリを再起動してください。")

except Exception as e:
    st.error("❌ 認証に失敗しました。")
    st.warning("Secretsのフォーマットがまだ間違っている可能性があります。ステップ1, 2を再度ご確認ください。")
    st.exception(e)
