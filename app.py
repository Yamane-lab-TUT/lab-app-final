# app.py

import streamlit as st
import gspread
import pandas as pd
import os
import io
import re
import json
import matplotlib.pyplot as plt
import numpy as np
from datetime import datetime, time, timedelta
from urllib.parse import quote as url_quote
from io import BytesIO

# Google API client libraries (設定ファイルが存在する場合のみ利用)
# from google.oauth2.service_account import Credentials 
# from googleapiclient.discovery import build
# from google.cloud import storage
# from google.auth.exceptions import DefaultCredentialsError
# from google.api_core import exceptions


# --- Global Configuration & Setup ---
st.set_page_config(page_title="山根研 便利屋さん", layout="wide")

# ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
# ↓↓↓↓↓↓ 【重要】ご自身の「バケット名」に書き換えてください ↓↓↓↓↓↓
CLOUD_STORAGE_BUCKET_NAME = "your-gcs-bucket-name" # Placeholder for Cloud Storage
# ↑↑↑↑↑↑ 【重要】ご自身の「バケット名」に書き換えてください ↑↑↑↑↑↑
# ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★


# --------------------------------------------------------------------------
# --- Data Loading and Caching ---
# 処理落ち対策: Streamlitのキャッシュ機能でデータ読み込みを高速化
@st.cache_data(show_spinner="データを読み込み中...")
def load_iv_data(uploaded_file):
    """アップロードされたIVデータファイル（TXT/CSV）をロバストに読み込む関数。"""
    
    file_name = uploaded_file.name
    
    # ファイルをバイナリとして読み込み、文字列にデコード（UTF-8, Shift-JISを試行）
    try:
        data_string = uploaded_file.getvalue().decode('utf-8')
    except UnicodeDecodeError:
        try:
            data_string = uploaded_file.getvalue().decode('shift_jis')
        except:
            # 処理できない場合はエラーをログに記録してNoneを返す
            return None, file_name

    try:
        # StringIOを使ってpd.read_csvに渡す
        data_io = io.StringIO(data_string)
        
        # ファイル形式の共通項として、最初のヘッダー行（VF(V) IF(A)）をスキップし、
        # タブ/スペース区切りで読み込む
        # header=Noneで読み込み、後で名前を付ける
        df = pd.read_csv(data_io, sep=r'\s+', skiprows=1, header=None, names=['VF(V)', 'IF(A)'])
        
        # 稀にヘッダーが2行目以降にある場合も考慮し、数値でない行をドロップ
        df['VF(V)'] = pd.to_numeric(df['VF(V)'], errors='coerce')
        df['IF(A)'] = pd.to_numeric(df['IF(A)'], errors='coerce')
        df.dropna(inplace=True)

        return df, file_name

    except Exception as e:
        st.error(f"ファイル '{file_name}' の処理中にエラーが発生しました。形式を確認してください。")
        # st.exception(e) # デバッグ用
        return None, file_name

# --------------------------------------------------------------------------
# --- Page Functions (実装済み: IVデータ解析) ---
# --------------------------------------------------------------------------

def page_iv_analysis():
    st.header("⚡ IV Data Analysis (IVデータ解析)")
    st.markdown("複数のIVデータファイルを選択し、グラフ描画と比較、統合データのエクスポートを行います。")

    uploaded_files = st.file_uploader(
        "IVデータファイル（.txt または .csv）を選択してください",
        type=['txt', 'csv'],
        accept_multiple_files=True
    )

    if uploaded_files:
        st.subheader("📊 IV Characteristic Plot")
        
        # 処理落ち対策: グラフ描画のための設定 (fig, ax)を事前に作成
        fig, ax = plt.subplots(figsize=(12, 7))
        
        all_data_frames = []
        
        # 全ファイルをループして処理
        for uploaded_file in uploaded_files:
            # キャッシュされた関数を使ってデータをロード
            df, file_name = load_iv_data(uploaded_file)
            
            if df is not None and not df.empty:
                voltage_col = 'VF(V)'
                current_col = 'IF(A)'
                
                # データ結合とエクスポートのために列名をリネームし、電圧と電流の2列のみ抽出
                df_to_merge = df[[voltage_col]].copy()
                df_to_merge[f'Current_A_{file_name}'] = df[current_col]

                all_data_frames.append(df_to_merge)
                
                # グラフにプロット
                ax.plot(df[voltage_col], df[current_col], label=file_name)

        
        # グラフ設定 (文字化け対策: すべて英語)
        ax.set_title('IV Characteristic Plot', fontsize=16)
        ax.set_xlabel('Voltage (V)', fontsize=14)
        ax.set_ylabel('Current (A)', fontsize=14)
        ax.grid(True, linestyle='--', alpha=0.6)
        ax.legend(title='File Name', loc='best')
        ax.ticklabel_format(style='sci', axis='y', scilimits=(0, 0)) # Y軸を科学表記に

        # Streamlitにグラフを表示
        st.pyplot(fig, use_container_width=True)
        
        # 処理落ち対策: Matplotlibのメモリを解放
        plt.close(fig)

        # --- データ結合とExcelエクスポート ---
        if all_data_frames:
            st.subheader("📝 統合データのエクスポート")
            
            # 最初のデータフレームを基準に結合を開始
            combined_df = all_data_frames[0].rename(columns={all_data_frames[0].columns[0]: 'Voltage_V'})
            
            # 2つ目以降のデータフレームを 'Voltage_V' をキーに結合
            for i in range(1, len(all_data_frames)):
                df_current = all_data_frames[i].rename(columns={all_data_frames[i].columns[0]: 'Voltage_V'})
                combined_df = pd.merge(combined_df, df_current, on='Voltage_V', how='outer')
            
            # 電圧順にソート
            combined_df.sort_values(by='Voltage_V', inplace=True)
            
            st.dataframe(combined_df.head())
            
            # Excel書き出しのためのBytesIOを使用 (xlsxwriterエラー対策)
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                combined_df.to_excel(writer, sheet_name='Combined_IV_Data', index=False)
            
            processed_data = output.getvalue()
            
            st.download_button(
                label="📈 結合Excelデータとしてダウンロード",
                data=processed_data,
                file_name=f"iv_analysis_combined_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("有効なデータファイルが見つかりませんでした。")


# --------------------------------------------------------------------------
# --- Page Functions (未実装/プレースホルダー) ---
# --------------------------------------------------------------------------

def page_note_recording():
    st.header("📝 エピノート記録")
    st.info("この機能は現在構築中です。")

def page_note_list():
    st.header("📚 エピノート一覧")
    st.info("この機能は現在構築中です。")

def page_calendar():
    st.header("🗓️ スケジュール・装置予約")
    st.info("この機能は現在構築中です。")

def page_pl_analysis():
    st.header("🔬 PLデータ解析")
    st.info("この機能は現在構築中です。")

def page_minutes():
    st.header("議事録・ミーティングメモ")
    st.info("この機能は現在構築中です。")

def page_qa_forum():
    st.header("💡 知恵袋・質問箱")
    st.info("この機能は現在構築中です。")
    
def page_handoff_notes():
    st.header("🤝 装置引き継ぎメモ")
    st.info("この機能は現在構築中です。")

def page_trouble_report():
    st.header("🚨 トラブル報告")
    st.info("この機能は現在構築中です。")

def page_contact():
    st.header("✉️ 連絡・問い合わせ")
    st.info("この機能は現在構築中です。")


# --------------------------------------------------------------------------
# --- Main App Execution ---
# --------------------------------------------------------------------------
def main():
    st.sidebar.title("山根研 ツールキット")
    
    # アプリ内の日本語表示は維持
    menu_selection = st.sidebar.radio("機能選択", [
        "📝 エピノート記録", "📚 エピノート一覧", "🗓️ スケジュール・装置予約", 
        "⚡ IVデータ解析", "🔬 PLデータ解析",
        "議事録・ミーティングメモ", "💡 知恵袋・質問箱", "🤝 装置引き継ぎメモ", 
        "🚨 トラブル報告", "✉️ 連絡・問い合わせ"
    ])
    
    if menu_selection == "📝 エピノート記録": 
        page_note_recording()
    elif menu_selection == "📚 エピノート一覧": 
        page_note_list()
    elif menu_selection == "🗓️ スケジュール・装置予約": 
        page_calendar()
    elif menu_selection == "⚡ IVデータ解析": 
        page_iv_analysis()
    elif menu_selection == "🔬 PLデータ解析": 
        page_pl_analysis()
    elif menu_selection == "議事録・ミーティングメモ": 
        page_minutes()
    elif menu_selection == "💡 知恵袋・質問箱": 
        page_qa_forum()
    elif menu_selection == "🤝 装置引き継ぎメモ": 
        page_handoff_notes()
    elif menu_selection == "🚨 トラブル報告": 
        page_trouble_report()
    elif menu_selection == "✉️ 連絡・問い合わせ": 
        page_contact()

if __name__ == "__main__":
    main()
    
