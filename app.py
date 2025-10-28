# app.py (page_iv_analysis 関数内)

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
        
        fig, ax = plt.subplots(figsize=(12, 7))
        
        all_data_for_export = [] # 各ファイルのDFとファイル名を格納
        
        # 1. データの読み込みとグラフ描画
        for uploaded_file in uploaded_files:
            # キャッシュされた関数を使ってデータをロード
            df, file_name = load_iv_data(uploaded_file)
            
            if df is not None and not df.empty:
                voltage_col = 'VF(V)'
                current_col = 'IF(A)'
                
                # グラフにプロット
                ax.plot(df[voltage_col], df[current_col], label=file_name)
                
                # エクスポート用に[Voltage_V, Current_A_filename]のDFをリストに追加
                df_export = df.rename(columns={voltage_col: 'Voltage_V', current_col: f'Current_A_{file_name}'})
                all_data_for_export.append({'name': file_name, 'df': df_export})

        
        # グラフ設定 (文字化け対策: すべて英語)
        ax.set_title('IV Characteristic Plot', fontsize=16)
        ax.set_xlabel('Voltage (V)', fontsize=14)
        ax.set_ylabel('Current (A)', fontsize=14)
        ax.grid(True, linestyle='--', alpha=0.6)
        ax.legend(title='File Name', loc='best')
        ax.ticklabel_format(style='sci', axis='y', scilimits=(0, 0))
        
        # Streamlitにグラフを表示
        st.pyplot(fig, use_container_width=True)
        # 処理落ち対策: Matplotlibのメモリを解放
        plt.close(fig)

        # ------------------------------------------------------------------
        # 2. データ結合とExcelエクスポート (メモリ負荷軽減版)
        # ------------------------------------------------------------------
        if all_data_for_export:
            st.subheader("📝 データのエクスポート")
            
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                
                # --- 各ファイルを別シートに出力 (メモリ負荷 小) ---
                for data_item in all_data_for_export:
                    file_name = data_item['name']
                    df_export = data_item['df']
                    
                    # Excelのシート名制限（31文字）に対応
                    sheet_name = file_name.replace('.txt', '').replace('.csv', '')
                    if len(sheet_name) > 31:
                         sheet_name = sheet_name[:28] + '...' 
                    
                    # 最初の2列 (Voltage_V, Current_A_filename) のみ書き出し
                    df_export.to_excel(writer, sheet_name=sheet_name, index=False)

                # --- 結合データも最終シートに出力 (メモリ負荷 高) ---
                st.info("💡 結合データを作成中です。ファイルが多い場合、数秒かかることがあります。")
                
                # 最初のデータフレームを基準に結合を開始
                combined_df = all_data_for_export[0]['df'][['Voltage_V']].copy()
                
                # 2つ目以降のデータフレームを 'Voltage_V' をキーに結合
                for item in all_data_for_export:
                    df_current = item['df']
                    combined_df = pd.merge(combined_df, df_current, on='Voltage_V', how='outer')
                
                # 電圧順にソート
                combined_df.sort_values(by='Voltage_V', inplace=True)
                
                # 結合DFのプレビュー
                st.dataframe(combined_df.head())
                
                # 結合DFを最終シートに出力
                combined_df.to_excel(writer, sheet_name='__COMBINED_DATA__', index=False)
                
                # 処理落ち対策: 結合DFのメモリを直後に解放
                del combined_df
                
            
            processed_data = output.getvalue()
            
            st.download_button(
                label="📈 結合/個別データを含むExcelファイルとしてダウンロード",
                data=processed_data,
                file_name=f"iv_analysis_combined_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            st.info("🎉 ダウンロードされるExcelファイルには、**各データが個別シート**として、また**全データが結合されたシート**として保存されます。")
        else:
            st.warning("有効なデータファイルが見つかりませんでした。")
