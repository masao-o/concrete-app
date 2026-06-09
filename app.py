# --- 5. AI解析処理 ---
    with col2:
        st.markdown("### 📊 高精密診断レポート")
        if uploaded_files and 'execute_analysis' in locals() and execute_analysis:
            if not api_key:
                st.error("APIキーが設定されていません。StreamlitのSecrets設定を確認してください。")
            else:
                with st.spinner("🔍 熟練コンクリート診断士AI(Gemini)が過去の実績に基づき統合解析中..."):
                    try:
                        genai.configure(api_key=api_key)
                        # 環境依存のエラーを回避するため「-latest」を付与したモデル名に変更
                        model = genai.GenerativeModel('gemini-1.5-flash-latest')
                        
                        env_text = "、".join(env_location) if env_location else "指定なし"
                        wet_text = "、".join(wet_status) if wet_status else "指定なし"
                        photo_comments_text = "\n".join(photo_comments)
                        
                        # 手動入力があればそれを適用
                        dim_info = "手動指定なし（写真内にスケールが無ければ測定不可として質問してください）"
                        if manual_width > 0 or manual_length > 0:
                            dim_info = f"【診断士による実測確定値】ひび割れ幅: {manual_width} mm, 長さ: {manual_length} cm"

                        # 超強力なプロンプトの構築
                        prompt = f"""
あなたは最高峰の「コンクリート診断士」です。官公庁や大手コンサルタントに提出する公式な報告書を作成してください。
以下の環境条件、入力情報、および各写真（Photo No.1〜）とそのコメントを踏まえて、テンプレートを排した完全オーダーメイドの重厚な工学的推測文（400文字以上）を出力してください。

【環境・現場入力情報】
- 構造物: {struct_type}
- 環境・湿潤: {env_text} / {wet_text}
- セメント種類: {cement_type}
- 供用年数: {elapsed_years}
- 主たる劣化症状: {crack_type}
- 気象・地域特有の環境: {region_info}
- 人為的補足: {human_factors_text}
- 寸法情報: {dim_info}
- 写真ごとのコメント:
{photo_comments_text}

【絶対厳守命令（ハルシネーション・寸法捏造の完全禁止）】
1. 手動指定寸法が無く、かつ写真内に「クラックスケール」や明確な寸法基準が確認できない場合、絶対に寸法を捏造しないでください。必ず文章の冒頭で「写真から正確な寸法を測定するための基準が確認できないため、勝手に判断せず保留します。正確な測定のために実測値または縮尺基準の提供を求めます」と回答・逆質問してください。
2. テンプレート回答を避け、塩害、凍害、中性化、ASR、不同沈下、せん断応力などの支配的メカニズムを、不動態被膜、遊離石灰、膨張圧などの専門用語を用いて、対象環境に合わせた推測をしてください。

【補修工法および追跡調査の選定基準】
・ひび割れ幅が0.2mm未満（または軽度）の場合は「劣化度Ⅰ」とし表面含浸工法等の予防保全や経過観察とする。
・0.2mm以上1.0mm未満の場合は「注入工法（低圧エポキシ樹脂注入など）」を提案する。
・1.0mm以上の場合は「充填工法（ポリマーセメントモルタル充填など）」を提案する。
・さらに、コンクリート内部の劣化度を確認するため、シュミットハンマーによる反発硬度試験、コア採取による圧縮強度試験、ドリル削孔による中性化深さ試験などの詳細調査の必要性を必ずプロの視点で明記すること。

出力は、確認できた（または手動入力された）「確定ひび割れ幅: 〇〇 mm」を冒頭に示し、その後【劣化原因の詳細（気象条件の考慮含む）】、【写真ごとの個別見解】、【対策案および詳細調査の推奨】を長文で出力してください。JSONは不要です。
"""
                        request_contents = [prompt] + images
                        response = model.generate_content(request_contents)
                        full_result_text = response.text
                        
                        # アラートカラーの判定（手動入力優先、なければ0）
                        final_width = manual_width
                        if final_width == 0:
                            try:
                                if "確定ひび割れ幅:" in full_result_text:
                                    # 修正：split後の文字列操作を安全に修正
                                    part = full_result_text.split("確定ひび割れ幅:")[1].split("mm")[0].strip()
                                    final_width = float(part)
                            except Exception:
                                final_width = 0.0

                        if final_width >= 0.2:
                            color_code = "#EF4444"
                            status_title = f"🔴 【要精密補修】確定/推定ひび割れ幅: {final_width} mm"
                            alert_desc = "⚠️ 判定基準：0.2mm以上のひび割れのため、指針に基づく「注入工法」等の検討が必要です。"
                        elif final_width > 0:
                            color_code = "#EAB308"
                            status_title = f"🟡 【経過観察】確定/推定ひび割れ幅: {final_width} mm"
                            alert_desc = "💡 判定基準：0.2mm未満のひび割れのため、表面含浸や経過観察（劣化度Ⅰ）に該当します。"
                        else:
                            color_code = "#3B82F6"
                            status_title = "🔵 【寸法未確定・質問あり】"
                            alert_desc = "ℹ️ スケールが不明、または寸法が入力されていないため、実測値の確認が求められています。"
                        
                        st.markdown(f"<div class='status-card'><h3 style='color: {color_code} !important; margin:0; font-size:22px;'>{status_title}</h3><p style='color: #F1F5F9 !important; font-size: 14px; margin: 8px 0 0 0; font-weight: bold;'>{alert_desc}</p></div>", unsafe_allow_html=True)
                        st.markdown("<h4 style='color: white; margin-top:20px;'>📑 AI Suite Pro 統合解析レポート</h4>", unsafe_allow_html=True)
                        st.info(full_result_text)

                        # --- 6. 複数枚対応 Excel出力（調査状況写真台帳フォーマット） ---
                        wb = openpyxl.Workbook()
                        ws = wb.active
                        ws.title = "調査状況写真台帳"
                        
                        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
                        ws.page_setup.paperSize = ws.PAPERSIZE_A4
                        
                        # 修正：インデックス指定を追加
                        ws.views.sheetView[0].showGridLines = False

                        font_header = Font(name="MS ゴシック", size=14, bold=True)
                        font_label = Font(name="MS ゴシック", size=11, bold=True)
                        font_data = Font(name="MS ゴシック", size=11)
                        
                        # 修正：style="thin" から border_style="thin" へ修正
                        thin_side = Side(border_style="thin")
                        border_cell = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

                        ws.column_dimensions['A'].width = 15
                        ws.column_dimensions['B'].width = 45
                        ws.column_dimensions['C'].width = 2
                        ws.column_dimensions['D'].width = 60
                        
                        p_name = project_name if project_name else "コンクリート構造物躯体調査"
                        l_name = location_name if location_name else "現場写真"

                        start_row = 1
                        for idx, img in enumerate(images):
                            ws.merge_cells(f"A{start_row}:D{start_row}")
                            ws[f"A{start_row}"] = f"{p_name} 状 況 写 真"
                            ws[f"A{start_row}"].font = font_header
                            ws.merge_cells(f"A{start_row+1}:D{start_row+1}")
                            ws[f"A{start_row+1}"] = f"施設名： {l_name}"
                            ws[f"A{start_row+1}"].font = font_header
                            ws[f"A{start_row+1}"].alignment = Alignment(horizontal="right")

                            info_labels = ["写真No.", "撮影箇所", "工 種", "位 置", "記 事（AI見解）"]
                            article_text = full_result_text if idx == 0 else photo_comments[idx]
                            info_values = [str(idx + 1), f"現場撮影写真 {idx+1}", "劣化状況調査", l_name, article_text]

                            for i, (label, value) in enumerate(zip(info_labels, info_values)):
                                r = start_row + 3 + i
                                ws[f"A{r}"] = label
                                ws[f"B{r}"] = value
                                ws[f"A{r}"].font = font_label
                                ws[f"B{r}"].font = font_data
                                ws[f"A{r}"].border = border_cell
                                ws[f"B{r}"].border = border_cell
                                ws[f"A{r}"].alignment = Alignment(horizontal="center", vertical="center")
                                ws[f"B{r}"].alignment = Alignment(wrap_text=True, vertical="top")

                            ws.row_dimensions[start_row + 7].height = 200

                            img_buffer = io.BytesIO()
                            img.save(img_buffer, format="PNG")
                            img_buffer.seek(0)
                            xl_img = ExcelImage(img_buffer)
                            xl_img.width, xl_img.height = 420, 310
                            ws.add_image(xl_img, f"D{start_row + 3}")

                            start_row += 35 

                        output = io.BytesIO()
                        wb.save(output)
                        
                        st.markdown("---")
                        st.download_button(
                            label="📥 官庁・役所・提出用 Excel写真台帳をダウンロード",
                            data=output.getvalue(),
                            file_name=f"【調査状況写真】{project_name if project_name else 'コンクリート調査'}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    except Exception as e:
                        st.error(f"解析中にエラーが発生しました: {e}")
