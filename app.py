import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import os

st.set_page_config(page_title="출고 생산성 분석 (디버깅 모드)", layout="centered")
st.title("🏭 출고 생산성 분석 프로그램 (에러 추적 모드)")

uploaded_files = st.file_uploader(
    "출고내역 엑셀 파일(xlsx)을 업로드해주세요.",
    type=["xlsx"],
    accept_multiple_files=True
)

target_seconds = st.number_input("🔄 작업시간 기준(초)", min_value=1, max_value=3600, value=60)

def generate_5_sheets(df_source, target_sec):
    bins = [0, 5, 10, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60, 90, 120, 150, 180, 360, 540, 720, np.inf]
    labels = ['0~5초', '5~10초', '10~15초', '15~20초', '20~25초', '25~30초', '30~35초', '35~40초', '40~45초', '45~50초', '50~55초', '55~60초', '60~90초', '90~120초', '120~150초', '150~180초', '180~360초', '360~540초', '540~720초', '720초~']

    if df_source.empty:
        empty_s1 = pd.DataFrame(columns=['작업자명', '작업수', f'0~{target_sec}초 작업 수', f'{target_sec}초이후 작업 수', f'0~{target_sec}초 작업시간 총합', f'{target_sec}초이후 작업시간 총합', '생산성(초)', '생산성(시간)'])
        empty_s2 = pd.DataFrame(columns=['작업자명'] + labels + ['총수량'])
        empty_s3 = pd.DataFrame(columns=df_source.columns if not df_source.empty else ['작업자', '작업일시', '주문번호', '주문 유형'])
        empty_s4 = pd.DataFrame(columns=df_source.columns if not df_source.empty else ['작업자', '작업일시', '주문번호', '주문 유형'])
        empty_s5 = pd.DataFrame()
        return empty_s1, empty_s2, empty_s3, empty_s4, empty_s5

    df_src = df_source.copy()
    df_src['작업일시'] = pd.to_datetime(df_src['작업일시'])
    s3_df = df_src.sort_values(by=['작업자', '작업일시'], ascending=[True, True]).reset_index(drop=True)
    s4_df = s3_df.copy().sort_values(by=['작업자', '작업일시'], ascending=[True, True]).drop_duplicates(subset=['작업자', '주문번호'], keep='first').reset_index(drop=True)

    processors = s4_df['작업자'].unique()
    columns_to_combine, stat_records, detailed_records = [], [], []

    for processor in processors:
        p_df = s4_df[s4_df['작업자'] == processor].copy().sort_values('작업일시', ascending=True).reset_index(drop=True)
        p_df['주문번호_전'] = p_df['주문번호'].shift(1)
        p_df['작업일시_전'] = p_df['작업일시'].shift(1)
        p_df = p_df.rename(columns={'주문번호': '주문번호_후', '작업일시': '작업일시_후'})
        p_df['작업간격_초'] = (p_df['작업일시_후'] - p_df['작업일시_전']).dt.total_seconds().fillna(0).astype(int)
        p_df['작업일시_전_str'] = p_df['작업일시_전'].dt.strftime('%Y-%m-%d %H:%M:%S').fillna('-')
        p_df['작업일시_후_str'] = p_df['작업일시_후'].dt.strftime('%Y-%m-%d %H:%M:%S')
        p_df['주문번호_전'] = p_df['주문번호_전'].fillna('-')

        p_res = p_df[['주문번호_전', '주문번호_후', '작업일시_전_str', '작업일시_후_str', '작업간격_초']].rename(columns={
            '주문번호_전': f"{processor}_주문번호_전", '주문번호_후': f"{processor}_주문번호_후",
            '작업일시_전_str': f"{processor}_작업일시_전", '작업일시_후_str': f"{processor}_작업일시_후", '작업간격_초': f"{processor}_작업간격_초"
        }).reset_index(drop=True)
        columns_to_combine.append(p_res)

        df_under_target = p_df[(p_df['작업간격_초'] >= 0) & (p_df['작업간격_초'] <= target_sec)]
        count_under_target = df_under_target.shape[0]
        sum_time_under_target = df_under_target['작업간격_초'].sum()
        df_over_target = p_df[p_df['작업간격_초'] > target_sec]
        count_over_target = df_over_target.shape[0]
        sum_time_over_target = df_over_target['작업간격_초'].sum()
        job_count = count_under_target + count_over_target

        productivity_sec = count_under_target / sum_time_under_target if sum_time_under_target > 0 else 0
        productivity_hour = productivity_sec * 3600

        stat_records.append({
            '작업자명': processor, '작업수': job_count,
            f'0~{target_sec}초 작업 수': count_under_target, f'{target_sec}초이후 작업 수': count_over_target,
            f'0~{target_sec}초 작업시간 총합': int(sum_time_under_target), f'{target_sec}초이후 작업시간 총합': int(sum_time_over_target),
            '생산성(초)': round(productivity_sec, 4), '생산성(시간)': round(productivity_hour, 1)
        })

        p_df['구간'] = pd.cut(p_df['작업간격_초'], bins=bins, labels=labels, include_lowest=True)
        counts = p_df['구간'].value_counts().reindex(labels, fill_value=0)
        detailed_record = {'작업자명': processor}
        for label in labels:
            detailed_record[label] = counts[label]
        detailed_record['총수량'] = len(p_df)
        detailed_records.append(detailed_record)

    s1_df = pd.DataFrame(stat_records)
    s2_df = pd.DataFrame(detailed_records)
    s5_df = pd.concat(columns_to_combine, axis=1) if columns_to_combine else pd.DataFrame()

    if not s1_df.empty:
        s1_df = s1_df[s1_df['작업자명'].notna() & (s1_df['작업자명'].astype(str).str.strip() != '') & (s1_df['작업자명'].astype(str).str.lower() != 'nan')]
    if not s2_df.empty:
        s2_df = s2_df[s2_df['작업자명'].notna() & (s2_df['작업자명'].astype(str).str.strip() != '') & (s2_df['작업자명'].astype(str).str.lower() != 'nan')]

    return s1_df, s2_df, s3_df, s4_df, s5_df

if uploaded_files:
    if st.button("작업자별 출고 작업시간 생산성 분석 시작"):
        # 실시간 분석 로그를 찍어줄 빈 공간 확보
        log_space = st.empty()
        
        try:
            output = BytesIO()
            sorted_files = sorted(uploaded_files, key=lambda x: x.name)
            
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                for idx, uploaded_file in enumerate(sorted_files):
                    base_file_name = os.path.splitext(uploaded_file.name)[0]
                    
                    # 💡 실시간 추적 확인용 로그출력
                    log_space.info(f"⏳ [{idx+1}/{len(sorted_files)}] 파일 분석 중: {uploaded_file.name}")
                    
                    df_main = pd.read_excel(uploaded_file)
                    
                    # 💡 원본 필수 컬럼 검증
                    required_cols = ['작업자', '작업일시', '주문번호', '주문 유형']
                    missing_cols = [c for c in required_cols if c not in df_main.columns]
                    if missing_cols:
                        st.error(f"❌ '{uploaded_file.name}' 파일에 {missing_cols} 열이 없습니다. 열 이름을 확인하세요.")
                        st.stop()

                    df_dang = df_main[df_main['주문 유형'] == '당특'].copy()
                    df_ilban = df_main[df_main['주문 유형'] == '일반'].copy()
                    
                    # 데이터 건수 중간 점검 로그
                    st.write(f"🔍 `{base_file_name}` 건수 분할 확인 ➔ 당특: {len(df_dang)}건 / 일반: {len(df_ilban)}건")

                    dang_s1, dang_s2, dang_s3, dang_s4, dang_s5 = generate_5_sheets(df_dang, target_seconds)
                    ilban_s1, ilban_s2, ilban_s3, ilban_s4, ilban_s5 = generate_5_sheets(df_ilban, target_seconds)

                    dang_summary_part = dang_s1.copy() if not dang_s1.empty else pd.DataFrame()
                    if not dang_summary_part.empty:
                        dang_summary_part.insert(0, '주문 유형', '당특')
                        dang_summary_part = dang_summary_part.rename(columns={'작업자명': '작업자'})

                    ilban_summary_part = ilban_s1.copy() if not ilban_s1.empty else pd.DataFrame()
                    if not ilban_summary_part.empty:
                        ilban_summary_part.insert(0, '주문 유형', '일반')
                        ilban_summary_part = ilban_summary_part.rename(columns={'작업자명': '작업자'})

                    if not dang_summary_part.empty or not ilban_summary_part.empty:
                        overall_summary = pd.concat([dang_summary_part, ilban_summary_part], axis=0).reset_index(drop=True)
                    else:
                        # 💡 요약 데이터프레임이 강제로 만들어지게 컬럼과 더미 데이터를 수동 주입
                        overall_summary = pd.DataFrame(columns=['주문 유형', '작업자', '작업수'], data=[['데이터 없음', '데이터 없음', 0]])

                    prefix = f"{base_file_name}_" if len(sorted_files) > 1 else ""

                    # 시트 내보내기 강제 실행
                    overall_summary.to_excel(writer, sheet_name=f'{prefix}유형별전체요약', index=False)
                    dang_s1.to_excel(writer, sheet_name=f'{prefix}당특_생산성분석', index=False)
                    dang_s2.to_excel(writer, sheet_name=f'{prefix}당특_생산성_상세', index=False)
                    dang_s3.to_excel(writer, sheet_name=f'{prefix}당특_작업자별정렬', index=False)
                    dang_s4.to_excel(writer, sheet_name=f'{prefix}당특_작업자별주문유니크', index=False)
                    dang_s5.to_excel(writer, sheet_name=f'{prefix}당특_상세작업시간', index=False)
                    ilban_s1.to_excel(writer, sheet_name=f'{prefix}일반_생산성분석', index=False)
                    ilban_s2.to_excel(writer, sheet_name=f'{prefix}일반_생산성_상세', index=False)
                    ilban_s3.to_excel(writer, sheet_name=f'{prefix}일반_작업자별정렬', index=False)
                    ilban_s4.to_excel(writer, sheet_name=f'{prefix}일반_작업자별주문유니크', index=False)
                    ilban_s5.to_excel(writer, sheet_name=f'{prefix}일반_상세작업시간', index=False)

            processed_data = output.getvalue()
            st.balloons()
            st.success("🎉 모든 파일 분석 완료!")
            
            st.download_button(
                label="📥 가공된 유형별 통합 엑셀 다운로드",
                data=processed_data,
                file_name=f"유형별_종합생산성분석_{target_seconds}초기준.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"⚠️ 처리 중 오류가 발생했습니다: {e}")
            st.info("정확히 어떤 파일의 데이터를 처리하다가 문제가 생겼는지 위 로그 화면을 확인해 보세요.")
