import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import os

# 1. 웹페이지 레이아웃 및 타이틀 설정
st.set_page_config(page_title="작업자별 출고 작업시간 생산성 분석", layout="centered")

st.title("🏭 주문 유형별 통합 생산성 분석 프로그램")
st.write("엑셀 파일을 업로드하고 기준 초를 지정하여 마스터 요약이 포함된 11개 시트의 리포트를 만들어 보세요.")

# 2. 파일 다중 업로드 및 기준 초 입력 섹션
uploaded_files = st.file_uploader(
    "출고내역 엑셀 파일(xlsx)을 업로드해주세요. (여러 파일 동시 업로드 가능)",
    type=["xlsx"],
    accept_multiple_files=True
)

# 사용자가 기준 초를 직접 입력할 수 있는 숫자 입력창 (기본값은 60초)
target_seconds = st.number_input(
    "🔄 작업시간 기준(초)를 입력해 주세요.",
    min_value=1,
    max_value=3600,
    value=60,
    step=1,
    help="해당 기준 초를 초과하는 작업은 (쉬는 시간 or 특이)로 간주하여 생산성 계산식 내 (작업시간)에서 제외합니다."
)


# =====================================================================
# [기능 정의] 특정 데이터프레임을 받아 5개 분석 시트셋을 만드는 함수
# =====================================================================
def generate_5_sheets(df_source, target_sec):
    bins = [0, 5, 10, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60, 90, 120, 150, 180, 360, 540, 720, np.inf]
    labels = [
        '0~5초', '5~10초', '10~15초', '15~20초', '20~25초', '25~30초',
        '30~35초', '35~40초', '40~45초', '45~50초', '50~55초', '55~60초',
        '60~90초', '90~120초', '120~150초', '150~180초', '180~360초', '360~540초', '540~720초', '720초~'
    ]

    # 데이터가 완전히 비어있어도 헤더를 수동 정의하여 openpyxl의 시트 소멸 에러 원천 차단
    if df_source.empty:
        empty_s1 = pd.DataFrame(columns=['작업자명', '작업수', f'0~{target_sec}초 작업 수', f'{target_sec}초이후 작업 수', f'0~{target_sec}초 작업시간 총합', f'{target_sec}초이후 작업시간 총합', '생산성(초)', '생산성(시간)'])
        empty_s2 = pd.DataFrame(columns=['작업자명'] + labels + ['총수량'])
        empty_s3 = pd.DataFrame(columns=['작업자', '작업일시', '주문번호', '주문 유형'])
        empty_s4 = pd.DataFrame(columns=['작업자', '작업일시', '주문번호', '주문 유형'])
        empty_s5 = pd.DataFrame(columns=['안내'], data=[['데이터 없음']])
        return empty_s1, empty_s2, empty_s3, empty_s4, empty_s5

    df_src = df_source.copy()
    df_src['작업일시'] = pd.to_datetime(df_src['작업일시'])

    # 1. 작업자별정렬
    s3_df = df_src.sort_values(by=['작업자', '작업일시'], ascending=[True, True]).reset_index(drop=True)

    # 2. 작업자별주문유니크정렬
    s4_df = s3_df.copy().sort_values(by=['작업자', '작업일시'], ascending=[True, True])
    s4_df = s4_df.drop_duplicates(subset=['작업자', '주문번호'], keep='first').reset_index(drop=True)

    processors = s4_df['작업자'].unique()

    columns_to_combine = []
    stat_records = []
    detailed_records = []

    for processor in processors:
        p_df = s4_df[s4_df['작업자'] == processor].copy().sort_values('작업일시', ascending=True).reset_index(drop=True)

        p_df['주문번호_전'] = p_df['주문번호'].shift(1)
        p_df['작업일시_전'] = p_df['작업일시'].shift(1)

        p_df = p_df.rename(columns={'주문번호': '주문번호_후', '작업일시': '작업일시_후'})
        p_df['작업간격_초'] = (p_df['작업일시_후'] - p_df['작업일시_전']).dt.total_seconds()
        p_df['작업간격_초'] = p_df['작업간격_초'].fillna(0).astype(int)

        p_df['작업일시_전_str'] = p_df['작업일시_전'].dt.strftime('%Y-%m-%d %H:%M:%S').fillna('-')
        p_df['작업일시_후_str'] = p_df['작업일시_후'].dt.strftime('%Y-%m-%d %H:%M:%S')
        p_df['주문번호_전'] = p_df['주문번호_전'].fillna('-')

        col_prev_order = f"{processor}_주문번호_전"
        col_next_order = f"{processor}_주문번호_후"
        col_prev_time = f"{processor}_작업일시_전"
        col_next_time = f"{processor}_작업일시_후"
        col_diff_sec = f"{processor}_작업간격_초"

        p_res = p_df[['주문번호_전', '주문번호_후', '작업일시_전_str', '작업일시_후_str', '작업간격_초']].rename(columns={
            '주문번호_전': col_prev_order, '주문번호_후': col_next_order,
            '작업일시_전_str': col_prev_time, '작업일시_후_str': col_next_time, '작업간격_초': col_diff_sec
        }).reset_index(drop=True)
        columns_to_combine.append(p_res)

        # 시트 1 계산
        df_under_target = p_df[(p_df['작업간격_초'] >= 0) & (p_df['작업간격_초'] <= target_sec)]
        count_under_target = df_under_target.shape[0]
        sum_time_under_target = df_under_target['작업간격_초'].sum()

        df_over_target = p_df[p_df['작업간격_초'] > target_sec]
        count_over_target = df_over_target.shape[0]
        sum_time_over_target = df_over_target['작업간격_초'].sum()

        job_count = count_under_target + count_over_target

        if sum_time_under_target > 0:
            productivity_sec = count_under_target / sum_time_under_target
            productivity_hour = productivity_sec * 3600
        else:
            productivity_sec = 0
            productivity_hour = 0

        stat_records.append({
            '작업자명': processor, '작업수': job_count,
            f'0~{target_sec}초 작업 수': count_under_target, f'{target_sec}초이후 작업 수': count_over_target,
            f'0~{target_sec}초 작업시간 총합': int(sum_time_under_target), f'{target_sec}초이후 작업시간 총합': int(sum_time_over_target),
            '생산성(초)': round(productivity_sec, 4), '생산성(시간)': round(productivity_hour, 1)
        })

        # 시트 2 계산
        p_df['구간'] = pd.cut(p_df['작업간격_초'], bins=bins, labels=labels, include_lowest=True)
        counts = p_df['구간'].value_counts().reindex(labels, fill_value=0)

        detailed_record = {'작업자명': processor}
        for label in labels:
            detailed_record[label] = counts[label]
        detailed_record['총수량'] = len(p_df)
        detailed_records.append(detailed_record)

    s1_df = pd.DataFrame(stat_records)
    s2_df = pd.DataFrame(detailed_records)
    s5_df = pd.concat(columns_to_combine, axis=1) if columns_to_combine else pd.DataFrame(columns=['안내'], data=[['상세 데이터 없음']])

    # 보고서 요약본 내 유령 작업자(NaN) 걸러내는 필터링
    if not s1_df.empty:
        s1_df = s1_df[s1_df['작업자명'].notna() & (s1_df['작업자명'].astype(str).str.strip() != '') & (s1_df['작업자명'].astype(str).str.lower() != 'nan')]
    if not s2_df.empty:
        s2_df = s2_df[s2_df['작업자명'].notna() & (s2_df['작업자명'].astype(str).str.strip() != '') & (s2_df['작업자명'].astype(str).str.lower() != 'nan')]

    return s1_df, s2_df, s3_df, s4_df, s5_df


# =====================================================================
# [단계 3] 웹 프로세스 실구동 제어
# =====================================================================
if uploaded_files:
    st.success(f"총 {len(uploaded_files)}개의 파일이 정상적으로 로드되었습니다!")

    if st.button("작업자별 출고 작업시간 생산성 분석 시작"):
        try:
            with st.spinner('데이터를 분리하여 마스터 대조 요약 및 유형별 5개 시트 리포트를 생성 중입니다...'):

                output = BytesIO()
                sorted_files = sorted(uploaded_files, key=lambda x: x.name)
                file_names_summary = []

                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    for uploaded_file in sorted_files:
                        base_file_name = os.path.splitext(uploaded_file.name)[0]
                        file_names_summary.append(base_file_name)

                        df_main = pd.read_excel(uploaded_file)

                        # 필수 컬럼 검증
                        required_cols = ['작업자', '작업일시', '주문번호', '주문 유형']
                        missing_cols = [c for c in required_cols if c not in df_main.columns]
                        if missing_cols:
                            st.error(f"❌ '{uploaded_file.name}' 파일에 {missing_cols} 열이 없습니다. 열 이름을 확인하세요.")
                            st.stop()

                        # '당특'과 '일반' 분할
                        df_dang = df_main[df_main['주문 유형'] == '당특'].copy()
                        df_ilban = df_main[df_main['주문 유형'] == '일반'].copy()

                        # 함수 기동 및 결과 도출
                        dang_s1, dang_s2, dang_s3, dang_s4, dang_s5 = generate_5_sheets(df_dang, target_seconds)
                        ilban_s1, ilban_s2, ilban_s3, ilban_s4, ilban_s5 = generate_5_sheets(df_ilban, target_seconds)

                        # '유형별전체요약' 마스터 탭 데이터 빌드
                        dang_summary_part = dang_s1.copy() if not dang_s1.empty else pd.DataFrame()
                        if not dang_summary_part.empty:
                            dang_summary_part.insert(0, '주문 유형', '당특')
                            dang_summary_part = dang_summary_part.rename(columns={'작업자명': '작업자'})

                        ilban_summary_part = ilban_s1.copy() if not ilban_s1.empty else pd.DataFrame()
                        if not ilban_summary_part.empty:
                            ilban_summary_part.insert(0, '주문 유형', '일반')
                            ilban_summary_part = ilban_summary_part.rename(columns={'작업자명': '작업자'})

                        # 💡 [핵심 해결 지점] 양쪽 데이터 요약이 모두 비어있을 때 '데이터 없음' 행을 강제 주입하여 openpyxl 시트 완전 보존
                        if not dang_summary_part.empty or not ilban_summary_part.empty:
                            overall_summary = pd.concat([dang_summary_part, ilban_summary_part], axis=0).reset_index(drop=True)
                        else:
                            overall_summary = pd.DataFrame(
                                columns=['주문 유형', '작업자', '작업수', f'0~{target_seconds}초 작업 수', f'{target_seconds}초이후 작업 수', f'0~{target_seconds}초 작업시간 총합', f'{target_seconds}초이후 작업시간 총합', '생산성(초)', '생산성(시간)'],
                                data=[['데이터 없음', '데이터 없음', 0, 0, 0, 0, 0, 0, 0]]
                            )

                        prefix = f"{base_file_name}_" if len(sorted_files) > 1 else ""

                        # 11개 시트 가시화 패키징 빌드
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
            st.success("🎉 분석 성공")

            # 다운로드 파일명 빌드 세팅
            display_name = f"{file_names_summary[0]}_외" if len(file_names_summary) > 1 else f"{file_names_summary[0]}"
            final_download_name = f"{display_name}_작업자별_출고_작업시간_{target_seconds}초기준.xlsx"

            st.download_button(
                label="📥 가공된 유형별 통합 마스터 엑셀 다운로드",
                data=processed_data,
                file_name=final_download_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"⚠️ 처리 중 오류가 발생했습니다: {e}")
