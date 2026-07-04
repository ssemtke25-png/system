import streamlit as st
import pandas as pd
import io
import re

# 파일명에서 불필요한 단어를 빼고 '시군명'만 똑똑하게 추출해 내는 함수
def extract_city_name(filename):
    # 확장자 제거
    name = re.sub(r'\.xlsx|\.xls|\.csv', '', filename)
    # 불필요한 단어들 제거 (명단, 참석자, 행사, 총괄표, 취합 등)
    name = re.sub(r'명단|참석자|행사|총괄표|취합|제출', '', name).strip()
    # 괄호와 괄호 안 내용 제거
    name = re.sub(r'\[.*?\]|\(.*?\)', '', name).strip()
    return name

def render():
    st.header("📝 ⑨ 행사 명단 취합 (순서대로 무조건 복붙!)")
    st.info("총괄표 양식을 먼저 올리고 시군 파일들을 올리면, **열(Column) 이름이 달라도 사람이 복사+붙여넣기 하듯 총괄표 순서대로 무식하고 확실하게 꿰매줍니다.**")

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("#### 1. 총괄표 엑셀 업로드")
        master_file = st.file_uploader("행사 명단(총괄표) 엑셀 파일", type=['xlsx', 'xls', 'csv'], key="tab9_master")
        st.caption("※ 이 파일의 첫 번째 줄을 기준으로 데이터가 1열부터 순서대로 들어갑니다.")
        
    with col2:
        st.markdown("#### 2. 시군 제출 파일 다중 업로드")
        sub_files = st.file_uploader("각 시군 명단 파일 (여러 개 선택 가능)", type=['xlsx', 'xls', 'csv'], accept_multiple_files=True, key="tab9_subs")
        st.caption("※ 빈칸으로 제출된 '시군명'은 파일명에서 추출해 자동으로 채워줍니다.")

    if master_file and sub_files:
        st.markdown("---")
        if st.button("🚀 총괄표 틀에 맞춰 순서대로 병합하기", type="primary", use_container_width=True):
            try:
                with st.spinner("엑셀 파일들을 하나로 이어 붙이고 있습니다..."):
                    # 1. 총괄표 읽어오기
                    if master_file.name.endswith('.csv'):
                        master_df = pd.read_csv(master_file)
                    else:
                        master_df = pd.read_excel(master_file)
                        
                    master_cols = master_df.columns.tolist()
                    combined_list = []
                    
                    # 총괄표 자체에 이미 작성된 데이터가 있다면 살려둡니다
                    cleaned_master = master_df.dropna(how='all')
                    if not cleaned_master.empty:
                        combined_list.append(cleaned_master)

                    # 시군명 역할을 하는 컬럼이름 찾기 ("시군명", "시군", "지역" 등)
                    region_col = next((col for col in master_cols if "시군" in str(col) or "지역" in str(col)), None)

                    # 2. 업로드된 시군 파일들 순회하며 병합
                    for file in sub_files:
                        if file.name.endswith('.csv'):
                            sub_df = pd.read_csv(file)
                        else:
                            sub_df = pd.read_excel(file)
                        
                        # 완전히 텅 빈 깡통 행은 제거
                        sub_df = sub_df.dropna(how='all')
                        if sub_df.empty:
                            continue

                        # [핵심 로직] 열 이름 무시! 오로지 '데이터 값'만 빼와서 위치대로 맞추기
                        data_values = sub_df.values.tolist()
                        processed_data = []
                        master_len = len(master_cols)
                        
                        for row in data_values:
                            # 만약 시군 파일의 열 개수가 총괄표보다 많으면 총괄표 개수만큼 자르기
                            if len(row) >= master_len:
                                processed_row = list(row[:master_len])
                            # 만약 시군 파일의 열 개수가 총괄표보다 적으면 모자란 만큼 빈칸(None) 채우기
                            else:
                                processed_row = list(row) + [None] * (master_len - len(row))
                            processed_data.append(processed_row)
                            
                        # 자르고 늘린 데이터를 총괄표 껍데기(컬럼)에 쏙 집어넣기
                        new_sub_df = pd.DataFrame(processed_data, columns=master_cols)
                        
                        # 시군명 칸이 비어있다면 파일명에서 추출해 채워주기
                        if region_col:
                            city_name = extract_city_name(file.name)
                            # 공백 문자로만 된 경우를 빈칸(NA)으로 확실히 처리 후 채움
                            new_sub_df[region_col] = new_sub_df[region_col].replace(r'^\s*$', pd.NA, regex=True)
                            new_sub_df[region_col] = new_sub_df[region_col].fillna(city_name)
                            
                        combined_list.append(new_sub_df)

                    # 3. 최종 데이터프레임 완성
                    if combined_list:
                        final_df = pd.concat(combined_list, ignore_index=True)
                    else:
                        final_df = pd.DataFrame(columns=master_cols)

                    st.success(f"✅ 총 {len(sub_files)}개의 파일을 무식하고 확실하게 이어 붙였습니다! (총 {len(final_df)}줄)")
                    
                    # 미리보기 제공
                    st.markdown("### 📊 취합 결과 미리보기")
                    st.dataframe(final_df, use_container_width=True)

                    # 4. 엑셀 다운로드 변환
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        final_df.to_excel(writer, index=False, sheet_name='취합명단')
                    
                    # 가운데 정렬 다운로드 버튼
                    st.markdown("<br>", unsafe_allow_html=True)
                    col_dl1, col_dl2, col_dl3 = st.columns([1, 2, 1])
                    with col_dl2:
                        st.download_button(
                            label="📥 완성된 총괄 명단 엑셀 다운로드",
                            data=output.getvalue(),
                            file_name="행사명단_최종_취합본.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )

            except Exception as e:
                st.error(f"명단 병합 중 오류가 발생했습니다: {e}")
                st.info("💡 오류 해결 팁: 제출된 엑셀 파일 중 손상되었거나 암호가 걸린 파일이 있는지 확인해 주세요.")