import streamlit as st
import pandas as pd
import io
import openpyxl # 엑셀 서식을 유지하기 위한 필수 도구

# ==========================================
# [1. 웹 페이지 기본 설정]
# ==========================================
st.set_page_config(page_title="데이터 자동 검증 및 취합 시스템", page_icon="📊", layout="wide")

# ==========================================
# [2. 강력한 보안 잠금장치 (비밀번호)]
# ==========================================
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.markdown("<h2 style='text-align: center;'>🔐 시스템 보안 접속</h2>", unsafe_allow_html=True)
    st.caption("<p style='text-align: center;'>본 시스템은 인가된 관리자만 접근할 수 있습니다.</p>", unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 1, 1])
    with col2:
        pwd = st.text_input("보안 비밀번호 4자리", type="password", placeholder="비밀번호 입력")
        if st.button("시스템 입장", use_container_width=True):
            if pwd == "7777":  # 🚨 원하시는 비밀번호로 변경하세요!
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("❌ 비밀번호가 일치하지 않습니다.")
    st.stop()

# ==========================================
# [3. 메인 화면: 파일 업로드 및 스마트 취합]
# ==========================================
st.markdown("### 📊 지적재조사 만능 데이터 취합 시스템")
st.info("💡 **[보안 안내]** 업로드된 파일은 서버 하드디스크에 1바이트도 저장되지 않으며, 계산 즉시 메모리에서 영구 삭제(휘발)됩니다.")

uploaded_files = st.file_uploader("취합할 시·군의 엑셀 파일(.xlsx)을 모두 드래그해서 놓으세요.", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    st.success(f"✅ 총 {len(uploaded_files)}개의 파일이 메모리에 안전하게 업로드되었습니다.")
    
    if st.button("🚀 만능 자동 취합 시작", type="primary", use_container_width=True):
        
        if len(uploaded_files) < 2:
            st.warning("⚠️ 합산하려면 최소 2개 이상의 파일이 필요합니다.")
        else:
            with st.spinner("🔍 스마트 스캐너 가동 중... 숫자만 찾아내어 서식 파괴 없이 합산합니다!"):
                try:
                    # 1. 첫 번째 파일을 '기준 뼈대'로 메모리에서 읽어오기 (서식 유지)
                    base_file = io.BytesIO(uploaded_files[0].read())
                    wb_base = openpyxl.load_workbook(base_file)
                    
                    # 2. 두 번째 파일부터 차례대로 읽으면서 합산
                    for f in uploaded_files[1:]:
                        temp_file = io.BytesIO(f.read())
                        wb_temp = openpyxl.load_workbook(temp_file, data_only=True) # 수식 대신 결과값만 읽기
                        
                        # 파일 안의 모든 시트(Sheet)를 똑같이 순회
                        for sheet_name in wb_base.sheetnames:
                            if sheet_name in wb_temp.sheetnames:
                                ws_base = wb_base[sheet_name]
                                ws_temp = wb_temp[sheet_name]
                                
                                # 모든 칸을 샅샅이 뒤지며 숫자만 더하기
                                for row in range(1, ws_base.max_row + 1):
                                    for col in range(1, ws_base.max_column + 1):
                                        val_base = ws_base.cell(row=row, column=col).value
                                        val_temp = ws_temp.cell(row=row, column=col).value
                                        
                                        # 두 칸 모두 '숫자(정수 또는 소수)'일 때만 합산! (문자나 빈칸은 완벽 무시)
                                        if isinstance(val_base, (int, float)) and isinstance(val_temp, (int, float)):
                                            ws_base.cell(row=row, column=col).value = val_base + val_temp

                    # 3. 계산 완료된 파일을 허공(메모리)에 저장
                    output = io.BytesIO()
                    wb_base.save(output)
                    output.seek(0)
                    
                    st.success("✨ 취합이 완벽하게 끝났습니다! 아래 버튼을 눌러 결과물을 다운로드하세요.")
                    
                    # 4. 다운로드 버튼 생성
                    st.download_button(
                        label="📥 최종 취합본 다운로드 (엑셀)",
                        data=output,
                        file_name="경상북도_지적재조사_최종취합본.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                    
                except Exception as e:
                    st.error(f"🚨 엑셀 처리 중 오류가 발생했습니다. 모든 시군의 엑셀 양식이 동일한지 확인해주세요. (오류내용: {e})")
