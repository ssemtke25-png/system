import streamlit as st
import pandas as pd
import io
import openpyxl

# ==========================================
# [1. 웹 페이지 기본 설정]
# ==========================================
st.set_page_config(page_title="데이터 자동 검증 및 취합 시스템", page_icon="📊", layout="wide")

# ==========================================
# [2. 보안 잠금장치]
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
            if pwd == "7777": 
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("❌ 비밀번호가 일치하지 않습니다.")
    st.stop()

# ==========================================
# [3. 메인 화면: 파일 업로드 및 스마트 취합]
# ==========================================
st.markdown("### 📊 지적재조사 만능 데이터 취합 및 검증 시스템")
st.info("💡 **[보안 안내]** 업로드된 파일은 서버 하드디스크에 1바이트도 저장되지 않으며, 계산 즉시 메모리에서 영구 삭제(휘발)됩니다.")

st.markdown("---")
st.markdown("#### ⚙️ 시스템 작동 옵션")
run_validation = st.checkbox("🔍 [테스트용] A+B=C 오류 검증 기능 켜기 (체크 해제 시 '단순 취합'만 진행됩니다)", value=False)
st.markdown("---")

uploaded_files = st.file_uploader("취합할 시·군의 엑셀 파일(.xlsx)을 모두 드래그해서 놓으세요.", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    st.success(f"✅ 총 {len(uploaded_files)}개의 파일이 메모리에 업로드되었습니다.")
    
    if st.button("🚀 시스템 가동 시작", type="primary", use_container_width=True):
        
        with st.spinner("🔍 데이터 처리 중입니다..."):
            try:
                # 🚨 [1단계: 스위치가 켜져 있을 때만 오류 검증 진행] 🚨
                if run_validation:
                    for f in uploaded_files:
                        temp_file = io.BytesIO(f.read())
                        
                        try:
                            wb_temp = openpyxl.load_workbook(temp_file, data_only=True)
                        except Exception as e:
                            st.error(f"🚨 **파일 인식 실패! 강제 중단!** 🚨")
                            st.error(f"❌ 범인 파일명: **[{f.name}]**")
                            st.stop()
                        
                        for sheet_name in wb_temp.sheetnames:
                            ws_test = wb_temp[sheet_name]
                            
                            val_A = ws_test.cell(row=1, column=1).value or 0
                            val_B = ws_test.cell(row=1, column=2).value or 0
                            val_C = ws_test.cell(row=1, column=3).value or 0
                            
                            if isinstance(val_A, (int, float)) and isinstance(val_B, (int, float)) and isinstance(val_C, (int, float)):
                                if (val_A + val_B) != val_C:
                                    st.error(f"🚨 **검증 실패! 강제 중단!** 🚨")
                                    st.error(f"❌ 범인 파일명: **[{f.name}]** (범행 장소: **{sheet_name}** 시트)")
                                    st.warning(f"➡️ 눈에 보이지 않는 엑셀 수식 캐시 오류이거나 잘못된 값입니다. 파일을 엑셀로 열어 [저장] 후 다시 올려주세요!")
                                    st.stop() 
                else:
                    st.info("⏩ 오류 검증 단계를 건너뛰고 '스마트 취합'을 바로 시작합니다.")

                # ----------------------------------------------------
                # 🟢 [2단계: 스마트 취합 진행] 🟢
                for f in uploaded_files: f.seek(0) 
                
                base_file = io.BytesIO(uploaded_files[0].read())
                wb_base = openpyxl.load_workbook(base_file)
                
                for f in uploaded_files[1:]:
                    f.seek(0)
                    temp_file = io.BytesIO(f.read())
                    wb_temp = openpyxl.load_workbook(temp_file, data_only=True)
                    
                    for sheet_name in wb_base.sheetnames:
                        if sheet_name in wb_temp.sheetnames:
                            ws_base = wb_base[sheet_name]
                            ws_temp = wb_temp[sheet_name]
                            
                            for row in range(1, ws_base.max_row + 1):
                                for col in range(1, ws_base.max_column + 1):
                                    val_base = ws_base.cell(row=row, column=col).value
                                    val_temp = ws_temp.cell(row=row, column=col).value
                                    
                                    if isinstance(val_base, str) and val_base.startswith('='):
                                        continue
                                    
                                    is_base_num = isinstance(val_base, (int, float))
                                    is_temp_num = isinstance(val_temp, (int, float))
                                    
                                    if is_base_num or is_temp_num:
                                        num_base = val_base if is_base_num else 0
                                        num_temp = val_temp if is_temp_num else 0
                                        
                                        ws_base.cell(row=row, column=col).value = num_base + num_temp

                output = io.BytesIO()
                wb_base.save(output)
                output.seek(0)
                
                st.success("✨ 취합이 완벽하게 끝났습니다!")
                
                st.download_button(
                    label="📥 최종 취합본 다운로드 (엑셀)",
                    data=output,
                    file_name="경상북도_지적재조사_최종취합본.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
            except Exception as e:
                st.error(f"🚨 예상치 못한 시스템 오류가 발생했습니다: {e}")
