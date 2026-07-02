import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
import io
import zipfile
from copy import copy
from pathlib import Path

# 지역 순서 (합계.xlsx 기준)
REGION_ORDER = [
    '포항시남구', '포항시북구', '경주시', '구미시', '청도군', 
    '영천시', '칠곡군', '성주군', '고령군', '성중시'
]

def render():
    st.markdown("### 📊 과징금·이행강제금 취합 시스템")
    
    col1, col2 = st.columns(2)
    
    with col1:
        template_file = st.file_uploader(
            "📋 합계.xlsx (템플릿)",
            type=["xlsx"],
            key="tab8_template"
        )
    
    with col2:
        region_files = st.file_uploader(
            "📁 시군 파일 (다중 선택)",
            type=["xlsx"],
            accept_multiple_files=True,
            key="tab8_regions"
        )
    
    if template_file and region_files:
        if st.button("🚀 취합 시작", key="tab8_process"):
            try:
                # 시군 파일명으로 지역 자동 추출
                region_data = {}
                for file in region_files:
                    region_name = extract_region_name(file.name)
                    if region_name:
                        region_data[region_name] = file
                
                # 지역 순서대로 정렬
                sorted_regions = [r for r in REGION_ORDER if r in region_data]
                missing_regions = [r for r in REGION_ORDER if r not in region_data]
                
                if missing_regions:
                    st.warning(f"⚠️ 누락된 지역: {', '.join(missing_regions)}")
                
                # 취합 실행
                output_buffer = merge_files(
                    template_file,
                    [region_data[r] for r in sorted_regions],
                    sorted_regions
                )
                
                # 검증
                validation_result = validate_totals(output_buffer)
                
                if validation_result['status'] == 'error':
                    st.error(f"❌ 검증 실패: {validation_result['message']}")
                    st.warning(f"과징금 미수납액(시트2 G5): {validation_result['actual_g5']}")
                    st.warning(f"이행강제금 미수납액(시트2 J5): {validation_result['actual_j5']}")
                else:
                    st.success("✅ 취합 완료!")
                    st.info(f"✓ 과징금 미수납액: {validation_result['actual_g5']:,}")
                    st.info(f"✓ 이행강제금 미수납액: {validation_result['actual_j5']:,}")
                    
                    st.download_button(
                        label="📥 합계.xlsx 다운로드",
                        data=output_buffer.getvalue(),
                        file_name="합계_최종.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            
            except Exception as e:
                st.error(f"❌ 오류 발생: {str(e)}")
                st.exception(e)

def extract_region_name(filename):
    """파일명에서 지역명 추출"""
    filename = filename.replace('.xlsx', '')
    for region in REGION_ORDER:
        if region in filename:
            return region
    return None

def merge_files(template_file, region_files, region_names):
    """템플릿 + 시군 파일 병합"""
    
    # 템플릿 로드
    template_wb = load_workbook(template_file)
    sheet1_template = template_wb['현황보고']
    sheet2_template = template_wb['미수납조서']
    
    # 시트2 미수납조서 헤더 저장
    sheet2_headers = []
    for row in sheet2_template.iter_rows(min_row=1, max_row=3, values_only=False):
        sheet2_headers.append(row)
    
    sheet2_data_start = 4  # 미수납조서 데이터 시작행
    current_row = sheet2_data_start
    
    # 각 시군 파일 처리
    for region_file, region_name in zip(region_files, region_names):
        region_wb = load_workbook(region_file)
        region_sheet1 = region_wb['현황보고']
        region_sheet2 = region_wb['미수납조서']
        
        # ===== 시트1: 현황보고 데이터 복사 =====
        copy_sheet1_data(sheet1_template, region_sheet1, region_name)
        
        # ===== 시트2: 미수납조서 데이터 복사 (세로 이어붙이기) =====
        current_row = copy_sheet2_data(sheet2_template, region_sheet2, current_row)
    
    # 출력
    output = io.BytesIO()
    template_wb.save(output)
    output.seek(0)
    
    return output

def copy_sheet1_data(template_sheet, region_sheet, region_name):
    """
    시트1(현황보고)에서 입력 영역 데이터만 복사
    - 색칠된 셀과 수식은 보호
    - 입력 영역만 복사
    """
    
    # 지역 행 찾기 (A열에서 region_name 찾기)
    target_row = None
    for row in range(9, 25):  # 데이터 영역
        cell_value = template_sheet[f'A{row}'].value
        if cell_value and region_name in str(cell_value):
            target_row = row
            break
    
    if not target_row:
        st.warning(f"⚠️ {region_name}의 행을 찾을 수 없습니다.")
        return
    
    # 입력 영역 정의 (B:M, 과징금 누계 ~ 이행강제금 청구)
    # 행사 데이터만 복사 (구분열 제외, 색칠된 셀 제외)
    for col in range(2, 14):  # B ~ M
        source_cell = region_sheet.cell(row=9, column=col)  # 시군파일 시작행
        target_cell = template_sheet.cell(row=target_row, column=col)
        
        # 색칠된 셀 확인
        if is_cell_colored(target_cell):
            continue
        
        # 수식 확인
        if target_cell.data_type == 'f':
            continue
        
        # 데이터 복사
        if source_cell.value is not None:
            target_cell.value = source_cell.value
            copy_cell_style(source_cell, target_cell)

def copy_sheet2_data(template_sheet, region_sheet, start_row):
    """
    시트2(미수납조서) 데이터를 세로로 이어붙이기
    합계행(Row 4)은 제외하고, 데이터행(Row 5부터)부터 복사
    """
    
    current_row = start_row
    
    # 지역명 행 추가 (선택사항)
    region_name = region_sheet['A1'].value or "미분류"
    
    # 미수납조서 데이터 복사 (Row 5부터 마지막까지)
    for source_row in range(5, 100):  # 충분한 범위
        source_cell_a = region_sheet[f'A{source_row}']
        
        if source_cell_a.value is None:
            break
        
        # 행 복사 (A ~ M)
        for col in range(1, 14):  # A ~ M
            source_cell = region_sheet.cell(row=source_row, column=col)
            target_cell = template_sheet.cell(row=current_row, column=col)
            
            if source_cell.value is not None:
                target_cell.value = source_cell.value
                copy_cell_style(source_cell, target_cell)
        
        current_row += 1
    
    return current_row

def is_cell_colored(cell):
    """셀이 색칠되어 있는지 확인"""
    if cell.fill and cell.fill.start_color:
        color = str(cell.fill.start_color.rgb)
        return color not in ['00000000', 'FFFFFFFF', '00000000']
    return False

def copy_cell_style(source_cell, target_cell):
    """셀 스타일 복사"""
    if source_cell.font:
        target_cell.font = copy(source_cell.font)
    if source_cell.border:
        target_cell.border = copy(source_cell.border)
    if source_cell.alignment:
        target_cell.alignment = copy(source_cell.alignment)
    if source_cell.number_format:
        target_cell.number_format = copy(source_cell.number_format)

def validate_totals(output_buffer):
    """
    검증: 시트1 P9 (미수납액 합계) = 시트2 G5 (과징금) + J5 (이행강제금)
    """
    output_buffer.seek(0)
    wb = load_workbook(output_buffer)
    sheet1 = wb['현황보고']
    sheet2 = wb['미수납조서']
    
    # 시트2 G5, J5 값 추출
    g5_value = sheet2['G5'].value or 0
    j5_value = sheet2['J5'].value or 0
    
    total_expected = g5_value + j5_value
    
    # 시트1 P9 값 추출
    p9_value = sheet1['P9'].value or 0
    
    # 검증
    if isinstance(p9_value, str):
        try:
            p9_value = float(p9_value)
        except:
            p9_value = 0
    
    if abs(float(p9_value) - float(total_expected)) < 1:  # 오차 허용
        return {
            'status': 'success',
            'actual_g5': g5_value,
            'actual_j5': j5_value
        }
    else:
        return {
            'status': 'error',
            'message': f"미수납액 합계 불일치\n시트1 P9: {p9_value}\n기대값(G5+J5): {total_expected}",
            'actual_g5': g5_value,
            'actual_j5': j5_value
        }
