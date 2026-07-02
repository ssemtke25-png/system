import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.merge import MergedCellRange
import pandas as pd

# ✅ 정확한 경북 지역 순서 (실거래 기준)
REGION_ORDER_RE = [
    '포항남', '포항북', '경주', '김천', '안동', '구미', '영주', '영천', '상주', '문경',
    '경산', '군위', '의성', '청송', '영양', '영덕', '청도', '고령', '성주', '칠곡',
    '예천', '봉화', '울진', '울릉'
]

def collect_and_adjust_tab8(file_path, output_path):
    """
    Tab8 데이터를 수집하고 조정하는 함수
    - 모든 지역이 없어도 계속 진행
    - 지역순서대로 정렬
    """
    
    print("\n" + "="*60)
    print("📊 Tab8 데이터 수집 및 조정 시작")
    print("="*60)
    
    try:
        wb = openpyxl.load_workbook(file_path)
        ws_target = wb['tab8']
        
        # 기존 데이터 초기화 (헤더 제외)
        for row in ws_target.iter_rows(min_row=2):
            for cell in row:
                cell.value = None
        
        # 지역별 데이터 수집
        collected_data = []
        missing_regions = []
        error_regions = []
        
        for region in REGION_ORDER_RE:
            # 시트명 포맷: "tab8_지역명"
            sheet_name = f"tab8_{region}"
            
            try:
                if sheet_name not in wb.sheetnames:
                    print(f"⊘ {region:6} - 시트 없음 (스킵)")
                    missing_regions.append(region)
                    continue
                
                ws_source = wb[sheet_name]
                
                # 데이터 추출 (헤더 제외)
                region_data = []
                for row in ws_source.iter_rows(min_row=2, values_only=False):
                    row_data = [cell.value for cell in row]
                    if any(row_data):  # 빈 행 제외
                        region_data.append(row_data)
                
                if region_data:
                    collected_data.extend(region_data)
                    print(f"✓ {region:6} - {len(region_data):3}개 행 수집")
                else:
                    print(f"⚠ {region:6} - 데이터 없음 (스킵)")
                    
            except Exception as e:
                print(f"⚠ {region:6} - 오류: {str(e)[:30]} (스킵)")
                error_regions.append((region, str(e)))
                continue
        
        # 수집된 데이터를 tab8에 작성
        if collected_data:
            for row_idx, row_data in enumerate(collected_data, start=2):
                for col_idx, value in enumerate(row_data, start=1):
                    ws_target.cell(row=row_idx, column=col_idx, value=value)
            
            print(f"\n✓ 총 {len(collected_data)}개 행을 tab8에 작성")
        else:
            print("\n⚠ 수집된 데이터가 없습니다.")
        
        # 병합 셀 처리 (필요시)
        try:
            adjust_merged_cells(ws_target)
        except Exception as e:
            print(f"⚠ 병합 셀 조정 중 오류: {str(e)}")
        
        # 결과 요약
        print("\n" + "="*60)
        print(f"📈 수집 결과 요약")
        print("="*60)
        print(f"✓ 성공: {len(REGION_ORDER_RE) - len(missing_regions) - len(error_regions)}/{len(REGION_ORDER_RE)}")
        
        if missing_regions:
            print(f"⊘ 누락된 지역 ({len(missing_regions)}): {', '.join(missing_regions)}")
        
        if error_regions:
            print(f"⚠ 오류 발생 ({len(error_regions)}):")
            for region, error in error_regions:
                print(f"  - {region}: {error[:40]}")
        
        # 파일 저장
        wb.save(output_path)
        print(f"\n✓ 파일 저장 완료: {output_path}")
        print("="*60 + "\n")
        
        return True
        
    except Exception as e:
        print(f"\n❌ 전체 프로세스 오류: {str(e)}")
        return False

def adjust_merged_cells(ws):
    """
    병합된 셀 범위 재계산 및 조정
    """
    # 기존 병합 셀 제거
    merged_ranges = list(ws.merged_cells)
    for merged_range in merged_ranges:
        ws.unmerge_cells(str(merged_range))
    
    # 필요시 새로운 병합 셀 생성 로직 추가
    # (현재는 기본 구조만 유지)

if __name__ == "__main__":
    # 테스트
    collect_and_adjust_tab8("실거래가격정보.xlsx", "output.xlsx")
