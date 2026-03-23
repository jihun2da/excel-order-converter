import streamlit as st
import pandas as pd
import re
from io import BytesIO

# 웹앱의 기본 설정 (제목, 레이아웃 넓게)
st.set_page_config(page_title="주문서 변환기", layout="wide")
st.title("📦 엑셀 주문서 사이즈 분리 및 수량 합산기")
st.write("원본 엑셀 파일을 업로드하면, 사이즈를 분리하고 샘플 양식에 맞게 변환해 줍니다.")

# 파일 업로드 버튼 생성
uploaded_file = st.file_uploader("여기에 엑셀 파일을 드래그하거나 클릭해서 업로드하세요.", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    try:
        # 업로드된 파일을 판다스(데이터프레임)로 읽기
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
        
        processed_data = [] # 변환된 데이터를 담을 빈 리스트
        
        # 엑셀 데이터를 한 줄씩 읽으면서 처리
        for index, row in df.iterrows():
            # 컬럼 이름이 조금씩 다를 수 있으므로, 인덱스(순서)나 이름으로 안전하게 가져오기
            order_date = str(row.get('주문일', row.iloc[0] if len(row) > 0 else ''))
            order_id = str(row.get('주문자ID(주문번호)', row.iloc[1] if len(row) > 1 else ''))
            order_name = str(row.get('주문자명', row.iloc[2] if len(row) > 2 else ''))
            customer_name = str(row.get('고객명', row.iloc[3] if len(row) > 3 else ''))
            
            # 원본 파일 헤더에 'l브랜드'라고 되어 있는 것을 반영
            brand = str(row.get('l브랜드', row.iloc[4] if len(row) > 4 else '')) 
            prod_name = str(row.get('상품명', row.iloc[5] if len(row) > 5 else ''))
            color = str(row.get('옵션', row.iloc[6] if len(row) > 6 else ''))
            qty_str = str(row.get('주문수량(장수)', row.iloc[7] if len(row) > 7 else ''))
            
            # 브랜드명과 상품명을 띄어쓰기로 합치기
            combined_prod = f"{brand} {prod_name}".strip()
            
            # 수량 컬럼에 데이터가 있고 괄호 "("가 있는 경우만 분리
            if pd.notna(qty_str) and "(" in qty_str:
                items = str(qty_str).split('-')
                
                for item in items:
                    item = item.strip()
                    # 정규식을 이용해 맨 마지막 괄호 안의 숫자(수량)와 그 앞의 텍스트(사이즈) 분리
                    # 예: XS(1~2y)(3) -> 그룹1: XS(1~2y), 그룹2: 3
                    match = re.search(r'(.*)\((\d+)\)$', item)
                    
                    if match:
                        size_name = match.group(1).strip()
                        qty = int(match.group(2))
                        
                        # 옵션 텍스트 생성
                        final_option = f"색상: {color} / 사이즈: {size_name}"
                        
                        # 처리된 데이터를 리스트에 추가
                        processed_data.append({
                            '주문일': order_date,
                            '주문자ID(주문번호)': order_id,
                            '주문자명': order_name,
                            '고객명': customer_name,
                            '브랜드/상품명': combined_prod,
                            '옵션': final_option,
                            '수량': qty,
                            '옵션가': 0
                        })
        
        # 리스트를 다시 데이터프레임(표 형식)으로 변환
        result_df = pd.DataFrame(processed_data)
        
        if not result_df.empty:
            # 완전히 동일한 항목(주문일~옵션까지)의 수량을 합산 (Groupby)
            result_df = result_df.groupby(
                ['주문일', '주문자ID(주문번호)', '주문자명', '고객명', '브랜드/상품명', '옵션', '옵션가'], 
                as_index=False
            )['수량'].sum()
            
            # 샘플 엑셀 파일과 동일한 컬럼 순서로 정렬
            result_df = result_df[['주문일', '주문자ID(주문번호)', '주문자명', '고객명', '브랜드/상품명', '옵션', '수량', '옵션가']]
            
            st.success("데이터 변환 및 중복 합산이 완료되었습니다! 아래 표를 확인해주세요.")
            
            # 웹 화면에 결과 표 보여주기
            st.dataframe(result_df, use_container_width=True)
            
            # 엑셀 파일로 다운로드 할 수 있는 기능
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                result_df.to_excel(writer, index=False, sheet_name='변환완료')
            processed_excel = output.getvalue()
            
            st.download_button(
                label="📥 변환된 엑셀 파일 다운로드",
                data=processed_excel,
                file_name="변환완료_샘플양식.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("추출할 데이터가 없거나 조건에 맞는 데이터가 없습니다.")
            
    except Exception as e:
        st.error(f"파일을 처리하는 중 오류가 발생했습니다: {e}")