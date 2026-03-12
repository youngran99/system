import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import zipfile

# ----------------- 🌟 1. 웹 페이지 기본 설정 -----------------
st.set_page_config(page_title="연구비 관리 시스템", layout="wide")

# ----------------- 🛠️ 2. 데이터 처리 핵심 로직 -----------------

def process_data(df, df_b):
    """지출 내역과 예산 정보를 결합하여 계산"""
    # 컬럼 공백 제거
    df.columns = [col.strip() for col in df.columns]
    df_b.columns = [col.strip() for col in df_b.columns]
    
    # 사용액 숫자 변환
    if '사용액' in df.columns:
        df['사용액'] = df['사용액'].astype(str).str.replace(',', '').str.replace('원', '').str.strip()
        df['사용액'] = pd.to_numeric(df['사용액'], errors='coerce').fillna(0)
    
    # 예산 정보 딕셔너리 생성 (기본값 25,000,000원)
    df_b['배정예산'] = pd.to_numeric(df_b['배정예산'].astype(str).str.replace(',', ''), errors='coerce').fillna(25000000)
    budgets = dict(zip(df_b['교원별'].astype(str).str.strip(), df_b['배정예산']))
    
    # 교수별 지표 계산
    metrics = {}
    df['교원별'] = df['교원별'].astype(str).str.strip()
    
    for n in df['교원별'].unique():
        if n in ['nan', 'None', '']: continue
        b = int(budgets.get(n, 25000000))
        e = int(df[df['교원별'] == n]['사용액'].sum())
        metrics[n] = {
            '배정예산': b, 
            '사용액': e, 
            '잔액': b - e, 
            '집행률(%)': round((e/b*100 if b>0 else 0), 1)
        }
    return df, metrics

# ----------------- 💾 3. 엑셀 및 ZIP 생성 함수 -----------------

def create_excel_download(df, metrics, target_name=None):
    """
    target_name이 있으면: 해당 교수님의 단일 시트 파일 생성
    target_name이 없으면: 전체 요약 + 교수별 시트가 포함된 통합 파일 생성
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        if target_name:
            # [개별 파일용]
            df_target = df[df['교원별'] == target_name].copy()
            m = metrics[target_name]
            footer = pd.DataFrame([
                {'적요': '---', '사용액': 0}, # 구분선 역할
                {'적요': '총 배정 예산', '사용액': m['배정예산']},
                {'적요': '총 사용액 합계', '사용액': m['사용액']},
                {'적요': '현재 잔액', '사용액': m['잔액']}
            ])
            pd.concat([df_target, footer], ignore_index=True).to_excel(writer, index=False, sheet_name="지출내역")
        else:
            # [통합 파일용]
            summary_df = pd.DataFrame.from_dict(metrics, orient='index').reset_index().rename(columns={'index': '교원별'})
            summary_df.to_excel(writer, index=False, sheet_name="전체요약")
            for name in metrics.keys():
                df[df['교원별'] == name].to_excel(writer, index=False, sheet_name=name[:31])
    return output.getvalue()

def create_zip_download(df, metrics):
    """모든 교원의 개별 엑셀 파일을 하나의 ZIP으로 압축"""
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        for name in metrics.keys():
            excel_data = create_excel_download(df, metrics, target_name=name)
            file_name = f"연구비내역_{name}_{datetime.now().strftime('%Y%m%d')}.xlsx"
            zip_file.writestr(file_name, excel_data)
    return zip_buffer.getvalue()

# ----------------- 🖥️ 4. 웹 UI 구성 -----------------

st.sidebar.header("📁 데이터 업로드")
uploaded_file = st.sidebar.file_uploader("연구비 지출내역 엑셀 파일을 업로드하세요", type=['xlsx'])

if uploaded_file:
    try:
        # 데이터 로드
        raw_df = pd.read_excel(uploaded_file, sheet_name='지출내역')
        raw_budget = pd.read_excel(uploaded_file, sheet_name='예산관리')
        
        df, metrics = process_data(raw_df, raw_budget)
        
        # 메뉴 선택
        menu = st.sidebar.radio("메뉴 선택", ["📊 전체 현황", "🔍 교수별 상세 조회", "💾 데이터 내보내기"])

        if menu == "📊 전체 현황":
            st.title("📊 연구비 집행 전체 현황")
            
            # 상단 요약 지표
            total_b = sum(m['배정예산'] for m in metrics.values())
            total_e = sum(m['사용액'] for m in metrics.values())
            
            c1, c2, c3 = st.columns(3)
            c1.metric("총 배정 예산", f"{total_b:,.0f}원")
            c2.metric("총 사용액", f"{total_e:,.0f}원", f"-{total_e/total_b*100:.1f}%", delta_color="inverse")
            c3.metric("남은 잔액", f"{total_b - total_e:,.0f}원")

            st.divider()
            m_df = pd.DataFrame.from_dict(metrics, orient='index')
            st.subheader("교원별 요약 리스트")
            st.dataframe(m_df.style.format("{:,.0f}"), use_container_width=True)

        elif menu == "🔍 교수별 상세 조회":
            st.title("🔍 교수별 상세 조회")
            target = st.selectbox("교수님 선택", list(metrics.keys()))
            
            m = metrics[target]
            st.success(f"**{target}** 교수님 현황: 잔액 **{m['잔액']:,}원** (집행률 {m['집행률(%)']}%)")
            
            st.dataframe(df[df['교원별'] == target], use_container_width=True)

        elif menu == "💾 데이터 내보내기":
            st.title("💾 보고서 다운로드")
            
            col_left, col_right = st.columns(2)
            
            with col_left:
                st.subheader("📦 일괄 다운로드")
                # 통합 엑셀 다운로드
                st.download_button(
                    label="📥 통합 엑셀 (한 파일에 모든 시트)",
                    data=create_excel_download(df, metrics),
                    file_name=f"통합보고서_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.ms-excel",
                    use_container_width=True
                )
                # ZIP 다운로드
                st.download_button(
                    label="🗂️ 개별 엑셀 압축파일 (ZIP)",
                    data=create_zip_download(df, metrics),
                    file_name=f"연구비_개별파일모음_{datetime.now().strftime('%Y%m%d')}.zip",
                    mime="application/zip",
                    use_container_width=True
                )

            with col_right:
                st.subheader("👤 개별 다운로드")
                sel_name = st.selectbox("추출할 교수님 선택", list(metrics.keys()))
                st.download_button(
                    label=f"📥 {sel_name} 교수님 파일만 저장",
                    data=create_excel_download(df, metrics, target_name=sel_name),
                    file_name=f"연구비내역_{sel_name}.xlsx",
                    mime="application/vnd.ms-excel",
                    use_container_width=True
                )

    except Exception as e:
        st.error(f"파일을 처리하는 중 오류가 발생했습니다: {e}")
        st.info("엑셀 시트 이름이 '지출내역'과 '예산관리'인지 확인해 주세요.")

else:
    st.title("🧪 å연구비 관리 시스템")
    st.warning("왼쪽 사이드바에서 엑셀 파일을 업로드해 주세요.")
    st.info("""
    **💡 사용 방법:**
    1. 관리하시는 엑셀 파일을 준비합니다.
    2. '지출내역' 시트와 '예산관리' 시트가 포함되어 있어야 합니다.
    3. 파일을 업로드하면 실시간으로 잔액이 계산됩니다.
    """)
