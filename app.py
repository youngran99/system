import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import os

# ----------------- 🌟 설정 및 초기화 -----------------
st.set_page_config(page_title="연구비 관리 시스템", layout="wide")

# 세션 상태 초기화 (데이터 유지용)
if 'data_loaded' not in st.session_state:
    st.session_state.data_loaded = False

# ----------------- 🛠️ 핵심 로직 함수 (기존 로직 이식) -----------------

def process_data(df, df_b):
    """데이터 정제 및 지표 계산 로직"""
    # 컬럼명 공백 제거
    df.columns = [col.strip() for col in df.columns]
    
    # 사용액 숫자 변환
    if '사용액' in df.columns:
        df['사용액'] = df['사용액'].astype(str).str.replace(',', '').str.replace('원', '').str.strip()
        df['사용액'] = pd.to_numeric(df['사용액'], errors='coerce').fillna(0)
    
    # 예산 정보 정리
    df_b.columns = [col.strip() for col in df_b.columns]
    df_b['배정예산'] = pd.to_numeric(df_b['배정예산'].astype(str).str.replace(',', ''), errors='coerce').fillna(25000000)
    budgets = dict(zip(df_b['교원별'].astype(str).str.strip(), df_b['배정예산']))
    
    # 지표 계산
    metrics = {}
    df['교원별'] = df['교원별'].astype(str).str.strip()
    pool_total = 3600000000 # 전체 총 예산
    
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

def create_excel_download(df, metrics, target_name=None):
    """엑셀 파일 생성 (메모리 내 작업)"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        if target_name: # 특정 교수님 한 명만
            df_target = df[df['교원별'] == target_name].copy()
            # 푸터 추가 로직 (기존 make_footer 유사)
            m = metrics[target_name]
            footer = pd.DataFrame([
                {'적요': '총 배정 예산', '사용액': m['배정예산']},
                {'적요': '총 사용액 합계', '사용액': m['사용액']},
                {'적요': '현재 잔액', '사용액': m['잔액']}
            ])
            df_final = pd.concat([df_target, footer], ignore_index=True)
            df_final.to_excel(writer, index=False, sheet_name=target_name)
        else: # 전체 통합 보고서
            summary_df = pd.DataFrame.from_dict(metrics, orient='index').reset_index().rename(columns={'index': '교원별'})
            summary_df.to_excel(writer, index=False, sheet_name="전체요약")
            for name, m in metrics.items():
                df[df['교원별'] == name].to_excel(writer, index=False, sheet_name=name)
    return output.getvalue()

# ----------------- 🖥️ 웹 UI 구성 -----------------

st.sidebar.header("📁 데이터 업로드")
uploaded_file = st.sidebar.file_uploader("연구비 내역 엑셀 파일을 선택하세요", type=['xlsx'])

if uploaded_file:
    try:
        # 엑셀 로드
        raw_df = pd.read_excel(uploaded_file, sheet_name='지출내역')
        raw_budget = pd.read_excel(uploaded_file, sheet_name='예산관리')
        
        # 데이터 처리
        df, metrics = process_data(raw_df, raw_budget)
        st.session_state.data_loaded = True
        
        # 사이드바 메뉴
        menu = st.sidebar.radio("메뉴 선택", ["전체 현황", "교수별 상세 조회", "데이터 내보내기"])

        if menu == "전체 현황":
            st.title("📊 전체 예산 집행 현황")
            
            # 카드형 요약 정보
            col1, col2, col3 = st.columns(3)
            total_assigned = sum(m['배정예산'] for m in metrics.values())
            total_used = sum(m['사용액'] for m in metrics.values())
            
            col1.metric("총 배정 예산 합계", f"{total_assigned:,.0f}원")
            col2.metric("총 집행액", f"{total_used:,.0f}원", delta=f"-{total_used/total_assigned*100:.1f}%", delta_color="inverse")
            col3.metric("전체 잔액", f"{total_assigned - total_used:,.0f}원")

            # 메인 표
            m_df = pd.DataFrame.from_dict(metrics, orient='index')
            st.subheader("교원별 요약 표")
            st.dataframe(m_df.style.format("{:,.0f}"), use_container_width=True)

        elif menu == "교수별 상세 조회":
            st.title("🔍 교수별 상세 지출 내역")
            name = st.selectbox("조회할 교수님 성함을 선택하세요", list(metrics.keys()))
            
            m = metrics[name]
            st.info(f"📌 **{name}** | 배정: {m['배정예산']:,}원 | 잔액: **{m['잔액']:,}원** (집행률 {m['집행률(%)']}%)")
            
            detail_df = df[df['교원별'] == name].copy()
            st.write("지출 상세 리스트")
            st.dataframe(detail_df, use_container_width=True)

        elif menu == "데이터 내보내기":
            st.title("💾 보고서 다운로드")
            
            col_a, col_b = st.columns(2)
            
            with col_a:
                st.subheader("1. 전체 통합 보고서")
                all_excel = create_excel_download(df, metrics)
                st.download_button(
                    label="📥 전체 통합 엑셀 다운로드",
                    data=all_excel,
                    file_name=f"통합보고서_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.ms-excel"
                )
            
            with col_b:
                st.subheader("2. 특정 교수님 개별 파일")
                target = st.selectbox("파일을 생성할 교수님 선택", list(metrics.keys()))
                single_excel = create_excel_download(df, metrics, target_name=target)
                st.download_button(
                    label=f"📥 {target} 교수님 파일 다운로드",
                    data=single_excel,
                    file_name=f"연구비내역_{target}_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.ms-excel"
                )

    except Exception as e:
        st.error(f"파일을 처리하는 중 오류가 발생했습니다: {e}")
        st.info("시트 이름이 '지출내역'과 '예산관리'로 되어 있는지 확인해 주세요.")

else:
    st.title("🧪 연구비 관리 시스템")
    st.warning("왼쪽 사이드바에서 엑셀 파일을 업로드해 주세요.")
    st.info("""
    **💡 사용 방법:**
    1. 관리하시는 엑셀 파일을 준비합니다.
    2. '지출내역' 시트와 '예산관리' 시트가 포함되어 있어야 합니다.
    3. 파일을 업로드하면 실시간으로 잔액이 계산됩니다.
    """)
