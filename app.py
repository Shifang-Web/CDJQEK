import streamlit as st
import pandas as pd
import io
import datetime
from sqlalchemy import create_engine, text

# ================= 页面配置 =================
st.set_page_config(page_title="每日检查表自动汇总系统", layout="wide")

# ================= 初始化数据库引擎 =================
@st.cache_resource
def init_db():
    # 从 Streamlit Secrets 中读取连接池 URL
    db_url = st.secrets["supabase"]["DATABASE_URL"]
    try:
        # 创建数据库引擎 (使用 psycopg2)
        engine = create_engine(db_url)
        # 测试连接
        with engine.connect() as conn:
            pass
        return engine
    except Exception as e:
        st.error(f"❌ 数据库连接失败，请检查 Secrets 配置。\n错误信息：{e}")
        st.stop()

engine = init_db()

# ================= 密码保护模块 =================
def check_password():
    def password_entered():
        if st.session_state["password"] == "admin123":
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.text_input("🔐 请输入系统密码以继续：", type="password", on_change=password_entered, key="password")
        st.info("提示：默认密码为 admin123，建议在代码中修改。")
        return False
    elif not st.session_state["password_correct"]:
        st.text_input("🔐 密码错误，请重试：", type="password", on_change=password_entered, key="password")
        st.error("😕 密码不正确")
        return False
    else:
        return True

if not check_password():
    st.stop()

# ================= 数据获取与清理 =================
@st.cache_data(ttl=60)
def load_all_data():
    try:
        # 使用 pandas 读取数据库
        query = "SELECT * FROM checklist ORDER BY upload_time DESC"
        df = pd.read_sql(query, engine)
        return df
    except Exception as e:
        st.error(f"读取数据库失败：{e}")
        return pd.DataFrame()

def cleanup_old_data():
    try:
        thirty_days_ago = (datetime.datetime.now(datetime.timezone.utc) - datetime.timedelta(days=30)).isoformat()
        with engine.connect() as conn:
            # 注意：直接在 SQL 中使用参数化查询
            conn.execute(text("DELETE FROM checklist WHERE upload_time < :time"), {"time": thirty_days_ago})
            conn.commit()
    except Exception as e:
        pass # 忽略清理错误，不影响主流程

# ================= 主界面 =================
st.title("📋 每日检查表自动汇总系统")
st.markdown("""
**使用说明：**
1. 请确保上传的 Excel 文件**表头格式统一**（第一行是列名，不要有合并单元格）。
2. 系统会根据文件名自动识别检查类型（品控 / 巡查 / 新风）。
3. 支持一次拖拽上传多个文件。
""")

cleanup_old_data()

uploaded_files = st.file_uploader(
    "请选择或拖拽上传检查表文件（支持多选 .xlsx 或 .xls）",
    type=['xlsx', 'xls'],
    accept_multiple_files=True
)

if st.button("🚀 开始自动汇总", use_container_width=True):
    if not uploaded_files:
        st.warning("⚠️ 请先上传至少一个文件。")
    else:
        progress_bar = st.progress(0, text="正在处理文件...")
        records_to_insert = []
        
        for idx, file in enumerate(uploaded_files):
            try:
                df = pd.read_excel(file, engine='openpyxl')
                df.dropna(how='all', inplace=True)
                if not df.empty and df.iloc[0, 0] == df.columns[0]:
                    df = df.iloc[1:]
                
                filename = file.name.lower()
                if "品控" in filename or "问题记录" in filename:
                    check_type = "品控"
                elif "巡查" in filename or "记录仪" in filename:
                    check_type = "巡查"
                elif "新风" in filename or "新风升级" in filename:
                    check_type = "新风"
                else:
                    check_type = "其他"
                
                record = {
                    "filename": file.name,
                    "check_type": check_type,
                    "data": df.to_dict(orient='records') 
                }
                records_to_insert.append(record)
            except Exception as e:
                st.error(f"❌ {file.name} 处理失败：{e}")
            
            progress_bar.progress((idx + 1) / len(uploaded_files), text=f"已处理 {idx+1}/{len(uploaded_files)} 个文件")
        
        progress_bar.empty()
        
        if records_to_insert:
            try:
                # 使用 pandas 直接写入数据库
                df_to_insert = pd.DataFrame(records_to_insert)
                # 注意：data 是 JSONB 格式，直接写入
                df_to_insert.to_sql('checklist', engine, if_exists='append', index=False)
                st.success(f"✅ 成功上传！本次处理了 {len(records_to_insert)} 个文件。")
                st.cache_data.clear()
                st.rerun()
            except Exception as e:
                st.error(f"❌ 写入数据库失败：{e}")

# ================= 展示与下载 =================
all_data = load_all_data()

if not all_data.empty:
    st.divider()
    st.subheader("📥 历史汇总数据（云端保留最近30天）")
    st.write(f"当前数据库中累计有 {len(all_data)} 条文件记录。")
    
    display_df = all_data[['id', 'filename', 'check_type', 'upload_time']].copy()
    display_df.columns = ['ID', '文件名称', '检查类型', '上传时间']
    st.dataframe(display_df.head(15), use_container_width=True)
    
    st.markdown("**导出选项：**")
    col1, col2 = st.columns(2)
    
    with col1:
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            expanded_dfs = []
            for index, row in all_data.iterrows():
                if isinstance(row['data'], list):
                    temp_df = pd.DataFrame(row['data'])
                    temp_df['来源文件'] = row['filename']
                    temp_df['检查类型'] = row['check_type']
                    temp_df['上传时间'] = row['upload_time']
                    expanded_dfs.append(temp_df)
            
            if expanded_dfs:
                final_export_df = pd.concat(expanded_dfs, ignore_index=True)
                final_export_df.to_excel(writer, index=False, sheet_name='汇总数据')
            else:
                all_data.to_excel(writer, index=False, sheet_name='汇总数据')
                
        st.download_button(
            label="📥 下载全部数据为 Excel (.xlsx)",
            data=buffer.getvalue(),
            file_name=f"历史汇总_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
else:
    st.info("数据库中暂无数据，等待成员上传。")