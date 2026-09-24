import streamlit as st
import pandas as pd
import io
import datetime
from supabase import create_client, Client

# ================= 页面配置 =================
st.set_page_config(page_title="每日检查表自动汇总系统", layout="wide")

# ================= 初始化 Supabase =================
@st.cache_resource
def init_supabase() -> Client:
    # 从 Streamlit Secrets 中读取密钥
    url = st.secrets["supabase"]["SUPABASE_URL"]
    key = st.secrets["supabase"]["SUPABASE_KEY"]
    return create_client(url, key)

try:
    supabase = init_supabase()
except Exception as e:
    st.error(f"❌ 数据库连接失败，请检查 Streamlit 的 Secrets 配置。\n错误信息：{e}")
    st.stop()

# ================= 密码保护模块 =================
def check_password():
    def password_entered():
        # 密码可以在这里修改，默认是 admin123
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

# ================= 数据获取与清理函数 =================
# 缓存数据，60秒内多次刷新不会重复请求数据库
@st.cache_data(ttl=60)
def load_all_data():
    try:
        response = supabase.table('checklist').select('*').order('upload_time', desc=True).execute()
        if response.data:
            return pd.DataFrame(response.data)
    except Exception as e:
        st.error(f"读取数据库失败：{e}")
    return pd.DataFrame()

# 自动清理超过30天的旧数据（每次应用加载时执行一次，使用缓存避免频繁请求）
@st.cache_resource
def cleanup_old_data():
    try:
        thirty_days_ago = (datetime.datetime.now(datetime.timezone.utc) - datetime.timedelta(days=30)).isoformat()
        supabase.table('checklist').delete().lt('upload_time', thirty_days_ago).execute()
    except Exception:
        pass

# ================= 主界面 =================
st.title("📋 每日检查表自动汇总系统")
st.markdown("""
**使用说明：**
1. 请确保上传的 Excel 文件**表头格式统一**（第一行是列名，不要有合并单元格）。
2. 系统会根据文件名自动识别检查类型（品控 / 巡查 / 新风）。
3. 支持一次拖拽上传多个文件。
""")

# 触发数据清理（在页面加载时）
cleanup_old_data()

# ================= 文件上传组件 =================
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
                # 读取 Excel 数据
                df = pd.read_excel(file, engine='openpyxl')
                df.dropna(how='all', inplace=True)
                if not df.empty and df.iloc[0, 0] == df.columns[0]:
                    df = df.iloc[1:]
                
                # 提取检查类型
                filename = file.name.lower()
                if "品控" in filename or "问题记录" in filename:
                    check_type = "品控"
                elif "巡查" in filename or "记录仪" in filename:
                    check_type = "巡查"
                elif "新风" in filename or "新风升级" in filename:
                    check_type = "新风"
                else:
                    check_type = "其他"
                
                # 构造要插入数据库的 JSON 数据
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
                # 批量插入数据库
                supabase.table('checklist').insert(records_to_insert).execute()
                st.success(f"✅ 成功上传！本次处理了 {len(records_to_insert)} 个文件，数据已同步到云端数据库。")
                st.cache_data.clear() # 清除缓存，让数据立刻刷新
                st.rerun()
            except Exception as e:
                st.error(f"❌ 写入数据库失败：{e}")

# ================= 数据展示与下载区域 =================
all_data = load_all_data()

if not all_data.empty:
    st.divider()
    st.subheader("📥 历史汇总数据（云端保留最近30天）")
    st.write(f"当前数据库中累计有 {len(all_data)} 条文件记录。")
    
    # 展示简化的列表
    display_df = all_data[['id', 'filename', 'check_type', 'upload_time']].copy()
    display_df.columns = ['ID', '文件名称', '检查类型', '上传时间']
    st.dataframe(display_df.head(15), use_container_width=True)
    
    st.markdown("**导出选项：**")
    col1, col2 = st.columns(2)
    
    with col1:
        # 导出为 Excel
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            # 展开所有上传的 JSON 数据以便于阅读
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
                # 如果没有可展开的数据，导出原始记录
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