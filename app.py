import streamlit as st
import pandas as pd
import io
import datetime
import re

# ================= 页面配置 =================
st.set_page_config(
    page_title="每日检查表自动汇总系统",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ================= 密码保护模块 =================
def check_password():
    """简单的密码验证，防止无关人员访问"""
    def password_entered():
        # 这里的密码修改为你自己的密码，例如 "admin123"
        if st.session_state["password"] == "admin123":
            st.session_state["password_correct"] = True
            del st.session_state["password"]  # 验证成功后清除密码
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.text_input(
            "请输入系统密码以继续：", 
            type="password", 
            on_change=password_entered, 
            key="password"
        )
        st.info("提示：默认密码为 admin123，建议在代码中修改。")
        return False
    elif not st.session_state["password_correct"]:
        st.text_input(
            "密码错误，请重试：", 
            type="password", 
            on_change=password_entered, 
            key="password"
        )
        st.error("😕 密码不正确")
        return False
    else:
        # 密码正确
        return True

if not check_password():
    st.stop()  # 密码验证不通过，停止执行后面的代码

# ================= 初始化 Session State =================
if 'master_df' not in st.session_state:
    st.session_state.master_df = pd.DataFrame()

# ================= 主界面 =================
st.title("📋 每日检查表自动汇总系统")
st.markdown("""
**使用说明：**
1. 请确保上传的 Excel 文件**表头格式统一**（第一行是列名，不要有合并单元格）。
2. 系统会根据文件名自动识别检查类型（品控 / 巡查 / 新风）。
3. 支持一次拖拽上传多个文件。
""")

# ================= 文件上传组件 =================
uploaded_files = st.file_uploader(
    "请选择或拖拽上传检查表文件（支持多选 .xlsx 或 .xls）",
    type=['xlsx', 'xls'],
    accept_multiple_files=True
)

# ================= 汇总处理逻辑 =================
if st.button("🚀 开始自动汇总", use_container_width=True):
    if not uploaded_files:
        st.warning("⚠️ 请先上传至少一个文件。")
    else:
        progress_bar = st.progress(0, text="正在处理文件...")
        new_data_list = []
        error_files = []
        
        for idx, file in enumerate(uploaded_files):
            try:
                # 1. 读取 Excel
                df = pd.read_excel(file, engine='openpyxl')
                
                # 2. 数据清洗
                # 去除完全空白的行
                df.dropna(how='all', inplace=True)
                # 去除可能存在的重复表头行
                if not df.empty and df.iloc[0, 0] == df.columns[0]:
                    df = df.iloc[1:]
                
                # 3. 提取检查类型（根据文件名关键词）
                filename = file.name.lower()
                if "品控" in filename or "问题记录" in filename:
                    check_type = "品控"
                elif "巡查" in filename or "记录仪" in filename:
                    check_type = "巡查"
                elif "新风" in filename or "新风升级" in filename:
                    check_type = "新风"
                else:
                    check_type = "其他"
                
                # 4. 添加标记字段
                df['检查类别'] = check_type
                df['来源文件'] = file.name
                df['上传时间'] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                
                new_data_list.append(df)
            except Exception as e:
                error_files.append(f"{file.name}: {str(e)}")
            
            # 更新进度条
            progress_bar.progress((idx + 1) / len(uploaded_files), text=f"已处理 {idx+1}/{len(uploaded_files)} 个文件")
        
        progress_bar.empty()
        
        # 处理错误信息
        if error_files:
            st.error("以下文件处理失败，请检查格式：\n" + "\n".join(error_files))
        
        # 合并数据
        if new_data_list:
            new_combined = pd.concat(new_data_list, ignore_index=True)
            
            if st.session_state.master_df.empty:
                st.session_state.master_df = new_combined
            else:
                st.session_state.master_df = pd.concat(
                    [st.session_state.master_df, new_combined], 
                    ignore_index=True
                )
            
            st.success(f"✅ 成功汇总！本次新增 {len(new_combined)} 条记录，当前系统累计 {len(st.session_state.master_df)} 条记录。")
            st.dataframe(new_combined.head(10))

# ================= 下载区域 =================
if not st.session_state.master_df.empty:
    st.divider()
    st.subheader("📥 下载汇总结果")
    
    col1, col2 = st.columns(2)
    
    with col1:
        # 下载 Excel
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            st.session_state.master_df.to_excel(writer, index=False, sheet_name='汇总结果')
        st.download_button(
            label="下载为 Excel 文件 (.xlsx)",
            data=buffer.getvalue(),
            file_name=f"检查汇总_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    
    with col2:
        # 下载 CSV
        csv_data = st.session_state.master_df.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            label="下载为 CSV 文件 (.csv)",
            data=csv_data,
            file_name=f"检查汇总_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            mime="text/csv",
            use_container_width=True
        )
    
    # 清空按钮
    if st.button("🗑️ 清空当前汇总数据，重新开始"):
        st.session_state.master_df = pd.DataFrame()
        st.rerun()