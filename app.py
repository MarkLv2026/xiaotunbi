# -*- coding: utf-8 -*-
"""
极简诊断版本 - 用于确认Streamlit Cloud环境是否正常
"""
import streamlit as st

st.set_page_config(page_title="小豚BI - 诊断模式", layout="wide")

st.title("🐷 小豚BI - 诊断模式")
st.success("✅ 应用已成功启动！")
st.info("如果看到这个页面，说明Streamlit Cloud环境正常。")
st.markdown("---")
st.subheader("诊断信息")
st.write("- Streamlit版本正常")
st.write("- Python运行环境正常")
st.write("- 无外部依赖导入")

st.markdown("---")
st.warning("这是极简诊断版本。如果此页面能正常显示，说明问题出在原app.py的代码中。")
