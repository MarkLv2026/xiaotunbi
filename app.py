# -*- coding: utf-8 -*-
"""
诊断版本 - 捕获所有异常显示错误信息
"""
import traceback
import sys

try:
    import streamlit as st
    st.set_page_config(page_title="诊断", layout="wide")
    st.success("✅ Streamlit导入成功")
    st.write(f"Python版本: {sys.version}")
    
    # 测试基本功能
    import datetime
    st.write(f"datetime导入成功")
    
    import pathlib
    st.write(f"pathlib导入成功")
    
    # 测试可选依赖
    try:
        import pandas as pd
        st.write(f"pandas {pd.__version__} 导入成功")
    except Exception as e:
        st.warning(f"pandas导入失败: {e}")
    
    try:
        import plotly
        st.write(f"plotly {plotly.__version__} 导入成功")
    except Exception as e:
        st.warning(f"plotly导入失败: {e}")
    
    try:
        import openpyxl
        st.write(f"openpyxl {openpyxl.__version__} 导入成功")
    except Exception as e:
        st.warning(f"openpyxl导入失败: {e}")
    
    try:
        from dashboard_core import parse_sales_workbook
        st.write(f"dashboard_core导入成功")
    except Exception as e:
        st.warning(f"dashboard_core导入失败: {e}")
    
    st.balloons()
    st.info("所有测试完成！如果看到这个页面，说明环境正常。")
    
except Exception as e:
    # 如果Streamlit都导入失败，打印到stderr
    print(f"FATAL ERROR: {e}", file=sys.stderr)
    print(traceback.format_exc(), file=sys.stderr)
    # 尝试用纯文本显示
    error_msg = f"错误: {e}\n\n{traceback.format_exc()}"
    try:
        import streamlit as st
        st.error(error_msg)
    except:
        print(error_msg)
