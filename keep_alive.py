"""
Streamlit Cloud Keep-Alive Script
用于防止 Streamlit Cloud 应用因休眠而无法访问

使用方法：
1. 部署到任意可定时访问的服务（如 UptimeRobot、自有服务器 cron）
2. 或者手动运行: python keep_alive.py
"""
import urllib.request
import sys

URL = "https://xiaotunbi-tmfhdkek237cxntwknq6ny.streamlit.app/"

def wake():
    try:
        req = urllib.request.Request(URL, headers={'User-Agent': 'Mozilla/5.0'})
        resp = urllib.request.urlopen(req, timeout=30)
        html = resp.read().decode('utf-8', errors='ignore')
        if "gone to sleep" in html or "Zzzz" in html:
            print("⚠️ App is sleeping, wake-up button detected")
            return False
        print(f"✅ App is awake (HTTP {resp.status})")
        return True
    except Exception as e:
        print(f"❌ Error: {e}")
        return False

if __name__ == "__main__":
    ok = wake()
    sys.exit(0 if ok else 1)
