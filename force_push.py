#!/usr/bin/env python3
"""
通过 GitHub API 强制推送本地提交到远程仓库
处理 data/targets.xlsx (12.6MB) 大文件
"""

import base64
import json
import os
import sys
import urllib.request
import urllib.error

# ⚠️ 禁用 SSL 验证（仅用于开发环境）
import ssl
ssl._create_default_https_context = ssl._create_unverified_context

# GitHub 配置
TOKEN = os.environ.get("GITHUB_TOKEN", "ghp_请通过环境变量设置TOKEN")
REPO_OWNER = "MarkLv2026"
REPO_NAME = "xiaotunbi"
BRANCH = "main"

# 本地提交信息（每次推送前请更新为最新提交）
LOCAL_COMMIT_SHA = "e4234470205dcf185e46dbf96e4b48c775c03e14"
LOCAL_TREE_SHA = "4b6a9b311534d10bf08adb0a569bb02eb7cd10eb"
LOCAL_PARENT_SHA = "877371e8dda259b8b941f7c0f4408897befee754"

API_BASE = f"https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}"

def api_request(method, url, data=None, accept="application/vnd.github.v3+json"):
    """发送 GitHub API 请求"""
    headers = {
        "Authorization": f"token {TOKEN}",
        "Accept": accept,
        "User-Agent": "ForcePushScript/1.0"
    }

    if data is not None:
        data = json.dumps(data).encode('utf-8')
        headers["Content-Type"] = "application/json"

    req = urllib.request.Request(url, data=data, headers=headers, method=method)

    try:
        with urllib.request.urlopen(req, timeout=300) as response:
            return json.loads(response.read().decode('utf-8'))
    except urllib.error.HTTPError as e:
        error_body = e.read().decode('utf-8')
        print(f"❌ HTTP Error {e.code}: {error_body}")
        raise

def get_remote_head():
    """获取远程分支 HEAD"""
    url = f"{API_BASE}/git/ref/heads/{BRANCH}"
    try:
        result = api_request("GET", url)
        return result["object"]["sha"]
    except Exception as e:
        print(f"⚠️ 无法获取远程 HEAD: {e}")
        return None

def create_blob(content, encoding="base64"):
    """创建 blob"""
    url = f"{API_BASE}/git/blobs"
    data = {
        "content": content,
        "encoding": encoding
    }
    result = api_request("POST", url, data)
    return result["sha"]

def create_tree(tree_items):
    """创建 tree"""
    url = f"{API_BASE}/git/trees"
    data = {
        "tree": tree_items
    }
    result = api_request("POST", url, data)
    return result["sha"]

def create_commit(tree_sha, parent_sha, message):
    """创建 commit"""
    url = f"{API_BASE}/git/commits"
    data = {
        "tree": tree_sha,
        "parents": [parent_sha],
        "message": message
    }
    result = api_request("POST", url, data)
    return result["sha"]

def update_ref(ref, commit_sha, force=True):
    """更新引用（强制推送）"""
    url = f"{API_BASE}/git/refs/heads/{BRANCH}"
    data = {
        "sha": commit_sha,
        "force": force
    }
    result = api_request("PATCH", url, data)
    return result

def get_commit_tree(commit_sha):
    """获取提交的 tree SHA"""
    url = f"{API_BASE}/git/commits/{commit_sha}"
    result = api_request("GET", url)
    return result["tree"]["sha"]

def main():
    print("=" * 60)
    print("🚀 开始强制推送本地更改到 GitHub")
    print("=" * 60)

    # 1. 获取远程 HEAD
    print(f"\n1️⃣ 获取远程 {BRANCH} 分支 HEAD...")
    remote_head = get_remote_head()
    if remote_head:
        print(f"   远程 HEAD: {remote_head[:8]}")
    else:
        print(f"   ⚠️ 无法获取远程 HEAD，将使用本地 parent")
        remote_head = LOCAL_PARENT_SHA

    # 2. 读取本地提交信息
    print(f"\n2️⃣ 读取本地提交 {LOCAL_COMMIT_SHA[:8]}...")
    import subprocess

    # 获取提交信息
    result = subprocess.run(
        ["git", "cat-file", "-p", LOCAL_COMMIT_SHA],
        capture_output=True,
        text=True,
        cwd=os.path.dirname(os.path.abspath(__file__))
    )
    commit_content = result.stdout
    print(f"   提交信息:")
    for line in commit_content.split('\n')[:10]:
        print(f"   {line}")
    if len(commit_content.split('\n')) > 10:
        print(f"   ...")

    # 3. 获取提交的 tree
    print(f"\n3️⃣ 获取提交 tree {LOCAL_TREE_SHA[:8]}...")
    result = subprocess.run(
        ["git", "ls-tree", "-r", LOCAL_TREE_SHA],
        capture_output=True,
        text=True,
        cwd=os.path.dirname(os.path.abspath(__file__))
    )
    tree_items = []
    for line in result.stdout.strip().split('\n'):
        if line:
            parts = line.split()
            mode = parts[0]
            obj_type = parts[1]
            sha = parts[2]
            path = ' '.join(parts[3:])
            tree_items.append({
                "mode": mode,
                "type": obj_type,
                "sha": sha,
                "path": path
            })

    print(f"   Tree 包含 {len(tree_items)} 个文件/目录")
    for item in tree_items:
        print(f"   - {item['path']} ({item['sha'][:8]})")

    # 4. 为所有文件创建 blob（如果不存在于远程）
    print(f"\n4️⃣ 检查并创建 blob...")
    new_tree_items = []

    for item in tree_items:
        if item['type'] == 'blob':
            # 读取文件内容
            file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), item['path'])
            if os.path.exists(file_path):
                with open(file_path, 'rb') as f:
                    content = f.read()

                # 检查文件大小
                if len(content) > 100 * 1024 * 1024:  # 100MB
                    print(f"   ⚠️ 文件 {item['path']} 太大 ({len(content)} bytes)，跳过")
                    new_tree_items.append(item)
                    continue

                # Base64 编码
                content_b64 = base64.b64encode(content).decode('utf-8')

                print(f"   创建 blob: {item['path']} ({len(content)} bytes)...")
                try:
                    blob_sha = create_blob(content_b64, "base64")
                    print(f"   ✅ Blob 创建成功: {blob_sha[:8]}")
                    new_tree_items.append({
                        "mode": item['mode'],
                        "type": item['type'],
                        "sha": blob_sha,
                        "path": item['path']
                    })
                except Exception as e:
                    print(f"   ❌ Blob 创建失败: {e}")
                    # 使用原始 SHA（可能已存在于远程）
                    new_tree_items.append(item)
            else:
                print(f"   ⚠️ 文件不存在: {item['path']}")
                new_tree_items.append(item)
        else:
            # 目录，直接使用原始 SHA
            new_tree_items.append(item)

    # 5. 创建新的 tree
    print(f"\n5️⃣ 创建新的 tree...")
    new_tree_sha = create_tree(new_tree_items)
    print(f"   ✅ Tree 创建成功: {new_tree_sha[:8]}")

    # 6. 创建新的 commit（使用远程 HEAD 作为 parent）
    print(f"\n6️⃣ 创建新的 commit...")
    commit_message = "feat: 上半年汇总改为华为三GAP方法论 + 单月详情增加达成指标图"
    
    # 使用远程 HEAD 作为 parent（确保在远程仓库中存在）
    parent_sha = remote_head if remote_head else LOCAL_PARENT_SHA
    print(f"   使用 parent: {parent_sha[:8]}")
    
    new_commit_sha = create_commit(new_tree_sha, parent_sha, commit_message)
    print(f"   ✅ Commit 创建成功: {new_commit_sha[:8]}")

    # 7. 强制更新引用
    print(f"\n7️⃣ 强制更新 {BRANCH} 分支引用...")
    result = update_ref(f"heads/{BRANCH}", new_commit_sha, force=True)
    print(f"   ✅ 引用更新成功!")
    print(f"   URL: {result.get('url', 'N/A')}")
    print(f"   Object SHA: {result.get('object', {}).get('sha', 'N/A')[:8]}")

    print("\n" + "=" * 60)
    print("🎉 强制推送完成!")
    print("=" * 60)
    print(f"\n📌 请等待 Streamlit Cloud 自动部署（约 2-3 分钟）")
    print(f"🔗 查看更改: https://github.com/{REPO_OWNER}/{REPO_NAME}/commit/{new_commit_sha}")

if __name__ == "__main__":
    main()
