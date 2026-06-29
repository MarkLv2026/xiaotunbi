import base64, json, requests, os

TOKEN = os.environ.get('GITHUB_TOKEN', 'ghp_请通过环境变量设置TOKEN')
REPO = 'MarkLv2026/xiaotunbi'
BRANCH = 'main'
FILE_PATH = 'app.py'
LOCAL_FILE = r'C:\Users\Gwell\WorkBuddy\2026-05-22-17-31-37\xiaotunbi\app.py'
COMMIT_MSG = 'fix: 全屏事件监听改为 capture 阶段，避免被 Streamlit 拦截'

headers = {
    'Authorization': f'token {TOKEN}',
    'Accept': 'application/vnd.github.v3+json'
}

# 1. Get current commit SHA
url = f'https://api.github.com/repos/{REPO}/git/ref/heads/{BRANCH}'
resp = requests.get(url, headers=headers)
resp.raise_for_status()
current_sha = resp.json()['object']['sha']
print(f'Current commit: {current_sha[:8]}')

# 2. Get current tree
url = f'https://api.github.com/repos/{REPO}/git/commits/{current_sha}'
resp = requests.get(url, headers=headers)
resp.raise_for_status()
tree_sha = resp.json()['tree']['sha']
print(f'Current tree: {tree_sha[:8]}')

# 3. Create blob
with open(LOCAL_FILE, 'rb') as f:
    content = f.read()
url = f'https://api.github.com/repos/{REPO}/git/blobs'
resp = requests.post(url, headers=headers, json={
    'content': base64.b64encode(content).decode(),
    'encoding': 'base64'
})
resp.raise_for_status()
blob_sha = resp.json()['sha']
print(f'Blob created: {blob_sha[:8]}')

# 4. Create new tree
url = f'https://api.github.com/repos/{REPO}/git/trees'
resp = requests.post(url, headers=headers, json={
    'base_tree': tree_sha,
    'tree': [{
        'path': FILE_PATH,
        'mode': '100644',
        'type': 'blob',
        'sha': blob_sha
    }]
})
resp.raise_for_status()
new_tree_sha = resp.json()['sha']
print(f'New tree: {new_tree_sha[:8]}')

# 5. Create commit
url = f'https://api.github.com/repos/{REPO}/git/commits'
resp = requests.post(url, headers=headers, json={
    'message': COMMIT_MSG,
    'tree': new_tree_sha,
    'parents': [current_sha]
})
resp.raise_for_status()
new_commit_sha = resp.json()['sha']
print(f'New commit: {new_commit_sha[:8]}')

# 6. Update ref
url = f'https://api.github.com/repos/{REPO}/git/refs/heads/{BRANCH}'
resp = requests.patch(url, headers=headers, json={
    'sha': new_commit_sha
})
resp.raise_for_status()
print('✅ Push successful!')
