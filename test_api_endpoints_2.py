import json
import requests

with open("config.json", encoding="utf-8") as f:
    cfg = json.load(f)

api_cfg = cfg["api"]
base = api_cfg["url"].split("/tms-report/")[0]  # https://gw-express.metfone.com.kh

headers = {
    "Authorization": f"Bearer {api_cfg['bearer_token']}",
    "Referer": api_cfg.get("referer", "https://opsexpress.metfone.com.kh/"),
    "Accept": "application/json, text/plain, */*",
    "Content-Type": "application/json",
    "x-client-id": api_cfg.get("x_client_id", "TMS_ANDROID"),
}

endpoints_to_test = [
    (f"{base}/vtp-user/api/v1/departments/posts/BATA001", "GET", None),
    (f"{base}/vtp-user/api/v1/departments/posts/detail?code=BATA001", "GET", None),
    (f"{base}/vtp-user/api/v1/departments/posts/get-by-code?code=BATA001", "GET", None),
    (f"{base}/vtp-user/api/v1/departments/posts/detail", "POST", {"code": "BATA001"}),
    (f"{base}/vtp-user/api/v1/departments/posts/get-by-code", "POST", {"code": "BATA001"}),
]

for url, method, body in endpoints_to_test:
    print(f"\nTesting: {method} {url}")
    try:
        if method == "GET":
            resp = requests.get(url, headers=headers, timeout=10)
        else:
            resp = requests.post(url, headers=headers, json=body, timeout=10)
            
        print(f"Status: {resp.status_code}")
        if resp.status_code == 200:
            data = resp.json()
            data_str = json.dumps(data, indent=2, ensure_ascii=False)
            print(data_str[:600])
        else:
            print(resp.text[:200])
    except Exception as e:
        print(f"Error: {e}")
