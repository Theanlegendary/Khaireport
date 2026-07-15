import requests
import json

def main():
    config_path = r"c:\Users\DELL\Downloads\Telegram Desktop\DataPusher 2\DataPusher\push_bot\config.json"
    with open(config_path, encoding="utf-8") as f:
        cfg = json.load(f)
    token = cfg["api"]["bearer_token"]
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json, text/plain, */*",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
    }
    
    order_id = "3003488041"
    out_data = {}
    
    # 1. Search API
    search_url = "https://gw-express.metfone.com.kh/tms-receiving/api/v1/orders/search"
    r_search = requests.get(search_url, params={"order_code": order_id}, headers=headers, timeout=15)
    out_data["search_status_code"] = r_search.status_code
    try:
        out_data["search_data"] = r_search.json()
    except Exception as e:
        out_data["search_error"] = str(e)
        out_data["search_text"] = r_search.text[:2000]
        
    # 2. Tracking API
    track_url = "https://gw-express.metfone.com.kh/tms-tracking/api/v1/order-tracking"
    r_track = requests.get(track_url, params={"order_id": order_id}, headers=headers, timeout=15)
    out_data["track_status_code"] = r_track.status_code
    try:
        out_data["track_data"] = r_track.json()
    except Exception as e:
        out_data["track_error"] = str(e)
        out_data["track_text"] = r_track.text[:2000]

    with open("order_output.json", "w", encoding="utf-8") as f:
        json.dump(out_data, f, indent=2, ensure_ascii=False)
    print("Saved to order_output.json")

if __name__ == "__main__":
    main()
