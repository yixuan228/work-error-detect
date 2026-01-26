import requests

# 根URL，所有接口都基于此URL
BASE_URL = "http://znsw.zxkj2006.com"

# 登入接口返回 token
login_url = BASE_URL + "/feed/api/admin/appLogin"
login_params = {
    "password": "123456",
    "username": "漯河汇兴"
}

login_resp = requests.post(login_url, json=login_params)
login_json = login_resp.json()
token = login_json["data"].get("token")

HEADERS = {
    "AuthorizationF": f"Bearer {token}"
}

# 协议6.5 导出猪只饲喂统计列表
EXP_FEED_DATA_URL = BASE_URL + '/feed/api/api/v1/data/exportRecordInfo'
