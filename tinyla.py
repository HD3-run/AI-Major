import requests

url = "http://127.0.0.1:11434/api/generate"
payload = {
    "model": "tinyllama",
    "prompt": "Tell me a joke.",
    "options": {"max_tokens": 50}
}

try:
    res = requests.post(url, json=payload)
    print("Status Code:", res.status_code)
    print("Response:", res.text)
except Exception as e:
    print("Failed to connect:", e)
