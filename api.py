import requests

response = requests.get("http://<192.168.1.10>:8000/catalogos")
print(response.json())
