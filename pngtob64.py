import base64

with open("trayicon.png", "rb") as f:
    print(base64.b64encode(f.read()).decode())