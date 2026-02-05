from app import app  # 这里 app 是你 app.py 里 Flask(__name__) 创建的那个变量
from flask import send_from_directory

@app.route("/")
def index():
    return send_from_directory("public", "index.html")
