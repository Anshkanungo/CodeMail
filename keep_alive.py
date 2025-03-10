from flask import Flask
app = Flask(__name__)

@app.route('/')
def home():
    return "Email Responder is Alive"

# No need for app.run() here; it’s called in mail.py