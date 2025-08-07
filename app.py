from flask import Flask, render_template, request
import json
import os
import openpyxl
import datetime
from fuzzywuzzy import fuzz

app = Flask(__name__)

with open("qa_data_kef_final_with_slang.json", "r", encoding="utf-8") as f:
    qa_data = json.load(f)

def match_question_fuzzy(user_question):
    user_question = user_question.strip().lower()
    best_score = 0
    best_answer = qa_data.get("default", "Sorry, I couldn’t understand that.")

    for question, answer in qa_data.items():
        score = fuzz.partial_ratio(user_question, question.lower())
        if score > best_score and score >= 60:
            best_score = score
            best_answer = answer

    return best_answer

def log_to_excel(user_msg, bot_response):
    file_name = "chat_log.xlsx"
    if not os.path.exists(file_name):
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "Chat Log"
        sheet.append(["Timestamp", "User", "Bot"])
        workbook.save(file_name)

    workbook = openpyxl.load_workbook(file_name)
    sheet = workbook.active
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    sheet.append([timestamp, user_msg, bot_response])
    workbook.save(file_name)

@app.route("/")
def home():
    return render_template("index.html")

@app.route("/get")
def get_bot_response():
    user_msg = request.args.get("msg")
    bot_response = match_question_fuzzy(user_msg)
    log_to_excel(user_msg, bot_response)
    return bot_response

PORT = int(os.environ.get("PORT", 10000))
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=PORT)
