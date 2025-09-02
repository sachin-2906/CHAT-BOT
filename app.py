from flask import Flask, render_template, request
import json
import os
import openpyxl
import datetime
from fuzzywuzzy import fuzz

app = Flask(__name__)

# ----------------------------
# Set base directory for files
# ----------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_PATH = os.path.join(BASE_DIR, "qa_data_kef_final_with_slang.json")
LOG_PATH = os.path.join(BASE_DIR, "chat_log.xlsx")

# ----------------------------
# Load the nested FAQ JSON with categories
# ----------------------------
try:
    with open(JSON_PATH, "r", encoding="utf-8") as f:
        qa_data = json.load(f)
except FileNotFoundError:
    qa_data = {}
    print(f"⚠️ JSON file not found at {JSON_PATH}")

# ----------------------------
# Function: Fuzzy Match Question
# ----------------------------
def match_question_fuzzy(user_question):
    user_question = user_question.strip().lower()
    best_score = 0
    best_answer = "Sorry, I couldn’t understand that."  # Default fallback

    for category, qas in qa_data.items():
        if isinstance(qas, dict):
            for question, answer in qas.items():
                score = fuzz.partial_ratio(user_question, question.lower())
                if score > best_score and score >= 60:
                    best_score = score
                    best_answer = answer
    return best_answer

# ----------------------------
# Function: Log conversation to Excel
# ----------------------------
def log_to_excel(user_msg, bot_response, teacher_input=""):
    if not os.path.exists(LOG_PATH):
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "Chat Log"
        sheet.append(["Timestamp", "User Message", "Bot Response", "Teacher Input"])
        workbook.save(LOG_PATH)

    workbook = openpyxl.load_workbook(LOG_PATH)
    sheet = workbook.active
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    sheet.append([timestamp, user_msg, bot_response, teacher_input])
    workbook.save(LOG_PATH)

# ----------------------------
# Routes
# ----------------------------
@app.route("/")
def home():
    return render_template("index.html")  # index.html must be inside /templates/

@app.route("/get")
def get_bot_response():
    user_msg = request.args.get("msg", "").strip()
    teacher_input = request.args.get("teacher_input", "").strip()

    if not user_msg and not teacher_input:
        return "Please select a question or enter teacher input."

    if user_msg:
        bot_response = match_question_fuzzy(user_msg)
    else:
        bot_response = "Your input has been noted. KEF will review it."

    log_to_excel(user_msg, bot_response, teacher_input)
    return bot_response

# ----------------------------
# Run app
# ----------------------------
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
