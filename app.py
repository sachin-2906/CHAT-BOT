from flask import Flask, render_template, request, jsonify
import json
import os
import openpyxl
import datetime

app = Flask(__name__)
app.debug = True

# Load JSON file with categorized FAQs
with open("kef_faq_chatbot_final.json", "r", encoding="utf-8") as f:
    faq_data = json.load(f)

# Excel logging function
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

# Home route
@app.route("/")
def home():
    return render_template("index.html", categories=list(faq_data.keys()))

# Get questions for a selected category
@app.route("/get_questions", methods=["POST"])
def get_questions():
    category = request.json.get("category")
    questions = list(faq_data.get(category, {}).keys())
    return jsonify(questions)

# Get answer for a selected question
@app.route("/get_answer", methods=["POST"])
def get_answer():
    category = request.json.get("category")
    question = request.json.get("question")
    answer = faq_data.get(category, {}).get(question, "Sorry, no answer found.")
    
    # Add "More Help" option at the end
    answer += "<br><br><b>Need more help?</b> <a href='https://kotakeducation.org/contact-us/' target='_blank'>Contact KEF Support</a>"

    log_to_excel(f"{category} → {question}", answer)
    return jsonify(answer)

# Run the app
PORT = int(os.environ.get("PORT", 10000))
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=PORT)
