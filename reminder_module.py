import os
import pandas as pd
import schedule
import time
import threading
from flask import Flask, render_template, request, redirect
import pyttsx3


app = Flask(__name__)

# Path to the Excel sheet
EXCEL_FILE = 'reminders_new.xlsx'

def save_to_excel(data):
    """Save new data to the Excel sheet."""
    file_exists = os.path.isfile(EXCEL_FILE)
    
    if file_exists:
        df = pd.read_excel(EXCEL_FILE)
    else:
        df = pd.DataFrame(columns=["Type", "Name", "Medication", "Reminder Time", "Appointment Time", "Hospital Name", "Hospital Address"])

    new_data = pd.DataFrame([data])
    df = pd.concat([df, new_data], ignore_index=True)
    df.to_excel(EXCEL_FILE, index=False)

def send_medication_reminder(name, medication):
    message = f"Hello {name}, it's time to take your medication: {medication}."
    speak_text(message)

def send_doctor_appointment_reminder(name, hospital_name, hospital_address):
    message = f"Hello {name}, you have a doctor's appointment at {hospital_name}, located at {hospital_address}."
    speak_text(message)

def speak_text(text):
    engine = pyttsx3.init()
    engine.say(text)
    engine.runAndWait()
    engine.stop()

def schedule_reminders():
    schedule.clear()

    if not os.path.isfile(EXCEL_FILE):
        return

    df = pd.read_excel(EXCEL_FILE)

    for _, row in df.iterrows():
        if row["Type"] == "medication":
            schedule.every().day.at(row['Reminder Time']).do(
                send_medication_reminder, row['Name'], row['Medication'])
        elif row["Type"] == "doctor_appointment":
            schedule.every().day.at(row['Appointment Time']).do(
                send_doctor_appointment_reminder, row['Name'], row['Hospital Name'], row['Hospital Address'])

def run_scheduler():
    while True:
        schedule.run_pending()
        time.sleep(1)

@app.route('/')
def home():
    return render_template('index.html') # Home page with buttons

@app.route('/scheduler')
def scheduler():
    return render_template('scheduler.html')

@app.route('/ai_assistant')
def ai_assistant():
    return render_template('ai_assistant.html') # Placeholder for AI Assistant page

@app.route('/submit', methods=['POST'])
def submit():
    reminder_type = request.form['reminder_type']

    if reminder_type == "medication":
        data = {
            "Type": "medication",
            "Name": request.form['name'],
            "Medication": request.form['medication'],
            "Reminder Time": request.form['reminder_time'],
            "Appointment Time": None,
            "Hospital Name": None,
            "Hospital Address": None
        }
    elif reminder_type == "doctor_appointment":
        data = {
            "Type": "doctor_appointment",
            "Name": request.form['name'],
            "Medication": None,
            "Reminder Time": None,
            "Appointment Time": request.form['appointment_time'],
            "Hospital Name": request.form['hospital_name'],
            "Hospital Address": request.form['hospital_address']
        }

    save_to_excel(data)
    schedule_reminders()

    return redirect('/scheduler')

if __name__ == '__main__':
    schedule_reminders()
    scheduler_thread = threading.Thread(target=run_scheduler, daemon=True)
    scheduler_thread.start()

    # Run the Flask app
    app.run(debug=True)