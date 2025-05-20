from flask import Flask, request
import csv
import os

app = Flask(__name__)

@app.route('/submit_email', methods=['POST'])
def submit_email():
    email = request.form.get('email')
    if email:
        with open('emails.csv', mode='a', newline='') as file:
            writer = csv.writer(file)
            writer.writerow([email])
        return 'Email saved successfully!', 200
    return 'Email not provided', 400

if __name__ == '__main__':
    if not os.path.exists('emails.csv'):
        with open('emails.csv', mode='w') as file:
            writer = csv.writer(file)
            writer.writerow(['Email'])  # Write header
    app.run(debug=True)
