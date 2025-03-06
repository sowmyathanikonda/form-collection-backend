from flask import Flask, request, jsonify
from flask_cors import CORS
import pymysql
import openpyxl
import os

app = Flask(__name__)
CORS(app)  # Enable CORS for all routes

# MySQL Configuration using environment variables
db_config = {
    'host': os.getenv('MYSQL_HOST', 'localhost'),  # Default to 'localhost' if not set
    'user': os.getenv('MYSQL_USER', 'root'),  # Default to 'root' if not set
    'password': os.getenv('MYSQL_PASSWORD', '12345'),  # Default to '12345' if not set
    'database': os.getenv('MYSQL_DATABASE', 'form_data'),  # Default to 'form_data' if not set
}

# Connect to MySQL
def get_db_connection():
    return pymysql.connect(**db_config)

# Create MySQL Table if it doesn't exist
def create_table():
    connection = get_db_connection()
    cursor = connection.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS form_submissions (
            id INT AUTO_INCREMENT PRIMARY KEY,
            name VARCHAR(255) NOT NULL,
            email VARCHAR(255) NOT NULL,
            message TEXT NOT NULL
        )
    ''')
    connection.commit()
    connection.close()

# Update Excel Sheet
def update_excel(data):
    excel_file = 'form_data.xlsx'
    if not os.path.exists(excel_file):
        # Create a new Excel file with headers if it doesn't exist
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = 'Submissions'
        sheet.append(['Name', 'Email', 'Message'])
    else:
        # Load the existing Excel file
        workbook = openpyxl.load_workbook(excel_file)
        sheet = workbook.active

    # Append new data to the Excel sheet
    sheet.append([data['name'], data['email'], data['message']])
    workbook.save(excel_file)

# Root route
@app.route('/')
def home():
    return "Welcome to the Flask backend!"

# Route to handle form submission
@app.route('/submit', methods=['POST'])
def submit_form():
    data = request.json  # Get JSON data from the request

    # Save to MySQL
    connection = get_db_connection()
    cursor = connection.cursor()
    cursor.execute('''
        INSERT INTO form_submissions (name, email, message)
        VALUES (%s, %s, %s)
    ''', (data['name'], data['email'], data['message']))
    connection.commit()
    connection.close()

    # Update Excel Sheet
    update_excel(data)

    # Return a success response
    return jsonify({'message': 'Form submitted successfully!'})

if __name__ == '__main__':
    create_table()  # Ensure the table exists
    app.run(host='0.0.0.0', port=5000, debug=True)  # Allow public access
