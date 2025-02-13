from flask import Flask, render_template, request, send_file
import pandas as pd
import asyncio, os, openpyxl
from twilio.rest import Client
import pymysql, math, datetime

# Twilio account credentials
account_sid = os.getenv("TWILIO_ACCOUNT_SID")
auth_token = os.getenv("TWILIO_AUTH_TOKEN")
twilio_client = Client(account_sid, auth_token)

# Function to read Excel file and convert to array
def read_excel_to_array(file_path):
    df = pd.read_excel(file_path)
    return df.values.tolist()

def header_read(file_path):
    df = pd.read_excel(file_path)
    return df.columns

def columns_read():
    wb = openpyxl.load_workbook('Marks1.xlsx')
    ws = wb.active
    return len(list(ws.iter_cols(values_only=True)))

def after_process_ese(file_path):
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active
    merged_ranges = list(ws.merged_cells.ranges)
    for merged_range in merged_ranges:
        ws.unmerge_cells(str(merged_range))
    for row in ws.iter_rows():
        for cell in row:
            if not isinstance(cell, MergedCell):
                cell.value = None
    wb.save(file_path)
    wb.close()

def after_process():
    wb = openpyxl.load_workbook('Marks1.xlsx') 
    ws = wb.active
    for row in ws.iter_rows():
        for cell in row:
            cell.value = None
            wb.save('Marks1.xlsx')
    return None

async def login_main(login, email, password):
    hod_email = os.getenv("HOD_EMAIL")
    hod_pwd = os.getenv("HOD_PWD")
    staff_email = os.getenv("STAFF_EMAIL")
    staff_pwd = os.getenv("STAFF_PWD")
    if str(login) == "HOD" and str(email) == hod_email and str(password) == hod_pwd:
        stat = "hod"
        return True
    elif str(login) == "Staff" and str(email) == staff_email and str(password) == staff_pwd:
        stat = "staff"
        return False
    else:
        stat = 'none'
        return stat

async def send_sms_message(name, count, sem, exam, year, ph_no, message, cursor, cnx):
    try:
        message = twilio_client.messages.create(
            from_='+13087734059',
            to=f"{ph_no}",
            body=message
        )
        query = "use status"
        cursor.execute(query)
        status = "DONE"
        query1 = "insert into status_data(name,arrear_count,sem,exam,year,Status) values (%s,%s,%s,%s,%s,%s)"
        values = [name, count, sem, exam, year, status]
        cursor.execute(query1, values)
        cnx.commit()
        query2 = "use status_rec"
        cursor.execute(query2)
        status = "DONE"
        now = datetime.datetime.now()
        query3 = "insert into status_data(name,arrear_count,sem,exam,year,Status,DATE) values (%s,%s,%s,%s,%s,%s,%s)"
        values = [name, count, sem, exam, year, status, now]
        cursor.execute(query3, values)
        cnx.commit()
        print(f"Message sent to {ph_no} regarding arrears.")
    except Exception as e:
        query = "use status"
        cursor.execute(query)
        status = "PENDING"
        query1 = "insert into status_data(name,arrear_count,sem,exam,year,Status) values (%s,%s,%s,%s,%s,%s)"
        values = [name, count, sem, exam, year, status]
        cursor.execute(query1, values)
        cnx.commit()
        query2 = "use status_rec"
        cursor.execute(query2)
        status = "PENDING"
        now = datetime.datetime.now()
        query3 = "insert into status_data(name,arrear_count,sem,exam,year,Status,DATE) values (%s,%s,%s,%s,%s,%s,%s)"
        values = [name, count, sem, exam, year, status, now]
        cursor.execute(query3, values)
        cnx.commit()
        print(f"Failed to send message to {ph_no}: {str(e)}")

def process_hod_data(year, sem, exam, arrear, cnx, cursor):
    data = None
    if arrear == 'three_arrear':
        cursor.execute("USE 3_arrear_data")
        query = "SELECT name, arrear_count,year,sem,exam FROM 3_arrear WHERE year = %s AND sem = %s AND exam = %s"
        cursor.execute(query, (year, sem, exam))
        data = cursor.fetchall()
    elif arrear == 'two_arrear':
        cursor.execute("USE 2_arrear_data")
        query = "SELECT name, arrear_count,year,sem,exam FROM 2_arrear WHERE year = %s AND sem = %s AND exam = %s"
        cursor.execute(query, (year, sem, exam))
        data = cursor.fetchall()
    elif arrear == 'one_arrear':
        cursor.execute("USE 1_arrear_data")
        query = "SELECT name, arrear_count,year,sem,exam FROM 1_arrear WHERE year = %s AND sem = %s AND exam = %s"
        cursor.execute(query, (year, sem, exam))
        data = cursor.fetchall()
    elif arrear == 'nil_arrear':
        cursor.execute("USE nil_arrear_data")
        query = "SELECT name, arrear_count,year,sem,exam FROM nil_arrear WHERE year = %s AND sem = %s AND exam = %s"
        cursor.execute(query, (year, sem, exam))
        data = cursor.fetchall()
    else:
        print("Invalid arrear type")
    return data

def clear_data(arrear, year, exam, sem):
    db_user = os.getenv("DB_USER")
    db_password = os.getenv("DB_PASSWORD")
    db_host = os.getenv("DB_HOST")
    cnx = pymysql.connect(
        cursorclass=pymysql.cursors.DictCursor,
        host=db_host,
        password=db_password,
        port=15274,
        user=db_user,
    )
    cursor = cnx.cursor()
    try:
        if arrear == 'three_arrear':
            cursor.execute("USE 3_arrear_data")
            quary = 'delete from 3_arrear where year=%s and exam=%s and sem=%s'
            values = (year, exam, sem)
            cursor.execute(quary, values)
        elif arrear == 'two_arrear':
            cursor.execute("USE 2_arrear_data")
            quary = 'delete from 2_arrear where year=%s and exam=%s and sem=%s'
            values = (year, exam, sem)
            cursor.execute(quary, values)
        elif arrear == 'one_arrear':
            cursor.execute("USE 1_arrear_data")
            quary = 'delete from 1_arrear where year=%s and exam=%s and sem=%s'
            values = (year, exam, sem)
            cursor.execute(quary, values)
        elif arrear == 'nil_arrear':
            cursor.execute("USE nil_arrear_data")
            quary = 'delete from nil_arrear where year=%s and exam=%s and sem=%s'
            values = (year, exam, sem)
            cursor.execute(quary, values)
        else:
            print("Invalid arrear type")

    finally:
        cnx.commit()
        cursor.close()
        cnx.close()
    return None

def staff_del_data():
    db_user = os.getenv("DB_USER")
    db_password = os.getenv("DB_PASSWORD")
    db_host = os.getenv("DB_HOST")
    cnx = pymysql.connect(
        cursorclass=pymysql.cursors.DictCursor,
        host=db_host,
        password=db_password,
        port=15274,
        user=db_user,
    )
    cursor = cnx.cursor()
    query = "USE all_data"
    cursor.execute(query)
    query1 = "DELETE FROM all_data1"
    cursor.execute(query1)
    cnx.commit()
    cursor.close()
    cnx.close()

def process_message_data():
    db_user = os.getenv("DB_USER")
    db_password = os.getenv("DB_PASSWORD")
    db_host = os.getenv("DB_HOST")
    cnx = pymysql.connect(
        cursorclass=pymysql.cursors.DictCursor,
        host=db_host,
        password=db_password,
        port=15274,
        user=db_user,
    )
    cursor = cnx.cursor()
    data1 = None 
    query = "USE all_data"
    cursor.execute(query)
    query1 = "SELECT * FROM all_data1"
    cursor.execute(query1)
    data1 = cursor.fetchall()
    cursor.close()
    cnx.close()
    return data1

def process_message_data1():
    db_user = os.getenv("DB_USER")
    db_password = os.getenv("DB_PASSWORD")
    db_host = os.getenv("DB_HOST")
    cnx = pymysql.connect(
        cursorclass=pymysql.cursors.DictCursor,
        host=db_host,
        password=db_password,
        port=15274,
        user=db_user,
    )
    cursor = cnx.cursor()
    data1 = None
    query = "USE all_data"
    cursor.execute(query)
    query1 = "SELECT * FROM all_data1"
    cursor.execute(query1)
    data1 = cursor.fetchall()
    cursor.close()
    cnx.close()
    return data1

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    file = request.files['file']
    file.save(os.path.join('uploads', file.filename))
    return 'File uploaded successfully'

if __name__ == '__main__':
    app.run(debug=True)
