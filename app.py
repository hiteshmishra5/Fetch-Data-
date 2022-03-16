from flask import Flask, render_template, Response
import psycopg2  
import psycopg2.extras
import io
import xlwt  
 
app = Flask(__name__)
 
DB_HOST = "localhost"
DB_NAME = "sampledb"
DB_USER = "postgres"
DB_PASS = "sai"
 
conn = psycopg2.connect(dbname=DB_NAME, user=DB_USER, password=DB_PASS, host=DB_HOST)
 
@app.route('/')
def Index():
    return render_template('index.html')
 
@app.route('/download/report/excel')
def download_report():
    cur = conn.cursor(cursor_factory=psycopg2.extras.DictCursor)
     
    cur.execute("SELECT * FROM technician1")
    result = cur.fetchall()
    for row in result:
       print(row)
 
    #output in bytes
    output = io.BytesIO()

    #create WorkBook object
    workbook = xlwt.Workbook()

    #add a sheet
    sh = workbook.add_sheet('Patient Report')
 
    #add headers
    sh.write(0, 0, 'sno')
    sh.write(0, 1, 'case_source')
    sh.write(0, 2, 'booking_date')
    sh.write(0, 3, 'booking_id')
    sh.write(0, 4, 'patient_name')
    sh.write(0, 5, 'type_of_test')
    sh.write(0, 6, 'case_done')
    sh.write(0, 7, 'case_reported')
    sh.write(0, 8, 'reported_by')
    sh.write(0, 9, 'technician_name')
    sh.write(0, 10, 'unit_no')
    sh.write(0, 11, 'bill_amount')
    sh.write(0, 12, 'bill_number')
    sh.write(0, 13, 'comments')
    idx = 0
    for row in result:
        sh.write(idx+1, 0, str(row['sno']))
        sh.write(idx+1, 1, row['case_source'])
        sh.write(idx+1, 2, row['booking_date'])
        sh.write(idx+1, 3, row['booking_id'])
        sh.write(idx+1, 4, row['patient_name'])
        sh.write(idx+1, 5, row['type_of_test'])
        sh.write(idx+1, 6, row['case_done'])
        sh.write(idx+1, 7, row['case_reported'])
        sh.write(idx+1, 8, row['reported_by'])
        sh.write(idx+1, 9, row['technician_name'])
        sh.write(idx+1, 10, row['unit_no'])
        sh.write(idx+1, 11, row['bill_amount'])
        sh.write(idx+1, 12, row['bill_number'])
        sh.write(idx+1, 13, row['comments'])
        idx += 1
 
    workbook.save(output)
    output.seek(0)
 
    return Response(output, mimetype="application/ms-excel", headers={"Content-Disposition":"attachment;filename=student_report.xls"})
 
if __name__ == "__main__":
    app.run(debug=True)