from flask import Flask, request, render_template, send_file
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import openpyxl
import time

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process():
    start_roll = int(request.form['start_roll'])
    end_roll = int(request.form['end_roll'])

    # Set up the Selenium WebDriver
    driver = webdriver.Chrome()

    # Create an Excel workbook and add a sheet
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "CGPAs"
    ws.append(["Roll No", "Final CGPA"])

    try:
        for roll_no in range(start_roll, end_roll + 1):
            driver.get('http://results.mvsrec.edu.in/SBLogin.aspx')

            # Wait for username field
            username_field = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.NAME, 'txtUserName'))
            )
            username_field.clear()
            username_field.send_keys(str(roll_no))

            # Wait for password field
            password_field = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.NAME, 'txtPassword'))
            )
            password_field.clear()
            password_field.send_keys(str(roll_no))

            # Submit form
            submit_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.NAME, 'btnSubmit'))
            )
            submit_button.click()

            # Navigate and fetch CGPA
            exams_button = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, "//input[@id='Stud_cpModules_imgbtnExams']"))
            )
            exams_button.click()

            sem_end_results_link = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.LINK_TEXT, 'Semester End Exam Result'))
            )
            sem_end_results_link.click()

            soup = BeautifulSoup(driver.page_source, 'html.parser')
            final_cgpa_element = soup.select_one('span#Stud_cpBody_lblCGPA')
            final_cgpa = final_cgpa_element.get_text(strip=True) if final_cgpa_element else 'N/A'

            ws.append([roll_no, final_cgpa])
            time.sleep(1)

        # Save Excel file
        file_path = "FinalCGPA.xlsx"
        wb.save(file_path)
        return send_file(file_path, as_attachment=True)

    except Exception as e:
        return f"An error occurred: {e}"

    finally:
        driver.quit()

if __name__ == '__main__':
    app.run(debug=True)
