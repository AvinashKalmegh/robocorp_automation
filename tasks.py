from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.PDF import PDF
import csv
import zipfile
import os

@task
def order_robots_from_RobotSpareBin():
    browser.configure(slowmo=200)
    try:
        open_robot_order_website()
        get_orders()
        fill_form_with_excel_data()
        #zip
        directory_path = 'output'
        output_zip_filename = 'output.zip'
        zip_pdf_files(directory_path, output_zip_filename)

    except Exception as e:
        print(f"Error occurred: {e}")

def open_robot_order_website():
    try:
        page = browser.page()
        browser.goto("https://robotsparebinindustries.com/#/robot-order")
        page.click("button:text('OK')")
    except Exception as e:
        print(f"Error navigating to website: {e}")

def export_as_pdf(key):
    try:
        page = browser.page()
        orders_results_html = page.locator("#receipt").inner_html()
        pdf = PDF()
        pdf.html_to_pdf(orders_results_html, f"output/order_results{key}.pdf")
    except Exception as e:
        print(f"Error exporting as PDF: {e}")

def collect_results(key):
    try:
        page = browser.page()
        page.screenshot(path=f"output/order_results{key}.png")
    except Exception as e:
        print(f"Error taking screenshot: {e}")

def add_screenshots_to_pdf(key):
    try:
        pdf = PDF()
        pdf.open_pdf(f"output/order_results{key}.pdf")
        pdf.add_files_to_pdf(files=[f'output/order_results{key}.png'], target_document=f"output/order_results{key}.pdf", append=True)
        pdf.close_all_pdfs()
    except Exception as e:
        print(f"Error adding screenshots to PDF: {e}")

def fill_and_submit_orders_form(sales_rep):
    try:
        page = browser.page()
        page.select_option("#head", str(sales_rep["Head"]))
        locator = f"#id-body-{sales_rep['Body']}"
        page.click(locator)  
        page.fill("input[type='number']", sales_rep["Legs"])
        page.fill("#address", sales_rep["Address"])
        page.click("button:text('Preview')")
        page.click("button:text('Order')")
        error_message = page.locator("div.alert.alert-danger")
        for i in range(10):
            if error_message.is_visible():
                page.click("button:text('Order')")
            else:
                break
            
    except Exception as e:
        print(f"Error filling and submitting form: {e}")

def get_orders():
    try:
        http = HTTP()
        http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)
    except Exception as e:
        print(f"Error downloading orders: {e}")

def fill_form_with_excel_data():
    try:
        with open('orders.csv', 'r', newline='') as file:
            reader = csv.DictReader(file)
            for row in reader:
                key = row['Order number']
                fill_and_submit_orders_form(row)
                export_as_pdf(key)
                collect_results(key)
                add_screenshots_to_pdf(key)
                page = browser.page()
                page.click("button:text('Order another robot')")
                page.click("button:text('OK')")
    except Exception as e:
        print(f"Error processing Excel data: {e}")

def zip_pdf_files(directory_path, output_zip_filename):
    try:
        if not os.path.exists(directory_path):
            raise FileNotFoundError(f"Directory '{directory_path}' does not exist.")
        pdf_files = [file for file in os.listdir(directory_path) if file.lower().endswith('.pdf')]
        if not pdf_files:
            raise FileNotFoundError(f"No PDF files found in '{directory_path}'.")
        with zipfile.ZipFile(output_zip_filename, 'w') as zipf:
            for pdf_file in pdf_files:
                file_path = os.path.join(directory_path, pdf_file)
                zipf.write(file_path, os.path.basename(file_path))
        print(f"Zip file '{output_zip_filename}' created successfully.")
    except Exception as e:
        print(f"Error creating ZIP file: {e}")

