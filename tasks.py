from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive
import time
import os
import glob

MAX_RETRIES = 5
INITIAL_BACKOFF = 1  # Initial backoff time in seconds

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    open_robot_order_website()
    browser.configure(slowmo=100)
    download_csv_from_url()
    remove_existing_files_in_folders()
    orders = get_orders()
    for order in orders:
        retry_with_backoff(process_order, order)
    archive_receipts()

def open_robot_order_website():
    """Navigates to the given URL and closes the annoying modal"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")
    close_annoying_modal()

def close_annoying_modal():
    """Closes the annoying modal on the website"""
    page = browser.page()
    page.click("text=OK")

def download_csv_from_url():
    """Download CSV from URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

def remove_existing_files_in_folders():
    """"Remove exisiting Pdf and Screenshot from folders"""
    pdf_files = glob.glob("receipts/*.pdf")
    for pdf_file in pdf_files:
        os.remove(pdf_file)

    screenshot_files = glob.glob("screenshots/*.png")
    for screenshot_file in screenshot_files:
        os.remove(screenshot_file)

def get_orders():
    """Reads the downloaded orders CSV file and returns it as a table"""
    tables = Tables()
    orders = tables.read_table_from_csv("orders.csv")
    return orders

def process_order(order):
    """Processes each order"""
    print(f"Processing order: {order}")
    fill_the_form(order)
    embed_screenshot_to_receipt(order['Order number'])

def fill_the_form(order):
    """Fills the robot order form and submits the robot order with error handling"""
    page = browser.page()

    try:
        if order["Head"] == "1":
            value = "Roll-a-thor head"
        elif order["Head"] == "2":
            value = "Peanut crusher head"
        elif order["Head"] == "3":
            value = "D.A.V.E head"
        elif order["Head"] == "4":
            value = "Andy Roid head"
        elif order["Head"] == "5":
            value = "Spanner mate head"
        else:
            value = "Drillbit 2000 head"

        page.select_option("#head", str(order["Head"]))

        if order["Body"] == "1":
            page.check("#id-body-1")
        elif order["Body"] == "2":
            page.check("#id-body-2")
        elif order["Body"] == "3":
            page.check("#id-body-3")
        elif order["Body"] == "4":
            page.check("#id-body-4")
        elif order["Body"] == "5":
            page.check("#id-body-5")
        else:
            page.check("#id-body-6")

        page.get_by_placeholder("Enter the part number for the legs").fill(str(order["Legs"]))
        page.fill("#address", str(order["Address"]))
        page.click('#preview')

        retries = 0
        backoff = INITIAL_BACKOFF
        while retries < MAX_RETRIES:
            try:
                page.click('#order')
                print(f"Order {order['Order number']} submitted successfully!")
                return
            except Exception as e:
                print(f"Error: {e}. Retrying in {backoff} seconds...")
                time.sleep(backoff)
                retries += 1
                backoff *= 2  # Exponential backoff
        print("Max retries reached. Continuing to next task.")
    except Exception as e:
        print(f"Error while filling the form for order {order['Order number']}: {e}")

def embed_screenshot_to_receipt(order_number):
    """Stores the order receipt as a PDF file"""
    page = browser.page()
    pdf = PDF()
    html_content = page.locator("#receipt").inner_html()
    pdf_path = f"receipts/order_{order_number}.pdf"
    pdf.html_to_pdf(html_content, pdf_path)
    print(f"Receipt stored as PDF: {pdf_path}")
    
    """Takes a screenshot of the robot"""
    robot_preview_selector = "#robot-preview-image"
    screenshot_path = f"screenshots/robot_{order_number}.png"
    element_handle = page.query_selector(robot_preview_selector)
    if element_handle:
        element_handle.screenshot(path=screenshot_path)
    print(f"Screenshot taken: {screenshot_path}")

    """Embeds the robot screenshot into the receipt PDF file"""
    pdf.add_watermark_image_to_pdf(
        image_path= screenshot_path,
        source_path= pdf_path,
        output_path= pdf_path
        )
    page.click("#order-another")
    page.click("text=OK")
    print(f"Screenshot embedded into receipt PDF: {pdf_path}")

def archive_receipts():
    """Create ZIP archive of PDF receipts and store in the output directory"""
    archive = Archive()
    pdf_folder = "receipts"
    zip_path = "output/receipts_archive.zip"
    
    if os.path.exists(zip_path):
        os.remove(zip_path)
        
    archive.archive_folder_with_zip(pdf_folder, zip_path)
    
def retry_with_backoff(func, *args, **kwargs):
    """Retry function with exponential backoff"""
    retries = 0
    backoff = INITIAL_BACKOFF
    while retries < MAX_RETRIES:
        try:
            func(*args, **kwargs)
            return
        except Exception as e:
            print(f"Error: {e}. Retrying in {backoff} seconds...")
            time.sleep(backoff)
            retries += 1
            backoff *= 2  # Exponential backoff
    print("Max retries reached. Continuing to next task.")
