from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.Excel.Files import Files
from RPA.PDF import PDF
from RPA.FileSystem import FileSystem
from RPA.Archive import Archive

@task
def order_robots_from_portal():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    browser.configure(
        slowmo=100,
    )
    
    file_system = FileSystem()
    
    file_system.create_directory(path="output/receipts")
    file_system.create_file(path="output/receipts/receipt.pdf", content="receipt", overwrite=True)
    orders = get_orders()
    open_the_intranet_website()
    for row in orders:
        close_annoying_module()
        fill_the_form(row)
    archive_receipts()
   
def open_the_intranet_website():
    browser.goto("https://robotsparebinindustries.com/#/robot-order")
    
def close_annoying_module():
    page = browser.page()
    page.click("button:text('OK')")    

def get_orders():
    """Read data from csv table and fill in the sales form"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)
    
    library = Tables()
    tab = library.read_table_from_csv( "orders.csv", header=True, columns=["Order number", "Head", "Body", "Legs", "Address"] )
    return tab

def fill_the_form(current_row):
    page = browser.page()
    
    page.select_option("#head", str(current_row["Head"]))
    page.click("#id-body-"+current_row['Body'])
    page.fill("//html/body/div[1]/div/div[1]/div/div[1]/form/div[3]/input", current_row["Legs"])
    page.fill("#address", str(current_row["Address"]))
    page.click("#preview")
    page.click("#order")

    if not page.is_visible("#order-another"):
        while not page.is_visible("#order-another"):
            page.click("#order")
    
    store_receipt_as_pdf(current_row["Order number"])
    page.click("#order-another")
   
def store_receipt_as_pdf(order_number):
    pdf = PDF()
    file_system = FileSystem()
    
    """file_system.create_file(path="output/receipts/receipt.pdf", content="receipt", overwrite=True)"""
    pdf_file = "output/receipts/receipt.pdf"
    screenshot = screenshot_robot(order_number)
    embed_screenshot_to_receipt(screenshot, pdf_file)
    
def screenshot_robot(order_number):
    page = browser.page()
    
    page.screenshot(path="output/receipts/receipt"+str(order_number)+".png")
    screenshot = "output/receipts/receipt"+str(order_number)+".png"
    return screenshot

def embed_screenshot_to_receipt(screenshot, pdf_file):
    file_system = FileSystem()
    
    file_system.append_to_file(path=pdf_file, content=screenshot, encoding="utf-8")
    """pdf.add_files_to_pdf(files = screenshot, target_document = pdf_file, append=True)"""
    
def archive_receipts():
    lib = Archive()
    
    lib.archive_folder_with_zip(folder="output/receipts",archive_name="output/receipts.zip")
    
    
    
