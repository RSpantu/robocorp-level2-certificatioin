from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
import shutil
import logging
from datetime import datetime



@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """

    browser.configure(
        slowmo=100,
        browser_engine='chrome'
    )
    main_url = "https://robotsparebinindustries.com/#/robot-order"
    orders_url = "https://robotsparebinindustries.com/orders.csv"
    current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
    archive_name = f"Archive-{current_time}"
    archive_format = "zip"

    open_robot_order_website(main_url)
    close_annoying_modal()
    orders = get_orders(orders_url)

    for order in orders:
        fill_the_form(order)
        submitted_order_number = scrape_order_number()
        store_receipt_as_pdf(submitted_order_number)
        order_another_robot()
    
    archive_receipts(archive_name=archive_name, archive_format=archive_format)
    clean_output_folder("output/receipts")

    

def open_robot_order_website(url):
    """Navigates to the given URL"""
    browser.goto(url)


def close_annoying_modal():
    """ Close the pop-up"""
    logging.info("Close annoying pop-up…")
    page = browser.page()
    page.click("button:text('OK')")


def get_orders(orders_url):
    """ Download input table"""
    logging.info("Download the input…")
    http = HTTP()
    http.download(url=orders_url, overwrite = True)
    return read_local_csv("orders.csv")


def read_local_csv(csv_path):
    """ instantiate tables"""
    logging.info("Read the input…")
    tables = Tables()
    """try reading the table and handle exception"""
    try:
        input_table = tables.read_table_from_csv(csv_path, header=True)
    except FileNotFoundError:
        print("File not found.")
    finally:
        return input_table   


def fill_the_form(order): 
    """Fill in the form with the order details"""
    logging.info(f"Filling the order {order['Order number']}")

    page = browser.page()
    """Select the head"""
    page.select_option("xpath=//select[@id='head']", order["Head"])
    """Select the body"""
    robot_body = str(order['Body'])
    page.click(f"xpath=//input[@type='radio'][@value='{robot_body}']")
    """Select the legs number"""
    page.fill("xpath=//input[@type='number']", order['Legs'])
    """Fill in the address"""
    page.fill("xpath=//input[@id='address']", order['Address'])
    """Preview the order"""
    page.click("xpath=//button[@id='preview']")
    """Place the Order"""
    page.click("xpath=//button[@id='order']")
    
    while error_occured():
        page.click("xpath=//button[@id='order']")
    

def store_receipt_as_pdf(order_number):
    """Stores the order receipt as a PDF file"""
    logging.info("Storing the receipt as pdf…")
    page = browser.page()
    order_receipt_html = page.locator("#receipt").inner_html()
    pdf = PDF()
    file_name = f"order_{order_number}_receipt.pdf"
    pdf.html_to_pdf(order_receipt_html, f"output/receipts/{file_name}")

    screenshot_robot(order_number)


def screenshot_robot(order_number):
    """Takes a screenshot of the robot"""
    logging.info("capturing the screenshot of")
    """Capture the current page"""
    page = browser.page()
    """Create the screenshot path"""
    screenshot_path = f"output/receipts/order_{order_number}_screenshot.png"
    """Screenshot the robot's image."""
    page.locator("xpath=//div[@id='robot-preview-image']").screenshot(path=screenshot_path)

    """Embed robot's image into the order"""
    embed_screenshot_to_receipt(screenshot_path, f"output/receipts/order_{order_number}_receipt.pdf")


def embed_screenshot_to_receipt(screenshot, pdf_file):
    logging.info("Embedding the robot's preview screenshot into the receipt PDF…")
    pdf = PDF()
    pdf.add_watermark_image_to_pdf(
        image_path=screenshot,
        source_path=pdf_file,
        output_path=pdf_file
    )


def archive_receipts(archive_name, archive_format):
    logging.info("Archiving receipts …")
    """Create the archive given the name and the requested format and save it to a variable for reuse."""
    archive_path = shutil.make_archive(archive_name, archive_format)
    """Print the archive path on the console"""
    print("Archive created at %s", archive_path)
    return archive_path


def scrape_order_number():
    """Capture the opened page"""
    page = browser.page()
    """Define the searched locator"""
    text_locator = page.locator("xpath=//p[@class='badge badge-success']")

    """Ensure the locator is visible"""
    text_locator.wait_for(timeout=5000, state="visible")
    assert text_locator.is_visible()

    """ Scrape the text"""
    scraped_text = text_locator.inner_text()
    print(f"Scraped text: {scraped_text}")
    return scraped_text


def error_occured():
    """Capture the opened page"""
    page = browser.page()
    
    """Define the searched locator"""
    # searched_error = page.locator("xpath=//div[@role='alert']")
    # error_occured = page.query_selector("xpath=//div[@role='alert']")
    error_occured = page.query_selector("//div[@class='alert alert-danger']")
    if error_occured:
        logging.warn("Error occured. Need to submit the order again.")
    else:
        logging.info("No error occured, safe to continue.")
    
    return error_occured


def order_another_robot():
    """ Click on order another robot"""
    logging.info("Order another robot…")
    """Capture the page"""
    page = browser.page()
    """Click button order another robot"""
    page.click("xpath=//button[@id='order-another']")
    close_annoying_modal()

def clean_output_folder(folder_path):
    """Remove the output folder """
    shutil.rmtree(folder_path)
    print(f"The directory {folder_path} has been removed successfully.")