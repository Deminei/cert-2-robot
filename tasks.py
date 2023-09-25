from robocorp.tasks import task
from robocorp import browser, http
from RPA.PDF import PDF
from RPA.Tables import Tables
from RPA.Archive import Archive


@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    browser.configure(slowmo=100)

    download_orders_file()
    open_robot_order_website()
    orders = get_orders()
    for row in orders:
        fill_the_form(row)
    archive_receipts()


def open_robot_order_website():
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")
    closs_annoying_modal()


def closs_annoying_modal():
    page = browser.page()
    page.click("text=OK")


def download_orders_file():
    """Downloads csv file from given URL"""
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

def get_orders():
    """Read data from the CSV file and return the orders"""
    library = Tables()
    orders = library.read_table_from_csv("orders.csv", columns=["Order number", "Head", "Body", "Legs", "Address"])
    return orders


def fill_the_form(row):
    """Fills in the order form"""
    page = browser.page()

    page.select_option("#head.custom-select", row["Head"])
    page.click("#id-body-" + row["Body"])
    page.fill("//html/body/div/div/div[1]/div/div[1]/form/div[3]/input", row["Legs"])
    page.fill("#address.form-control", row["Address"])

    page.click("text=Preview")

    order_number = row["Order number"]

    page.click("#order")

    while True:
        warning_message = page.query_selector(".alert.alert-danger")
        if warning_message is None:
            break  # Exit the loop if there is no warning message
        page.click("#order")  # Click #order again if there is a warning message

    store_receipt_as_pdf(order_number)


def store_receipt_as_pdf(order_number):
    page = browser.page()

    sales_receipt_html = page.locator("#receipt").inner_html()

    pdf = PDF()

    pdf_file = f"output/OrderReceipts/order_receipt_{order_number}.pdf"
    pdf.html_to_pdf(sales_receipt_html, pdf_file,)
    screenshot_robot(order_number, pdf_file)


def screenshot_robot(order_number, pdf_file):
    page = browser.page()

    screenshot = f"output/RobotScreenshots/robot_preview_{order_number}.png"

    robot_screenshot = page.query_selector("#robot-preview-image")
    robot_screenshot.screenshot(path=screenshot)

    embed_screenshot_to_receipt(screenshot, pdf_file)


def embed_screenshot_to_receipt(screenshot, pdf_file):
    pdf = PDF()
    pdf.add_watermark_image_to_pdf(image_path=screenshot, source_path=pdf_file, output_path=pdf_file)
    open_robot_order_website()


def archive_receipts():
    lib = Archive()
    lib.archive_folder_with_zip("./output/OrderReceipts", "output/OrderReceipts.zip", recursive=True)
