# Executes every monday 12pm using task scheduler.
# Must set "run when user is logged on" in task scheduler because selenium mimics user behavior.
# Scrapes programming life insta acc's previous week's posts (mon to sun) and insert it into a word file.
import time
import datetime
import pandas as pd
from docx import Document
from docx.oxml.ns import qn
from docx.enum.style import WD_STYLE_TYPE
from selenium import webdriver
from selenium.webdriver.common.by import By
from win10toast_persist import ToastNotifier
# Only used for showing intellisense autocomplete, would crash with other import if used
# from docx.document import Document

def RetrievePostComments(postCaption: str) -> str:
    commentsExist1 = len(driver.find_elements(By.XPATH, "//body[1]/div[5]/div[1]/div[1]/div[3]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/article[1]/div[1]/div[2]/div[1]/div[1]/div[2]/div[1]/ul[1]/ul"))
    for i in range(commentsExist1, 0, -1):
        postCaption += "\n" + driver.find_element(By.XPATH, f"//body[1]/div[5]/div[1]/div[1]/div[3]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/article[1]/div[1]/div[2]/div[1]/div[1]/div[2]/div[1]/ul[1]/ul[{i}]/div[1]/li[1]/div[1]/div[1]/div[2]/div[1]/span[1]").text
    commentsExist2 = len(driver.find_elements(By.XPATH, "//body[1]/div[6]/div[1]/div[1]/div[3]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/article[1]/div[1]/div[2]/div[1]/div[1]/div[2]/div[1]/ul[1]/div[3]/div[1]/div[1]/div"))  # Added on 20231204
    for i in range(commentsExist2, 0, -1):
        postCaption += "\n" + driver.find_element(By.XPATH, f"//body[1]/div[6]/div[1]/div[1]/div[3]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/article[1]/div[1]/div[2]/div[1]/div[1]/div[2]/div[1]/ul[1]/div[3]/div[1]/div[1]/div[{i}]/ul[1]/div[1]/li[1]/div[1]/div[1]/div[2]/div[1]/span[1]").text
    return postCaption

def ClickNextButtonFirst():
    nextButtonExist = len(driver.find_elements(By.XPATH, "//body[1]/div[6]/div[1]/div[1]/div[3]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/button[1]"))
    if nextButtonExist != 0:
        driver.find_element(By.XPATH, "//body[1]/div[6]/div[1]/div[1]/div[3]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/button[1]").click()
    else:
        driver.find_element(By.XPATH, "//body[1]/div[7]/div[1]/div[1]/div[3]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/button[1]").click()

def ClickNextButtonSubsequent():
    nextButtonExist = len(driver.find_elements(By.XPATH, "//body[1]/div[5]/div[1]/div[1]/div[3]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[2]/button[1]"))
    if nextButtonExist != 0:
        driver.find_element(By.XPATH, "//body[1]/div[5]/div[1]/div[1]/div[3]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[2]/button[1]").click()
    else:
        driver.find_element(By.XPATH, "//body[1]/div[6]/div[1]/div[1]/div[3]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[2]/button[1]").click()

def ScrapePostCaption() -> str:
    postCaptionExist1 = len(driver.find_elements(By.XPATH, "//body[1]/div[5]/div[1]/div[1]/div[3]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/article[1]/div[1]/div[2]/div[1]/div[1]/div[2]/div[1]/ul[1]/div[1]/li[1]/div[1]/div[1]/div[2]/div[1]/h1[1]"))
    if postCaptionExist1 != 0:
        postCaption = driver.find_element(By.XPATH, "//body[1]/div[5]/div[1]/div[1]/div[3]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/article[1]/div[1]/div[2]/div[1]/div[1]/div[2]/div[1]/ul[1]/div[1]/li[1]/div[1]/div[1]/div[2]/div[1]/h1[1]").text
    postCaptionExist2 = len(driver.find_elements(By.XPATH, "//body[1]/div[6]/div[1]/div[1]/div[3]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/article[1]/div[1]/div[2]/div[1]/div[1]/div[2]/div[1]/ul[1]/div[1]/li[1]/div[1]/div[1]/div[2]/div[1]/h1[1]"))
    if postCaptionExist2 != 0:
        postCaption = driver.find_element(By.XPATH, "//body[1]/div[6]/div[1]/div[1]/div[3]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/article[1]/div[1]/div[2]/div[1]/div[1]/div[2]/div[1]/ul[1]/div[1]/li[1]/div[1]/div[1]/div[2]/div[1]/h1[1]").text
    postCaptionExist3 = len(driver.find_elements(By.XPATH, "//body[1]/div[7]/div[1]/div[1]/div[3]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/article[1]/div[1]/div[2]/div[1]/div[1]/div[2]/div[1]/ul[1]/div[1]/li[1]/div[1]/div[1]/div[2]/div[1]/h1[1]"))
    if postCaptionExist3 != 0:
        postCaption = driver.find_element(By.XPATH, "//body[1]/div[7]/div[1]/div[1]/div[3]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/article[1]/div[1]/div[2]/div[1]/div[1]/div[2]/div[1]/ul[1]/div[1]/li[1]/div[1]/div[1]/div[2]/div[1]/h1[1]").text
    if postCaptionExist1 == 0 and postCaptionExist2 == 0 and postCaptionExist3 == 0:
        raise Exception("Can't find post caption.")
    return postCaption

try:
    # Set script initial status.
    scriptStatus = True
    
    # Notify me the scrapper operation has started
    notif = ToastNotifier()
    dateToday = datetime.date.today()
    startDate = (dateToday - pd.Timedelta('7d')).strftime("%Y%m%d")
    endDate = (dateToday - pd.Timedelta('1d')).strftime("%Y%m%d")
    notif.show_toast("Programming Life Insta Scrapper", f"Operation started.\n{startDate} ~ {endDate}", duration=5)
    
    # Configure driver options and services.
    options = webdriver.ChromeOptions()
    options.page_load_strategy = 'normal'
    options.add_argument("--headless")
    options.add_argument("--window-size=1920x1080")
    logPath = "C:/Users/Admin/Documents/Selenium/InstaScraping"
    service = webdriver.ChromeService(service_args=['--log-level=DEBUG', '--append-log', '--readable-timestamp'], log_output=logPath)
    driver = webdriver.Chrome(options, service)

    # Try to access the website
    driver.implicitly_wait(100)
    driver.get("https://www.instagram.com/")

    # Checks whether there are any title tags same as specified. If none, throw exception.
    time.sleep(5)
    title = driver.title
    assert title == "Instagram"

    # Find column with name=username and name=password then fill in value.
    time.sleep(5)
    driver.implicitly_wait(10)
    username = ""  # Insert own IG username.
    password = ""  # Insert own IG password. 
    driver.find_element(By.NAME, "username").send_keys(username)
    driver.find_element(By.NAME, "password").send_keys(password)

    # Click login button.
    time.sleep(2)
    driver.find_element(By.XPATH, "//button[@type='submit']").click()

    # Access the profile page again.
    time.sleep(10)
    driver.get(f"https://www.instagram.com/{username}/")

    # Click open 1st post.
    firstPostExist1 = len(driver.find_elements(By.XPATH, "//body[1]/div[2]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/section[1]/main[1]/div[1]/div[3]/article[1]/div[1]/div[1]/div[1]/div[1]/a[1]"))
    if firstPostExist1 != 0:
        driver.find_element(By.XPATH, "//body[1]/div[2]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/section[1]/main[1]/div[1]/div[3]/article[1]/div[1]/div[1]/div[1]/div[1]/a[1]").click()
    firstPostExist2 = len(driver.find_elements(By.XPATH, "//body[1]/div[2]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/section[1]/main[1]/div[1]/div[3]/article[1]/div[1]/div[1]/div[1]/div[1]/a[1]"))  # Added on 20231204
    if firstPostExist2 != 0:
        driver.find_element(By.XPATH, "//body[1]/div[2]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/section[1]/main[1]/div[1]/div[3]/article[1]/div[1]/div[1]/div[1]/div[1]/a[1]").click()
    firstPostExist3 = len(driver.find_elements(By.XPATH, "//body[1]/div[2]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/section[1]/main[1]/div[1]/div[3]/div[1]/div[1]/div[1]/a[1]"))  # Added on 20240212
    if firstPostExist3 != 0:
        driver.find_element(By.XPATH, "//body[1]/div[2]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/section[1]/main[1]/div[1]/div[3]/div[1]/div[1]/div[1]/a[1]").click()
    firstPostExist4 = len(driver.find_elements(By.XPATH, "//body[1]/div[2]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[2]/div[1]/div[2]/section[1]/main[1]/div[1]/div[2]/div[1]/div[1]/div[1]/a[1]"))  # Added on 20240524
    if firstPostExist4 != 0:
        driver.find_element(By.XPATH, "//body[1]/div[2]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[2]/div[1]/div[2]/section[1]/main[1]/div[1]/div[2]/div[1]/div[1]/div[1]/a[1]").click()
    if firstPostExist1 == 0 and firstPostExist2 == 0 and firstPostExist3 == 0 and firstPostExist4 == 0:
        raise Exception("Can't find first post to click.")

    # Copy 1st post caption & comments.
    postCaption1 = ScrapePostCaption()
    postCaption1 = RetrievePostComments(postCaption1)
    postCaptionList = []
    postCaptionList.append(postCaption1)

    # Click next post button then copy 2nd post caption & comments.
    ClickNextButtonFirst()
    postCaption2 = ScrapePostCaption()
    postCaption2 = RetrievePostComments(postCaption2)
    postCaptionList.append(postCaption2)

    # Click next post button then copy 3rd to 7th post caption & comments. (2 div's of next buttonxpath has [-1] compared to previous xpath, same for subsequent)
    for i in range(5): 
        ClickNextButtonSubsequent()
        postCaption = ScrapePostCaption()
        postCaption = RetrievePostComments(postCaption)
        postCaptionList.append(postCaption)

    # Close Selenium session.
    driver.quit()

    # Set docx file path.
    docxFilePath = f"C:/Users/Admin/Desktop/Programming Life/Programming Life Record/Programming Life Word Record {startDate} ~ {endDate}.docx"

    # Create word document instance.
    doc = Document()

    # Adding a custom style with a page break before it. Write english in Calibri, and chinese in KaiTi.
    customStyle = "page_break_before"
    style = doc.styles.add_style(customStyle, WD_STYLE_TYPE.PARAGRAPH)
    style.paragraph_format.page_break_before = True
    style.font.name = "Calibri"
    rPr = style.element.get_or_add_rPr()
    rPr.rFonts.set(qn('w:eastAsia'), 'KaiTi')

    # Paste copied post captions into document.
    postCaptionList.reverse()
    for postCaption in postCaptionList:
        doc.add_paragraph(postCaption, style=customStyle)
    doc.save(docxFilePath)
    
except Exception as exception:
    scriptStatus = False
    errorMessage = str(exception)
    
# Add a notification to notify me about the success or failure.
notif = ToastNotifier()
if scriptStatus:
    notif.show_toast("Programming Life Insta Scrapper", f"Operation Successful.\n{startDate} ~ {endDate}", duration=5)
    print("Script completed.") 
else:
    notif.show_toast("Programming Life Insta Scrapper", f"Operation Failed.\n{startDate} ~ {endDate}\n" + errorMessage, duration=5)
    print("Script failed.") 





