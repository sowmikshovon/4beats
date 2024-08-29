from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import load_workbook
from datetime import datetime


def get_current_day_of_the_week():
    return datetime.now().strftime('%A')


driver = webdriver.Chrome()

try:
    excelFile = "C:\\Users\\ssowm\\Desktop\\f1\\4beatsAutomation&DevOpsInter\\4BeatsQ1.xlsx"
    workbook = load_workbook(excelFile)
    sheet = workbook[get_current_day_of_the_week()]

    columnIndex = 3
    for rowIndex in range(3, 13):
        row = sheet[rowIndex]
        cell = row[columnIndex-1]
        keyword = cell.value

        driver.get('https://www.google.com/')
        searchInput = driver.find_element(By.NAME, 'q')
        searchInput.send_keys(keyword)

        wait = WebDriverWait(driver, 10)
        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "ul[role='listbox']")))

        suggestionsContainer = driver.find_element(By.CSS_SELECTOR, "ul[role='listbox']")
        suggestionElements = suggestionsContainer.find_elements(By.TAG_NAME, "li")

        longestOption = ""
        shortestOption = ""

        for suggestionElement in suggestionElements:
            suggestion = suggestionElement.text
            if len(suggestion) > len(longestOption):
                longestOption = suggestion
            if not shortestOption or len(suggestion) < len(shortestOption):
                shortestOption = suggestion

        row[columnIndex].value = longestOption
        row[columnIndex + 1].value = shortestOption

    workbook.save(excelFile)

finally:
    driver.quit()

