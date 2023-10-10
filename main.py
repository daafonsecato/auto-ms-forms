import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

driver = webdriver.Chrome()  # assuming Chrome browser


def fill_form(udemy_name, udemy_url, price, cloud_provider, du_name, region, email):
    # Initialize the browser and open the URL
    # replace with the actual URL
    # # Fill Udemy training name
    # driver.find_element(
    #     By.CSS_SELECTOR, 'input[aria-labelledby="QuestionId_raf4d57c0f0284c27aa377f0b38846ff9 QuestionInfo_raf4d57c0f0284c27aa377f0b38846ff9"]').send_keys(udemy_name)

    element = WebDriverWait(driver, 60).until(
        EC.presence_of_element_located(
            (By.CSS_SELECTOR, 'input[aria-labelledby="QuestionId_raf4d57c0f0284c27aa377f0b38846ff9 QuestionInfo_raf4d57c0f0284c27aa377f0b38846ff9"]'))
    )
    element.send_keys(udemy_name)
    # # Fill Udemy training URL
    driver.find_element(
        By.CSS_SELECTOR, 'input[aria-labelledby="QuestionId_r969508977068400995f8c6a675536532 QuestionInfo_r969508977068400995f8c6a675536532"]').send_keys(udemy_url)

    # Fill full price in GBP
    driver.find_element(
        By.CSS_SELECTOR, 'input[aria-labelledby="QuestionId_rcc6b4358e2d041248f9c575e3df50acc QuestionInfo_rcc6b4358e2d041248f9c575e3df50acc"]').send_keys(price)

    # Choose cloud provider
    driver.find_element(
        By.CSS_SELECTOR, f'input[value="{cloud_provider}"]').click()

    # Fill DU name
    driver.find_element(
        By.CSS_SELECTOR, 'input[aria-labelledby="QuestionId_r313b963268cb4a6590661723ac3f010d QuestionInfo_r313b963268cb4a6590661723ac3f010d"]').send_keys(du_name)

    # Choose region
    driver.find_element(By.CSS_SELECTOR, f'input[value="{region}"]').click()

    # Fill Endava email address
    driver.find_element(
        By.CSS_SELECTOR, 'input[aria-labelledby="QuestionId_r7925d381d7c94d77a277268ae2b3c216 QuestionInfo_r7925d381d7c94d77a277268ae2b3c216"]').send_keys(email)

    # Click submit button(uncomment if you want to actually submit)
    driver.find_element(
        By.CSS_SELECTOR, 'button[data-automation-id="submitButton"]').click()


# Read the Excel file
data = pd.read_excel('document.xlsx')

# URL of the form
URL = 'https://forms.office.com/Pages/ResponsePage.aspx?id=eME_CzC3i06YQ-gSWSN7d1EAI-pF_WtOuaWMpb37JgpUNU5ISUU2RDlUQzFXNEhLWlNQQlBBN0tYUC4u'

# Iterate through each row of the Excel file
for index, row in data.iterrows():
    if index == 0:
        driver.get(URL)
        # Wait for an element from the form to load after either loading or refreshing the page
        element_to_wait_for = 'input[aria-labelledby="QuestionId_raf4d57c0f0284c27aa377f0b38846ff9 QuestionInfo_raf4d57c0f0284c27aa377f0b38846ff9"]'
        WebDriverWait(driver, 60).until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, element_to_wait_for))
        )
    else:
        driver.refresh()

    fill_form(
        row['Udemy training name'],
        row['Udemy training URL'],
        row['Full price in GBP  (without discounts)'],
        row['Cloud Provider'],
        row['DU'],
        row['Region'],
        row['Endava email address of the person that should receive the training']
    )

# Keep the browser open for debugging
input("Press Enter to close the browser...")
driver.quit()
