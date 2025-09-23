# -*- coding: utf-8 -*-

import os
import time
import csv
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup

# --- LOAD CONFIGURATION ---
# Load variables from the .env file (where login and password are stored)
load_dotenv()

# Get login and password from environment variables
WEBSITE_LOGIN = os.getenv("WEBSITE_LOGIN")
WEBSITE_PASSWORD = os.getenv("WEBSITE_PASSWORD")

# --- SETTINGS ---
# URL of the login page
LOGIN_URL = "https://lepolek.pl/logowanie"

# Path to your chromedriver.exe.
DRIVER_PATH = ""

# Output filename for the results
OUTPUT_FILE = "scraped_data_auto.csv"

# --- MAIN SCRIPT CODE ---

def setup_driver():
    """Configures and starts the Selenium web driver."""
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")
    if DRIVER_PATH:
        service = Service(executable_path=DRIVER_PATH)
        driver = webdriver.Chrome(service=service, options=options)
    else:
        driver = webdriver.Chrome(options=options)
    return driver

def login_to_website(driver):
    """Performs automatic login and waits for manual confirmation."""
    try:
        print(f"Navigating to login page: {LOGIN_URL}")
        driver.get(LOGIN_URL)
        
        wait = WebDriverWait(driver, 20)
        
        username_field = wait.until(EC.presence_of_element_located((By.NAME, "username")))
        password_field = driver.find_element(By.NAME, "password")
        login_button = driver.find_element(By.CSS_SELECTOR, "button[type='submit']")

        print("Entering login and password...")
        username_field.send_keys(WEBSITE_LOGIN)
        password_field.send_keys(WEBSITE_PASSWORD)
        
        print("Clicking the 'Login' button...")
        login_button.click()
        
        # --- ИЗМЕНЕНО: Возвращено ручное подтверждение ---
        print("\n--- MANUAL CONFIRMATION REQUIRED ---")
        print("Please check the browser window to confirm login was successful.")
        print("After you see your account page, press Enter in this console to continue...")
        input() # Скрипт ждет, пока вы нажмете Enter
        
        print("Login confirmed by user.")
        return True
        
    except Exception as e:
        print("An error occurred during the login process.")
        print(f"Details: {e}")
        return False

def scrape_page_data(driver, url):
    """Navigates to the specified URL and scrapes its data."""
    try:
        print(f"Navigating to data page: {url}")
        driver.get(url)
        
        print("Waiting for content to load...")
        wait = WebDriverWait(driver, 15)
        # Ждем появления основного блока со статьями
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "main")))
        print("Content loaded.")
        
        time.sleep(2) # Дополнительная пауза для прогрузки всех элементов
        html_content = driver.page_source
        soup = BeautifulSoup(html_content, 'html.parser')
        
        # Находим все контейнеры с препаратами (тег <article>)
        items = soup.select('article')
        
        if not items:
            print("No articles found on the page. Please check the selectors if the page structure has changed.")
            return None

        scraped_results = []
        print(f"Found {len(items)} items to scrape.")

        for item in items:
            # Извлекаем данные из каждого контейнера
            try:
                # Название препарата находится в теге <h2>
                drug_name = item.select_one('h2').get_text(strip=True)
                # Действующее вещество находится в теге <p>
                active_substance = item.select_one('p').get_text(strip=True)
                
                scraped_results.append({
                    'drug_name': drug_name, 
                    'active_substance': active_substance
                })
            except AttributeError:
                # Пропускаем, если у элемента нет h2 или p
                continue
                
        return scraped_results

    except Exception as e:
        print(f"An error occurred while scraping data: {e}")
        return None

def save_to_csv(data, filename):
    """Saves data to a CSV file."""
    if not data:
        print("No data to save.")
        return
    with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=data[0].keys())
        writer.writeheader()
        writer.writerows(data)
    print(f"Data successfully saved to {filename}")


if __name__ == "__main__":
    if not WEBSITE_LOGIN or not WEBSITE_PASSWORD:
        print("Error: Login or password not found in the .env file. Please check it.")
    else:
        driver = None
        try:
            driver = setup_driver()
            if login_to_website(driver):
                target_url = input("Login complete. Now, please paste the URL of the page to scrape and press Enter: ")
                
                results = scrape_page_data(driver, target_url)
                if results:
                    save_to_csv(results, OUTPUT_FILE)
        finally:
            if driver:
                print("Closing browser...")
                driver.quit()
