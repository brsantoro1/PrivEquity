from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, WebDriverException, TimeoutException

def search_webpage_adamsstreet(name):
    url = "https://www.adamsstreetpartners.com/our-firm/team/"
    full_url = url + name.lower().replace(' ', '-') + '/'
    span_xpath = "/html/body/div[7]/div/div/div[1]/span[1]"  # Updated XPath to locate the desired element
    get_text(full_url, span_xpath)


def get_text(name_url, span_xpath):
    try:
        driver = webdriver.Chrome()  # Assuming Webdriver location is added to Path
        driver.get(name_url)

        try:
            # If there's a search box or a way to filter the results, use it
            # Otherwise, if you are directly accessing the element, you can skip this step
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, span_xpath)) # Adjust this XPath as needed
            )
            
            # Find the element with the specified XPath and extract text
            span_element = driver.find_element(By.XPATH, span_xpath)
            extracted_text = span_element.text

            return extracted_text

        except (NoSuchElementException, TimeoutException) as e:
            print(f"Error occurred: {e}")
            return "PERSON NOT FOUND"
        finally:
            driver.quit()  # Close the browser in any case
    except WebDriverException as e:
        print(f"WebDriverException occurred: {e}")
        return "Page not found"
