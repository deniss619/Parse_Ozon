import re
import time
from selenium import webdriver


class OzonUtils:
    def parse_page(self, driver):
        blocks = driver.find_elements_by_xpath("//div[@class='e3f5']")
        for i in range(8):
            driver.execute_script(f"scroll ({1080 * i}, {1080 * i + 1080});")
            time.sleep(0.5)

        data = []
        for block in blocks:
            name = block.find_element_by_xpath(
                ".//span[contains(@class,'j4 as3 az a0f2 f-tsBodyL item e3t0')]/span").get_attribute('innerHTML')
            price_block = block.find_elements_by_xpath(".//div[contains(@class,'b5v4 e3r9 item')]/span")
            if len(price_block) == 1:
                full_price = discount_price = price_block[0].get_attribute('innerHTML')
            else:
                full_price = price_block[1].get_attribute('innerHTML')
                discount_price = price_block[0].get_attribute('innerHTML')

            try:
                review = int(
                    re.findall('\d+', block.find_element_by_xpath(".//a[@class='a1r7']").get_attribute('innerHTML'))[0])
            except:
                review = '0'

            try:
                description = block.find_element_by_xpath(".//span[@class='j4 as3 a0f6 f-tsBodyM item e3t']/span").text
            except:
                description = ''
            data.append([name, description, review, re.sub(r'\s+', '', discount_price), re.sub(r'\s+', '', full_price)])
        return data

    def parse_query(self, user_query):
        driver = get_driver()
        try:
            driver.get("https://www.ozon.ru/")
            driver.find_element_by_xpath("//div[@class='f9j4']/input").send_keys(user_query)
            time.sleep(1)
            driver.find_element_by_xpath("//button[@class='f9k']").click()
            time.sleep(3)
            try:
                total = re.sub('[^0-9]', '', driver.find_element_by_class_name('b6r7').text)
                url = driver.current_url
                driver.get(url[:url.find('&') + 1] + 'sorting=rating' + '&' + url[url.find('&') + 1:])
                data = []
                for i in range(3):
                    time.sleep(3)
                    data.append(self.parse_page(driver))
                    try:
                        driver.find_element_by_class_name('b7t1').find_element_by_class_name('kxa6').click()
                    except:
                        break
                return data, total
            except Exception as e:
                print(e)
                print("По данному запросу ничего не найдено!")
                return None, None
        except Exception as e:
            print(e)
            print("Произошла ошибка попробуйте повторить")
            return None, None
        finally:
            driver.close()


def get_driver():
    chrome_options = webdriver.ChromeOptions()
    # chrome_options.add_argument("--headless")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--log-level=3")
    chrome_options.add_argument("--start-maximized")
    return webdriver.Chrome(options=chrome_options)
