import logging
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import openpyxl as excel
import urllib.parse
import cv2
import numpy as np
from PIL import Image

def readContacts(fileName):

    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
    logger = logging.getLogger()

    file = excel.load_workbook(fileName)
    sheet = file.active
    firstCol = sheet['A']
    secondCol = sheet['B']

    for cell in range(len(firstCol)):
        options = webdriver.ChromeOptions()
        options.add_argument('--user-data-dir=./userdata')
        driver = webdriver.Chrome(options=options)
        contact = str(firstCol[cell].value)
        message = str(secondCol[cell].value)
        driver.get("https://web.whatsapp.com/send/?phone=" + contact + "&text=" + urllib.parse.quote_plus(message))
        time.sleep(15)
        driver.save_screenshot('screenshot.png')

        logger.info('Passou dos 15 segundos')

        screenshot = cv2.imread('screenshot.png')
        template = cv2.imread('loading.png')

        if screenshot is None or template is None:
            logger.info("Erro ao carregar as imagens. Verifique os caminhos.")
            continue

        screenshot_gray = cv2.cvtColor(screenshot, cv2.COLOR_BGR2GRAY)
        template_gray = cv2.cvtColor(template, cv2.COLOR_BGR2GRAY)

        result = cv2.matchTemplate(screenshot_gray, template_gray, cv2.TM_CCOEFF_NORMED)

        threshold = 0.8
        loc = np.where(result >= threshold)

        if loc[0].size > 0:
            logger.info("O loading está presente na tela.")
            time.sleep(15)
        else:
            logger.info("O loading não foi encontrado na tela.")

        try:
            driver.save_screenshot('screenshot-tela.png')
        
            screenshot = cv2.imread('screenshot-tela.png')
            template = cv2.imread('btn wpp.png')

            if screenshot is None or template is None:
                logger.info("Erro ao carregar as imagens. Verifique os caminhos.")
                continue

            screenshot_gray = cv2.cvtColor(screenshot, cv2.COLOR_BGR2GRAY)
            template_gray = cv2.cvtColor(template, cv2.COLOR_BGR2GRAY)
        
            result = cv2.matchTemplate(screenshot_gray, template_gray, cv2.TM_CCOEFF_NORMED)
        
            threshold = 0.8
            loc = np.where(result >= threshold)

            if loc[0].size == 0: 
                time.sleep(15)

            if loc[0].size > 0:
                logger.info("O botão está presente na tela.")
                
                window_width = driver.execute_script("return window.innerWidth")
                window_height = driver.execute_script("return window.innerHeight")

                x = window_width - 15 
                y = window_height - 4 

                screenshot = cv2.imread('screenshot.png')
                if screenshot is not None:
                    region = screenshot[int(y)-52:int(y), int(x)-40:int(x)]
                    cv2.imwrite('region_to_click.png', region)
                    logger.info("Captura de tela da região a ser clicada salva como 'region_to_click.png'.")

            
                logger.info(f"clicando em: {x} {y}")
                actions = ActionChains(driver)
                actions.move_by_offset(x, y).click().perform()

                logger.info("Clicou no botão.")

                send_button = WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, 'span[data-icon="send"]'))
                )
                logger.debug('send button',send_button)
                actions.move_to_element(send_button).click().perform()

                time.sleep(15)
                logger.info("retry no clique.")
            else: 
                logger.info("O botão não foi encontrado na tela.")
        except Exception as e:
            logger.error("Failed to send message to %s due to %s", contact, e)
        finally:
            logger.info("Message finalizou")
            time.sleep(20)
            driver.quit()

readContacts("./contacts-message.xlsx")
