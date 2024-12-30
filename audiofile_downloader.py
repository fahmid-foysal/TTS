import time
from openpyxl import Workbook, load_workbook
from seleniumbase import Driver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys


    
class RunScrapper:
    def download_audios(self,downloadable_words):

        driver=Driver(uc=True)
        driver.get('https://www.narakeet.com/app/text-to-audio/?projectId=02152928-8c2f-4cf3-bacc-ce844ae2fee2')


        driver.click('body > header > nav > span:nth-child(3) > a.button.muted')
        driver.sleep(20)
        driver.click('#cfgVideoLanguage')
        driver.click('#cfgVideoLanguage > option:nth-child(20)')
        driver.click('#configureAudio > div > div.dialog-contents.fill > div.control-block > span > button')
        driver.click('#cfgAudioFormat')
        driver.click('#cfgAudioFormat > option:nth-child(2)')
        


        i=1
        for word in downloadable_words:
            
            driver.find_element(By.ID,"unparsedScriptEditor").send_keys(Keys.BACKSPACE+Keys.BACKSPACE+Keys.BACKSPACE+Keys.BACKSPACE+Keys.BACKSPACE+
                                                                        Keys.BACKSPACE+Keys.BACKSPACE+Keys.BACKSPACE+Keys.BACKSPACE+Keys.BACKSPACE+
                                                                        Keys.BACKSPACE+Keys.BACKSPACE+Keys.BACKSPACE+Keys.BACKSPACE+Keys.BACKSPACE+word)
            driver.sleep(2)
            driver.click('#workflowUploadForm > button.button.primary')
            driver.sleep(9)
            driver.click("body > main > div:nth-child(5) > div > div > div.actions.spaced.space-above > a:nth-child(3)")
            driver.sleep(6)
            driver.click('body > main > div:nth-child(5) > div > div > div.actions.spaced.space-above > a.button.back.muted')
            driver.sleep(4)
            print(i)
            i+=1
        driver.sleep(2)

