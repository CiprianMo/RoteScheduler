from selenium import webdriver
import time 
import os
import urllib

def convertPdfToExcel(path):
    ''' 
    Takes a path to a pdf file  
    
    Returns a tuple containing the path to the newly created
    data file as well as the resulting HTTPMessage object. 
    '''
    browser = webdriver.Firefox()

    browser.get('https://www.ilovepdf.com/pdf_to_excel')

    time.sleep(1)

    linkElement = browser.find_element_by_xpath("//div[@class='moxie-shim moxie-shim-html5']/input")

    linkElement.send_keys(os.getcwd()+'/Shiftlist-Oct2018(1).pdf')

    time.sleep(1)

    uploadElem = browser.find_element_by_id("uploadfiles")

    uploadElem.click()

    time.sleep(1)
    
    downld = browser.find_element_by_id("download")

    link = downld.get_attribute("href")

    result = urllib.request.urlretrieve(link, os.getcwd()+"/file.xlsx")

    browser.quit()

    return result


