from dotenv import load_dotenv
from os import environ,path,getcwd
from selenium.webdriver import Chrome
import time

envPath = path.join(getcwd(),'assets','.env_selenium')
load_dotenv(envPath)
webdriverPath = environ.get('WEBDRIVER')
webdriver = Chrome(webdriverPath)

def acessPage(path):
    webdriver.get(path)

def setForms(formCriteria):
    for key in formCriteria:
        webdriver.find_element_by_id(key).send_keys(formCriteria[key])

def arrayCheckClick(array,criteria):
     for item in array:
             try:
                 if item.text == criteria:
                     time.sleep(1)
                     item.click()
             except:
                 continue

def removeDuplicateDays(array):
    monthWeeks = False
    previousValue = 0
    outputArray = []
    for item in array:
        if int(item.text) < previousValue and monthWeeks:
            break

        if int(item.text) > 1 and monthWeeks == False:
            pass
        else:
            previousValue = int(item.text)
            monthWeeks = True
            outputArray.append(item)

    return outputArray

def setCalendar(calendarCriteria):
    iFrameLocation = webdriver.find_element_by_id("main").find_element_by_class_name("container-fluid")
    time.sleep(3)
    webdriver.switch_to.frame(iFrameLocation.find_element_by_tag_name("iframe"))

    calendarInput = webdriver.find_element_by_class_name("emulated-calendar__component")
    calendarInput.click()
    webdriver.find_element_by_class_name("react-calendar__navigation__label").click()
    webdriver.find_element_by_class_name("react-calendar__navigation__label").click()

    yearButtons = webdriver.find_elements_by_class_name("react-calendar__decade-view__years__year")
    arrayCheckClick(yearButtons,calendarCriteria["year"])

    monthButtons = webdriver.find_elements_by_class_name("react-calendar__year-view__months__month")
    arrayCheckClick(monthButtons,calendarCriteria["month"])

    daysButtons = webdriver.find_elements_by_class_name("react-calendar__month-view__days__day")
    daysButtons = removeDuplicateDays(daysButtons)
    arrayCheckClick(daysButtons,calendarCriteria["begin"])
    arrayCheckClick(daysButtons,calendarCriteria["end"])

    exportExcel = webdriver.find_element_by_class_name("MuiButton-outlined")
    exportExcel.click()

    time.sleep(2)
    closeModal = webdriver.find_element_by_class_name("jss79").find_element_by_tag_name("button")
    closeModal.click()

def clickButton(tag_id):
    webdriver.find_element_by_id(tag_id).click()

if __name__ == '__main__':
    acessPage(environ.get("LOGIN_URL"))
    setForms({
        "username":environ.get("EMAIL"),
        "password":environ.get("PASSWORD")
    })
    clickButton("kc-login")
    acessPage(environ.get("ORDER_URL"))
    setCalendar({
        "year":"2018",
        "month":"janeiro",
        "begin":"1",
        "end":"31"
    })
    webdriver.close()