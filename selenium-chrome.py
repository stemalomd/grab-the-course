from selenium import webdriver
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import pytesseract
import base64
from matplotlib import image as mpimg
import numpy
import cv2
from PIL import Image as pilImage

import sys
import time
from datetime import datetime
import os
import io


def delay(ms):
    time.sleep(ms / 1000)


def getMstime():
    return int(round(time.time() * 1000))


def getTimeStr():
    return datetime.now().strftime("%Y/%m/%d %H:%M:%S.%f")[:-3]


def testObjNotNoneToStr(obj):
    if obj:
        return "Yes"
    return "No"


def toClassSearchPage():
    driver.find_element(
        By.XPATH, '//*[@id="ctl00_MainContent_TabContainer1_tabCourseSearch_Label4"]'
    ).click()


def toSelectedPage():
    driver.find_element(
        By.XPATH, '//*[@id="ctl00_MainContent_TabContainer1_tabSelected_Label3"]'
    ).click()


def findEleFromClassSearchPage():
    toClassSearchPage()
    return driver.find_element(
        By.XPATH,
        '//*[@id="ctl00_MainContent_TabContainer1_tabCourseSearch"]',
    )


def findEleFromSelectedPage():
    toSelectedPage()
    return driver.find_element(
        By.XPATH,
        '//*[@id="ctl00_MainContent_TabContainer1_tabSelected"]',
    )


# 系統偵測異常 - 短時間內發出過量需求。請輸入驗證碼後，再執行加選動作，以確認使用狀態。
def isDetectionException():
    try:
        findEleFromSelectedPage().find_element(
            By.XPATH,
            f'.//*[@id="ctl00_MainContent_TabContainer1_tabSelected_CAPTCHA_imgCAPTCHA"]',
        )
        return True
    except:
        return False


def isCAPTCHANumberError():
    try:
        loginInfo = driver.find_element(
            By.XPATH,
            f'//*[@id="ctl00_Login1"]/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[5]/td',
        )
    except:
        return False

    if loginInfo.text != "驗證碼錯誤/Invalid Captcha":
        raise Exception()
    return True


def findClassOnFavoriteList(classCode: str):
    try:
        attentionClass = findEleFromSelectedPage().find_element(
            By.XPATH,
            f".//td[@class='gvAddWithdrawCellOne' and text()='{classCode}']/..",
        )
    except:
        raise Exception()
    return attentionClass


def isClassOnFavoriteList(classCode: str):
    try:
        findEleFromSelectedPage().find_element(
            By.XPATH,
            f".//td[@class='gvAddWithdrawCellOne' and text()='{classCode}']/..",
        )
        return True
    except:
        return False


def isClassCanAddOnFavoriteList(classCode: str):
    while 1:
        if isClassOnFavoriteList(classCode):
            break
        addClassToFavorite(classCode)
    return findClassOnFavoriteList(classCode).get_attribute("style") != "color: red;"


def searchClassAndAddOrDel(classCode: str, addOrDel: str):
    findPage = findEleFromSelectedPage
    # 已選課表搜尋框
    findPage().find_element(
        By.XPATH, './/*[@id="ctl00_MainContent_TabContainer1_tabSelected_tbSubID"]'
    ).send_keys(classCode)
    try:
        AddClassButton = findPage().find_element(
            By.XPATH,
            './/*[@id="ctl00_MainContent_TabContainer1_tabSelected_gvToAdd"]/tbody/tr[2]/td[1]/input',
        )
    except:
        AddClassButton = None
    try:
        DelClassButton = findPage().find_element(
            By.XPATH,
            './/*[@id="ctl00_MainContent_TabContainer1_tabSelected_gvToDel"]/tbody/tr[2]/td[1]/input',
        )
    except:
        DelClassButton = None
    if DelClassButton and addOrDel == "Del":
        DelClassButton.click()
        WebDriverWait(driver, 10).until(EC.alert_is_present())
        driver.switch_to.alert.accept()
    elif AddClassButton and addOrDel == "Add":
        AddClassButton.click()
        global actionCount
        actionCount += 1
    else:
        raise Exception()
    return


def addClassToFavorite(classCode: str):
    findPage = findEleFromClassSearchPage
    # 輸入
    searchBox = findPage().find_element(
        By.XPATH,
        './/*[@id="ctl00_MainContent_TabContainer1_tabCourseSearch_wcCourseSearch_tbSubID"]',
    )
    searchBox.clear()
    searchBox.send_keys(classCode)
    # 搜尋
    findPage().find_element(
        By.XPATH,
        './/*[@id="ctl00_MainContent_TabContainer1_tabCourseSearch_wcCourseSearch_btnSearchOther"]',
    ).click()
    # 確認搜尋結果
    result = findPage().find_element(
        By.XPATH,
        './/*[@id="ctl00_MainContent_TabContainer1_tabCourseSearch_wcCourseSearch_gvSearchResult"]/tbody/tr[2]/td[2]',
    )
    if result.text != classCode:
        raise Exception()
    button = findPage().find_element(
        By.XPATH,
        './/*[@id="ctl00_MainContent_TabContainer1_tabCourseSearch_wcCourseSearch_gvSearchResult_ctl02_btnAdd"]',
    )
    if button.get_attribute("value") != "關注":
        raise Exception()
    button.click()


def getSeatsLeftForClassOnFavoriteList(classCode: str):
    seatsLeftButton = findClassOnFavoriteList(classCode).find_elements(
        By.XPATH, ".//td[@class='gvAddWithdrawCellEight']/input"
    )[0]
    if seatsLeftButton.get_attribute("value") != "餘額查詢":
        print(f"餘額查詢按鈕冷卻中:{classCode}")
        return
    seatsLeftButton.click()
    WebDriverWait(driver, 15).until(EC.alert_is_present())
    confirm = driver.switch_to.alert
    text = confirm.text
    num = int(text[text.find("：") + 1 : text.rfind("/")])
    confirm.accept()
    return num


def addClassOnFavoriteList(classCode: str):
    classEle = findClassOnFavoriteList(classCode)
    if classEle.get_attribute("style") == "color: red;":
        raise Exception()
    classEle.find_element(By.XPATH, ".//td[1]/input").click()
    global actionCount
    actionCount += 1


def imgResize(img: numpy.ndarray, size=30000):
    h, w = img.shape[0], img.shape[1]
    mul = int((size / (h * w)) ** 0.5)
    h, w = int(h * mul), int(w * mul)
    return cv2.resize(img, (w, h))


def getBase64Img(xpath):
    img = driver.execute_async_script(
        f"""
        var canvas = document.createElement('canvas');
        var context = canvas.getContext('2d');
        var img = document.evaluate('{xpath}', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
        canvas.height = img.naturalHeight;
        canvas.width = img.naturalWidth;
        context.drawImage(img, 0, 0);
        
        callback = arguments[arguments.length - 1];
        callback(canvas.toDataURL());
        """
    )
    # 去掉多餘的內容
    return base64.b64decode(img.split(",")[1])


def getCAPTCHANumber(originalImg):
    bigImg = imgResize(originalImg)
    grayImg = cv2.cvtColor(bigImg, cv2.COLOR_BGR2GRAY)
    _, binarizationImg = cv2.threshold(grayImg, 130, 255, cv2.THRESH_BINARY)
    erosionImg = cv2.dilate(
        binarizationImg, numpy.ones((3, 3), numpy.uint8), iterations=2
    )
    dilationImg = cv2.erode(erosionImg, numpy.ones((3, 3), numpy.uint8), iterations=4)
    test_message = pilImage.fromarray(dilationImg)
    text = pytesseract.image_to_string(test_message)
    return text[:4]


def login():
    while 1:
        driver.get("https://course.fcu.edu.tw/")
        driver.find_element(By.XPATH, '//*[@id="ctl00_Login1_UserName"]').send_keys(
            "D0123456"
        )
        driver.find_element(By.XPATH, '//*[@id="ctl00_Login1_Password"]').send_keys(
            "abc123"
        )

        img = getBase64Img('//*[@id="ctl00_Login1_Image1"]')
        img = io.BytesIO(img)
        img = mpimg.imread(img, format="PNG")
        img = img * 255
        img = img.astype(numpy.uint8)

        number = getCAPTCHANumber(img)
        driver.find_element(By.XPATH, '//*[@id="ctl00_Login1_vcode"]').send_keys(number)
        driver.find_element(By.XPATH, '//*[@id="ctl00_Login1_LoginButton"]').click()

        if not isCAPTCHANumberError():
            break


def addClassStateTransition(i):
    global info, successNum
    while 1:
        mode = info[i]["模式"]
        oldCls = info[i]["舊課"]
        newCls = info[i]["新課"]
        state = info[i]["狀態"]
        newState = ""
        match state:
            case "舊課程":
                if mode == "舊換新":
                    # 新課程有座位
                    if getSeatsLeftForClassOnFavoriteList(newCls):
                        # 退選舊課程
                        searchClassAndAddOrDel(oldCls, "Del")
                        # 加選新課程
                        addClassOnFavoriteList(newCls)
                        # 加選新課程失敗
                        if isClassCanAddOnFavoriteList(newCls):
                            print(
                                f'加選新課程失敗 {newCls} {datetime.now().strftime("%Y/%m/%d %H:%M:%S.%f")[:-3]}'
                            )
                            newState = "新的選不到"
                        else:
                            newState = "新課程"
                    else:
                        newState = "舊課程"
                else:
                    raise Exception()
            case "新課程":
                if mode == "舊換新" or mode == "加選新":
                    newState = "新課程"
                    successNum += 1
                else:
                    raise Exception()
            case "舊的選不到":
                if mode == "舊換新":
                    addClassOnFavoriteList(newCls)
                    if isClassCanAddOnFavoriteList(newCls):
                        newState = "新的選不到"
                    else:
                        newState = "新課程"
                else:
                    raise Exception()
            case "新的選不到":
                if mode == "舊換新":
                    addClassOnFavoriteList(oldCls)
                    if isClassCanAddOnFavoriteList(oldCls):
                        newState = "舊的選不到"
                    else:
                        addClassToFavorite(oldCls)  # 下次加不到新課時要加回舊課所用
                        newState = "舊課程"
                elif mode == "加選新":
                    addClassOnFavoriteList(newCls)
                    if isClassCanAddOnFavoriteList(newCls):
                        newState = "新的選不到"
                    else:
                        newState = "新課程"
                else:
                    raise Exception()
            case _:
                raise Exception()

        info[i]["狀態"] = newState

        match state:
            case "舊課程":
                if mode == "舊換新":
                    if newState == "新的選不到":
                        pass
                    else:
                        break
                else:
                    raise Exception()
            case "新課程":
                break
            case "舊的選不到":
                break
            case "新的選不到":
                break
            case _:
                raise Exception()


def addClassProcess():
    global successNum, clsNum, actionCount, info
    global processStartTime, processActionTotal, processTenMinCount, processExceptCount
    maxActionNumInRound = 2
    actionTotal = 0
    # 將需要用到的課程加到關注清單
    for i in range(clsNum):
        oldCls = info[i]["舊課"]
        newCls = info[i]["新課"]
        mode = info[i]["模式"]
        if mode == "舊換新":
            if not isClassOnFavoriteList(oldCls):
                addClassToFavorite(oldCls)
            if not isClassOnFavoriteList(newCls):
                addClassToFavorite(newCls)
        elif mode == "加選新":
            if not isClassOnFavoriteList(newCls):
                addClassToFavorite(newCls)
        else:
            raise Exception()

    while 1:
        successNum = 0
        for i in range(clsNum):
            startT = getMstime()
            startState = info[i]["狀態"]
            actionCount = 0
            addClassStateTransition(i)
            if info[i]["狀態"] != startState:
                print(
                    f'{i} old:{info[i]["舊課"]:4} new:{info[i]["新課"]:4} s1:{startState:5} s2:{info[i]["狀態"]:5} {getTimeStr()}'
                )

            actionTotal += actionCount
            processActionTotal += actionCount

            nowT = getMstime()
            if nowT - processStartTime > 10 * 60 * 1000:
                processStartTime = nowT
                processTenMinCount += 1
                print(
                    f"過了第{processTenMinCount}次10分鐘 共 加選{processActionTotal}次,意外{processExceptCount}次 {getTimeStr()}"
                )

            if actionTotal + maxActionNumInRound > 30:
                return "需要重新登入"
            if isDetectionException():
                if actionTotal < 2:
                    return "意外狀況 遭系統阻擋"
                return "意外狀況 加選滿30次前跳驗證"

        if successNum == clsNum:
            return "已完成加選"


def main():
    global driver, processExceptCount
    options = Options()
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
    options.add_argument("headless")
    options.add_argument("disable-gpu")
    driver = webdriver.Chrome(
        service=Service(
            ChromeDriverManager(
                path=f"./drivers/{os.path.basename(__file__)}"
            ).install()
        ),
        options=options,
    )
    while 1:
        try:
            login()
            state = addClassProcess()
            if state == "需要重新登入":
                pass
            elif state == "意外狀況 加選滿30次前跳驗證":
                processExceptCount += 1
                pass
            elif state == "意外狀況 遭系統阻擋":
                processExceptCount += 1
                return state
            elif state == "已完成加選":
                print(f"{state} {getTimeStr()}")
                return state
            else:
                print(f"addClassProcess返回值錯誤 {getTimeStr()}")
                pass
        except:
            print(f"意外狀況 login或addClassProcess 發生錯誤 {getTimeStr()}")


if __name__ == "__main__":
    os.environ["WDM_LOG_LEVEL"] = "0"
    pytesseract.pytesseract.tesseract_cmd = (
        r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    )
    driver = None
    info = [
        {"模式": "加選新", "舊課": "0", "新課": "0123", "狀態": "新的選不到"},  # 舊課程,新課程,舊的選不到,新的選不到
    ]
    clsNum = len(info)
    actionCount = None
    successNum = None
    processStartTime = getMstime()
    processActionTotal = 0
    processTenMinCount = 0
    processExceptCount = 0
    while 1:
        try:
            state = main()
            if state == "意外狀況 遭系統阻擋":
                driver.quit()
            elif state == "已完成加選":
                break
            else:
                print(f"main返回值錯誤 {getTimeStr()}")
                driver.quit()
        except:
            print(f"main error {getTimeStr()}")
            if driver:
                driver.quit()
    driver.quit()
