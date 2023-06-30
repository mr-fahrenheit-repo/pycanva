import pyautogui
import time
import random

url = "https://www.canva.com/"

# random.randint(5,10)
list_name = ["Richie Norman",
"Eileen Wood",
"Arron Buckley",
"Ned Hansen",
"Jade Ingram",
"Fahad Ramirez",
"Lulu Pittman",
"Lilia Nguyen",
"Crystal Davies",
"Kallum Hicks"]

login = {
	"a" : 1183,
	"b" : 136,
	"c" : 477,
	"d" : 554,
	"e" : 478,
	"f" : 452
}

email_column = {
    "a" : 625,
    "b" : 452,
    "c" : 886,
    "d" : 409,
    "e" : 576,
    "f" : 323,
    "g" : 448,
    "h" : 771,
    "i" : 412
}

dashboard = {
	"a" : 1359,
	"b" : 574,
	"c" : 396
}

allowcook = {
	"a" : 108,
	"b" : 645
}

edit = {
    'a' : 888,
    "b" :393,
    "c" : 1341,
    "d" :189
}

randomflyer = {
    "a" : 1248,
    "b" : 165
}

email = "designasal@outlook.com"
password = "Karma111"


def openchrome():
    pyautogui.press('win')
    time.sleep(round(random.uniform(0.5,2.0),2))
    pyautogui.write("chrome", interval = 0.25)
    time.sleep(round(random.uniform(0.5,1.5),2))
    pyautogui.press("enter")

def opencanva(url):
	pyautogui.write(url, interval=0.25)
	pyautogui.press("enter")
 
def allowcookies(cookies_pic):
    pyautogui.click(pyautogui.locateCenterOnScreen(cookies_pic))

def loginbutton(log_in_pic, continue_another_way_pic):
    pyautogui.click(pyautogui.locateCenterOnScreen(log_in_pic))
    time.sleep(round(random.uniform(1.5,3.0),2))
    pyautogui.click(pyautogui.locateCenterOnScreen(continue_another_way_pic))
    time.sleep(round(random.uniform(1.5,3.0),2))

def loginemail(microsoft_pic, next_pic, login_pic, yes_pic):
    pyautogui.click(pyautogui.locateCenterOnScreen(microsoft_pic))
    time.sleep(round(random.uniform(10.0,15.0),2))
    pyautogui.write(email, interval=0.25)
    pyautogui.click(pyautogui.locateCenterOnScreen(next_pic))
    time.sleep(round(random.uniform(10.0,15.0),2))
    pyautogui.write(password, interval=0.25)
    pyautogui.click(login_pic)
    time.sleep(round(random.uniform(10.0,15.0),2))
    pyautogui.click()
    time.sleep(round(random.uniform(0.5,2.0),2))
    pyautogui.click(pyautogui.locateCenterOnScreen(yes_pic))

def opendesign(c):
	pyautogui.keyDown("down")
	time.sleep(7)
	pyautogui.keyUp('down')
	time.sleep(7)
	pyautogui.click(c,c)
    
def editdesign(a,b,c,d, elmnt_pic):
    elmnt = pyautogui.locateCenterOnScreen(elmnt_pic)
    pyautogui.click(elmnt)
    time.sleep(2)
    pyautogui.click(c,d)
    time.sleep(2)
    pyautogui.doubleClick(a,b)
    pyautogui.hotkey("ctrl","a", interval=0.25)
    pyautogui.write(random.choice(list_name), interval=0.15)
    
def download(a,b,share_pic, download_pic, download_png_pic):
    share = pyautogui.locateCenterOnScreen(share_pic)
    pyautogui.click(share)
    pyautogui.keyDown("down")
    time.sleep(2)
    pyautogui.keyUp("up")
    download = pyautogui.locateCenterOnScreen(download_pic)
    pyautogui.click(download)
    time.sleep(2)
    download_png = pyautogui.locateCenterOnScreen(download_png_pic)
    pyautogui.click(download_png)
    time.sleep(10)
    pyautogui.click(a,b)

openchrome()
time.sleep(round(random.uniform(5.0,7.0),2))
opencanva(url)
time.sleep(round(random.uniform(5.0,7.0),2))
allowcookies("accept_all_cookies.png")
time.sleep(round(random.uniform(5.0,7.0),2))
loginbutton("log_in.png", "continue_another_way.png")
time.sleep(round(random.uniform(5.0,7.0),2))
loginemail("continue_microsoft_way.png", "next_login.png", "login.png", "yes.png")





# time.sleep(5)
# opencanva(url,searchx,searchy)
# time.sleep(13)
# allowcookies(allowcook["a"], allowcook["b"])
# time.sleep(5)
# loginbutton(login['a'],login['b'], "other_way.png")
# time.sleep(5)
# loginemail("microsoft.png", "next_login.png", "login.png", "yes.png")
# time.sleep(15)
# opendesign(dashboard["c"])
# time.sleep(15)
# for i in range(10):
#     editdesign(edit["a"], edit["b"],edit["c"], edit["d"],"elmnt.png")
#     time.sleep(3)
#     download(randomflyer["a"],randomflyer["b"],"share.png", "down.png", "dwnld.png")