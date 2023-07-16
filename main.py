# Import Libraries
import pyautogui
pyautogui.FAILSAFE = False
import time
import random
import os
import glob
import pandas as pd
import datetime
from termcolor import colored

# Target URL for the automation 
url = "www.canva.com/"

# Time and date stamp for logging
def time_stamp(options):
    option = options
    
    if option == "time":
        current_time = datetime.datetime.now()
        formatted_time = current_time.strftime("%H:%M %p")
        return formatted_time
    elif option == "date":
        current_time = datetime.datetime.now()
        formatted_time = current_time.strftime("%Y-%m-%d")
        return formatted_time

# Create Log dataframe
def log():
    columns = ['USER', 'date','start', 'end', 'account', 'download']
    data_dict = {}
    df = pd.DataFrame(data_dict, columns=columns)
    return df

# Inputing the logging data
def input_log(numb_index, df, user, date, start, end, account, download):
    df.at[numb_index, 'USER'] =  user
    df.at[numb_index, 'date'] =  date
    df.at[numb_index, 'start'] =  start
    df.at[numb_index, 'end'] =  end
    df.at[numb_index, 'account'] =  account
    df.at[numb_index, 'download'] =  download

# Loading account dataframe
def batch():
    print(colored(text="Scanning databases in directory", color="green"))
    directory = os.getcwd()
    search_pattern = os.path.join(directory, f"*.xlsx")
    files = glob.glob(search_pattern)
    while True:
        for i,file in enumerate(files):
            substrings = file.split("\\")
            time.sleep(0.5)
            print(f"[{i}] -",substrings[-1])
        index = input(colored(text="Please provide the account database to be executed! ", color="yellow"))
        try:
            selected_files = files[int(index)]
        except:
            print("Please provide the correct database")
            pass
        else:
            try:
                df = pd.read_excel(selected_files)
                if set(df.columns) == set(['email', 'password', 'country', 'vpn_code', 'premium']):
                    print("**************************************** DONE YA BANG! ***************************************\n")
                    break
                else:
                    print("Please provide the correct database!")
                    pass
            except:
                print("Please provide the correct database!")
                pass
    return df

# Random choose account from account dataframe
def accounts(df):
    email_account = random.choice(df['email'])
    index = df.index[df['email'] ==  email_account][0]
    password_account = df.loc[index, 'password']
    vpn_account = df.loc[index, 'vpn_code']
    account = {
        'email' : email_account,
        'password' : password_account,
        'vpn' : vpn_account
    }
    return account

# Username input 
def username():
    print(colored(text="Before starting the program you have to choose username based on options provided down below.", color="green"))
    time.sleep(0.75)
    print("[0] - Rifky\n[1] - Fitra")
    time.sleep(0.75)
    while True:
        user = int(input(colored(text="Please enter your username! ", color="yellow")))
        if user == 0:
            user = "Rifky"
            print(f"Welcome {user}")
            print("Press any key to continue or 0 to start over...")
            start = input()
            if start == '0':
                pass
            else:
                print("**************************************** DONE YA BANG! ***************************************\n")
                break
        elif user == 1:
            user = "Fitra"
            print(f"Welcome {user}")
            print("Press any key to continue or 0 to start over...")
            start = input()
            if start == '0':
                pass
            else:
                print("**************************************** DONE YA BANG! ***************************************\n")
                break
        else:
            print("Please input the correct username!")
            pass
        
    return user

# Changing vpn and flushing DNS
def vpn(code):
    os.system('"ipconfig /flushdns"')
    os.system(f'"mullvad relay set hostname {code}"')
    os.system('"mullvad connect"')
    hold(10)
    os.system('"mullvad status"')
    print("===================================================")

# Hold between two commands
def hold(stop):
    start = stop * 0.7
    time.sleep(round(random.uniform(float(start), float(stop)),2))

# Checking for element loaded
def check_element(x_1, x_2, y, r, g, b):
    for tries in range(50):
        x_random = random.randint(x_1,x_2)
        pyautogui.moveTo(x_random,y,1, pyautogui.easeOutQuad)
        param = pyautogui.pixelMatchesColor(x_random, y, (r, g, b), tolerance=10)
        if param == True:
            break
        else:
            pass

# Checking for element loaded
def check_element_msft(x, y, r, g, b):
    hold(2)
    for tries in range(50):
        xp = x + 10
        y = y
        pyautogui.moveTo(x,y,1, pyautogui.easeOutQuad)
        param_x = pyautogui.pixelMatchesColor(x, y, (r, g, b), tolerance=10)
        param_xp = pyautogui.pixelMatchesColor(xp, y, (r, g, b), tolerance=10)
        if param_x == True:
            break
        elif param_xp == True:
            break
        else:
            y = y - 10
            continue


# Pick the login button      
def pick_login(x,y,r,g,b):
    y = y 
    for tries in range(10):
        pyautogui.moveTo(x,y,0.65, pyautogui.easeOutQuad)
        hold(1)
        param = pyautogui.pixelMatchesColor(x, y, (r, g, b), tolerance=10)
        if param == True:
            pyautogui.click()
            break
        else:
            y = y - 57
            
            continue

# Locate element by x,y coordinate and RGB value
def locate_element(x,y,r,g,b):
    for i in range(5):
        pyautogui.moveTo(x,y,1, pyautogui.easeOutQuad)
        param = pyautogui.pixelMatchesColor(x, y, (r, g, b), tolerance=50)
        if param == True:
            pyautogui.click()
            break
        else:
            x_random = random.randint(100,1820)
            y_random = random.randint(100,980)
            pyautogui.moveTo(x_random,y_random,0.5, pyautogui.easeOutQuad)
            pass

# Locate element by x,y coordinate and RGB value with y axis adjustment     
def y_locate_element(x,y,r,g,b):
    y = y 
    for tries in range(10):
        pyautogui.moveTo(x,y,0.4, pyautogui.easeOutQuad)
        param = pyautogui.pixelMatchesColor(x, y, (r, g, b), tolerance=10)
        if param == True:
            pyautogui.click()
            break
        else:
            y = y - 10
            continue

# Opening up Chrome browser
def openchrome():
    print("1. Opening chrome browser")
    pyautogui.press('win')
    hold(1)
    pyautogui.write("chrome", interval = 0.125)
    hold(1)
    pyautogui.press("enter")
    hold(6)

# Open the target URL
def opencanva():
    print("2. Opening Canva")
    pyautogui.write(url, interval=0.25)
    hold(0.5)
    pyautogui.press("enter")
    hold(1)
    pyautogui.press("f11")
    check_element(1795,1861,23,115,0,230)

# Locate and click Accept Cookies prompt         
def allowcookies():
    print("3. Cookies prompt")
    x_random = random.randint(312,445)
    locate_element(x_random,1006,173,200,231)
    
# Locate and click Log-in button   
def login(email, password):
    print("4. Logging In")
    x_random = random.randint(1700,1765)
    locate_element(x_random,24,225,228,231)
    hold(1)
    check_element(849,915,405,225,228,231)
    x_random = random.randint(849,915)
    pick_login(x_random,644,242,243,245) 
    hold(0.75)
    y_locate_element(633,573,255,185,0)
    hold(1)
    pyautogui.press("f11")
    check_element_msft(788,431,255,185,0)
    # check_element(798,804,383,255,185,0)
    hold(1)
    pyautogui.write(email, interval=0.25)
    hold(0.75)
    pyautogui.press("enter")
    hold(1)
    check_element_msft(788,431,255,185,0)
    pyautogui.write(password, interval=0.25)
    hold(0.75)
    pyautogui.press("enter")
    hold(1)
    check_element_msft(788,431,255,185,0)
    x_random_list = [random.randint(1032,1129),random.randint(921,1018)]
    y_random = random.randint(628,649)
    pyautogui.moveTo(random.choice(x_random_list,),y_random)
    pyautogui.click()
    hold(2)
    check_element(1682,1807,18,115,0,230)
    hold(5)

# Configuring window 
def conf_windows():
    print("5. Configuring windows for editing")
    locate_element(395,708,241,243,245)
    check_element(5,12,588,24,25,27)
    hold(2)
    y_random_list = [589,518,444,376,295,234,150]
    y_random_choice = random.choice(y_random_list)
    locate_element(12,y_random_choice,24,25,27)
    hold(1)
    locate_element(1895,84,255,255,255)
    hold(1)
    
# Random choosing name from given file
def random_name():
    with open('name list.txt') as input_file:
        long_list = [line.strip() for line in input_file]
    random_line = random.choice(long_list)
    return random_line

# Edit the design     
def edit_design(number):
    print(f"Editing design #{number}")
    pyautogui.moveTo(1182,536,1, pyautogui.easeOutQuad)
    hold(0.75)
    pyautogui.click()
    hold(0.75)
    pyautogui.click()
    hold(0.75)
    pyautogui.hotkey('ctrl','a')
    hold(1)
    pyautogui.write(random_name(), interval=0.15)
    hold(0.75)
    random_x = random.randint(641,1537)
    locate_element(random_x,131,235,236,240)

# Search for download button
def search_download(number):
    for x in range(number):
        pyautogui.press('tab')
        hold(0.5)
    pyautogui.press('enter')

# Download the design
def download(number):
    print(f"Downloading design #{number}")
    pyautogui.moveTo(1851,15,1, pyautogui.easeOutQuad)
    locate_element(1851,15,255,255,255)
    hold(2)
    search_download(10)
    hold(2)
    search_download(7)
    hold(13)
    pyautogui.moveTo(1902,1051,1, pyautogui.easeOutQuad)
    pyautogui.click()
    hold(1)
    x_random = random.randint(437,610)
    y_random = random.randint(117,1006)
    pyautogui.moveTo(x_random,y_random,1, pyautogui.easeOutQuad)
    pyautogui.click()
    hold(1)

# Closing the chrome web browser
def closechrome():
    pyautogui.press("f11")
    hold(2)
    pyautogui.moveTo(1897,12,1, pyautogui.easeOutQuad)
    pyautogui.click()

# Main Program
def main():
    print("*************************************** PyCanva v1.0.0 ***************************************")
    print("************************************* User Configuration *************************************")
    user = username()
    account_data = batch()
    df_log = log()
    print("************************************** Automation Start **************************************")
    print(colored(text="To stop the program from running you can press 'ctrl' + 'c' on your keyboard.", color="red"))
    for x in range(len(account_data)):
        date = time_stamp('date')
        start = time_stamp('time')
        account =  accounts(account_data)
        vpn(account['vpn'])
        openchrome()
        opencanva()
        allowcookies()
        login(account['email'], account['password'])
        conf_windows()
        random_download = random.randint(33,35)
        for i in range(random_download):
            edit_design(i + 1)
            download(i + 1 )
        closechrome()
        end = time_stamp('time')
        input_log(x, df_log, user, date, start, end, account['email'], random_download)
    df_name = f"Download_log_{str(date)}_user_{user}"
    df_log.to_excel(f"{df_name}.xlsx")
    
# Running the code
if __name__ == '__main__':
    main()
