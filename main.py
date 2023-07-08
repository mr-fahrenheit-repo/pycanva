# Import Libraries
import pyautogui
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
    
# Locate element by x,y coordinate and RGB value
def locate_element(x,y,r,g,b):
    pyautogui.moveTo(x,y,1, pyautogui.easeOutQuad)
    while True:
        param = pyautogui.pixelMatchesColor(x, y, (r, g, b), tolerance=10)
        if param == True:
            pyautogui.click()
            break

# Locate element by x,y coordinate and RGB value with y axis adjustment     
def y_locate_element(x,y,r,g,b):
    x = x
    y = y 
    r = r
    g = g
    b = b
    while True:
        pyautogui.moveTo(x,y,1, pyautogui.easeOutQuad)
        hold(1,1.5)
        param = pyautogui.pixelMatchesColor(x, y, (r, g, b), tolerance=10)
        if param == True:
            pyautogui.click()
            break
        else:
            y = y - 20
            continue

# Hold between two commands
def hold(start, stop):
	time.sleep(round(random.uniform(float(start), float(stop)),2))

# Opening up Chrome browser
def openchrome():
    pyautogui.press('win')
    hold(0.75,1)
    pyautogui.write("chrome", interval = 0.12)
    hold(0.75,1)
    pyautogui.press("enter")

# Open the target URL
def opencanva():
    pyautogui.write(url, interval=0.25)
    hold(0.35,0.5)
    pyautogui.press("enter")
    hold(2,3)
    pyautogui.press("f11")
    
# Locate and click Accept Cookies prompt         
def allowcookies():
    locate_element(381,1009,173,200,231)

# Locate and click Log-in button   
def login(email, password):
    locate_element(1730,36,13,103,172)
    hold(1,2)
    y_locate_element(880,626,241,243,245)
    hold(0.5,0.75)
    y_locate_element(633,573,255,185,0)
    hold(4,5)
    pyautogui.press("f11")
    hold(1,2)
    pyautogui.write(email, interval=0.25)
    hold(1,2)
    pyautogui.press("enter")
    hold(5,7)
    pyautogui.write(password, interval=0.25)
    hold(0.5,0.75)
    pyautogui.press("enter")
    hold(2,3)
    pyautogui.press("enter")
   
# Configuring window 
def conf_windows():
    locate_element(395,708,241,243,245)
    hold(4,6)
    locate_element(12,150,24,25,27)
    hold(1,2)
    locate_element(1895,84,255,255,255)
    hold(0.75,1)
    
# Edit the design     
def edit_design():
    pyautogui.moveTo(1182,536,1, pyautogui.easeOutQuad)
    hold(0.35, 0.75)
    pyautogui.click()
    hold(0.35, 0.75)
    pyautogui.click()
    hold(0.35, 0.75)
    pyautogui.hotkey('ctrl','a')
    hold(0.75,1)
    pyautogui.write("Paijo", interval=0.15)
    hold(0.35, 0.75)
    locate_element(1151,131,235,236,240)
    
# Download the design
def download():
    pyautogui.moveTo(1851,15,1, pyautogui.easeOutQuad)
    locate_element(1851,15,255,255,255)
    pyautogui.moveTo(1703,536,1, pyautogui.easeOutQuad)
    locate_element(1703,536,241,243,245)
    pyautogui.moveTo(1763,402,1, pyautogui.easeOutQuad)
    locate_element(1763,402,115,0,230)
    hold(5,7)
    pyautogui.moveTo(1903,1052,1, pyautogui.easeOutQuad)
    pyautogui.click()
    hold(0.35, 0.75)
    locate_element(545,491,235,236,240)
    pyautogui.click()
    hold(1,3)

# Main Program
def main():
    print("*************************************** PyCanva v1.0.0 ***************************************")
    print("************************************* User Configuration *************************************")
    user = username()
    account_data = batch()
    df_log = log()
    print("************************************** Automation Start **************************************")
    print(colored(text="To stop the program from running you can press 'ctrl' + 'q' on your keyboard.", color="red"))
    for x in range(len(account_data)):
        date = time_stamp('date')
        start = time_stamp('time')
        account =  accounts(account_data)
        vpn(account['vpn'])
        hold(10,15)
        openchrome()
        hold(10,15)
        opencanva()
        hold(7,9)
        # allowcookies()
        hold(0.75,1.5)
        login(account['email'], account['password'])
        hold(7,9)
        conf_windows()
        hold(1,2)
        random_download = random.randint(75,150)
        for x in range(random_download):
            edit_design()
            hold(0.75,1)
            download()
            hold(1,2)
        end = time_stamp('time')
        input_log(x, df_log, user, date, start, end, account['email'], random_download)
    df_name = f"Download_log_{str(date)}_user_{user}"
    df_log.to_excel(f"{df_name}.xlsx")
    
# Running the code
if __name__ == '__main__':
    main()