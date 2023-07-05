import time
import pandas as pd
from termcolor import colored
import random
# user, date, start, end, account, download
# def log():
#     columns = ['USER', 'date','start', 'end', 'account', 'download']
#     data_dict = {}
#     df = pd.DataFrame(data_dict, columns=columns)
#     return df

# df = log()
# df.at[0, 'USER'] = "rifky"
# df.at[0, 'download'] = 50
# df.at[1, 'download'] = 50
# print(df)


def logging(numb_index, df, user, date, start, end, account, download):
    df.at[numb_index, 'USER'] =  user
    df.at[numb_index, 'date'] =  date
    df.at[numb_index, 'start'] =  start
    df.at[numb_index, 'end'] =  end
    df.at[numb_index, 'account'] =  account
    df.at[numb_index, 'download'] =  download

# import datetime

# def time_stamp(options):
#     option = options
    
#     if option == "time":
#         current_time = datetime.datetime.now()
#         formatted_time = current_time.strftime("%H:%M %p")
#         return formatted_time
#     elif option == "date":
#         current_time = datetime.datetime.now()
#         formatted_time = current_time.strftime("%Y-%m-%d")
#         return formatted_time


# print(time_stamp('time'))
# print(time_stamp('date'))

# Loading account dataframe
def batch():
    while True:
        print("Please provide the path to the account database!")
        dbs_input = input()
        dbs = dbs_input.replace('"', '')
        print("The provided database path is ===>", colored(dbs, 'red', attrs=['reverse', 'blink']))
        time.sleep(1)
        print("Are you sure?\n[y] - Yes\n[n] - No")
        confirm = input()
        if confirm == 'y':
            print("Processing...")
            time.sleep(1)
            try:
                df = pd.read_excel(dbs)
                if set(df.columns) == set(['email', 'password', 'country', 'vpn_code', 'premium']):
                    print("***************** DONE YA BANG! *****************")
                    break
                else:
                    print("Please provide the correct database!")
                    pass
            except:
                print("Please provide the correct database!")
            else:
                print("***************** DONE YA BANG! *****************")
                break
        elif confirm == 'n':
            pass
        else:
            print("Please provide the correct answer!")
            time.sleep(2)
            pass
    df = df.reset_index()
    return df

account_data = batch()

def account(df):
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



account = account(account_data)

print(account['email'])
print(account['password'])
print(account['vpn'])