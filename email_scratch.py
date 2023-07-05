# Importing libraries
import imaplib
import email

def get_email_code():
    # from credentials.yml import user name and password
    my_credentials = {
        'user' : "asharirifky@gmail.com",
        'password' : "rivdxqlxcctkzkhm"
    }
    #Load the user name and passwd from yaml file
    user, password = my_credentials["user"], my_credentials["password"]

    #URL for IMAP connection
    imap_url = 'imap.gmail.com'

    # Connection with GMAIL using SSL
    my_mail = imaplib.IMAP4_SSL(imap_url)

    # Log in using your credentials
    my_mail.login(user, password)

    # Select the Inbox to fetch messages
    my_mail.select('Inbox')

    #Define Key and Value for email search
    #For other keys (criteria): https://gist.github.com/martinrusev/6121028#file-imap-search
    key = 'FROM'
    value = 'no-reply@canva.com'
    _, data = my_mail.search(None, key, value)  #Search for emails with specific key and value

    mail_id_list = data[0].split()  
    msgs = [] 
    for num in mail_id_list:
        typ, data = my_mail.fetch(num, '(RFC822)') #RFC822 returns whole message (BODY fetches just body)
        msgs.append(data)

    code = []

    for msg in msgs[::-1]:
        for response_part in msg:
            if type(response_part) is tuple:
                my_msg=email.message_from_bytes((response_part[1]))
                # print (my_msg['subject'])
                res = ''.join(filter(lambda i: i.isdigit(), my_msg['subject']))
                code.append(res)
    
    return code