# to access your credentials (optional)
import email_credential
# to access emails using IMAP email protocol 
import imaplib
# to access and manipulate the emails
import email
# to manipulate the files on the local desktop
import os
# to get D0
import datetime as dt

def createConnection(email, password, imapurl='imap.gmail.com'):
    '''
    INPUT: email and password
    OUTPUT: IMAP connection
    '''
    # this is done to make SSL connection with gmail
    connection = imaplib.IMAP4_SSL(imapurl)
    # login to the gmail account
    connection.login(email, password)
    # this is dont to check for emails under this label
    connection.select('Inbox') # 'Sent'
    # return created connection
    return connection

def searchMail(connection, Date, savepath=''):
    '''
    INPUT: connection, filter criteria (one or more), savepath (optional)
    OUTPUT: List containing a string for every line of the message body
    FUNCTIONALITY:
    Search email with multiple criteria (Can be Subject, Date, From, To and BBC)
    If input contains valid savepath, attachments will be downloaded to savepath
    '''
    # Get mails considering only one filter
    try:
        result, msgnums = connection.search(None,f'(DATE "{Date}")')
    except:
        raise Exception('Invalid input date')
    # Iterate through all mails returned from first search
    for msgnum in msgnums[0].split():
        # Get mail message
        result, raw_data = connection.fetch(msgnum, "(RFC822)")
        data = raw_data[0][1].decode('utf-8')
        message = email.message_from_string(data)
        '''
        # Apply all filters to skip unwanted messages
        if Date and (Date not in message.get("Date")): continue
        if Subject and (Subject not in message.get("Subject")): continue
        if From and (From not in message.get("From")): continue
        if To and (To not in message.get("To")): continue
        if BCC and (BCC not in message.get("BCC")): continue
        '''
        # Walk through the whole message
        for part in message.walk():
            # Skip useless message parts
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue
            # Get filename
            fileName = part.get_filename()
            # Save file in savepath
            if bool(fileName) and savepath:
                filePath = os.path.join(savepath, fileName)
                fp = open(filePath, 'wb')
                fp.write(part.get_payload(decode=True))
                fp.close()
                print(f'Downloaded "{fileName}" from email titled "{message.get("Subject")} at {savepath}"')


#USAGE EXAMPLE:

# Get D0 date
today = dt.date.today()
print(f"Running for {today}")

# Format date to use it in filepath
formated_date = str(today).replace('-', '')

# Format date to use as search parameter
month_dict = {'1': 'Jan', '2': 'Fev', '3': 'Mar', '4': 'Apr', '5': 'May', '6': 'June', '7': 'July', '8': 'Aug','9': 'Sep', '10': 'Oct', '11': 'Nov', '12': 'Dec'}
search_date = f'{today.day} {today.month} {today.year}'

# Connects to imap
connection = createConnection(email_credential.email, email_credential.password)

# Search mail with multiple criteria (Can be Subject, Date, From, To and BBC) (if input a savepath, attachments will be downloaded)
mail = searchMail(connection, Date='7 Oct 2022' ,savepath='H:/Teste/emailData/formated_date')