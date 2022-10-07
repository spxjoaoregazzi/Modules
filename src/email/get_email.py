import email_credential
# to access emails using IMAP email protocol 
import imaplib
# to access and manipulate the emails
import email
# to manipulate the files on the local desktop
import os

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

def searchMail(connection, Subject=None, Date=None, From=None, To=None, BCC=None, savepath=''):
    '''
    INPUT: connection, filter criteria (one or more), savepath (optional)
    OUTPUT: List containing a string for every line of the message body
    FUNCTIONALITY:
    Search email with multiple criteria (Can be Subject, Date, From, To and BBC)
    If input contains valid savepath, attachments will be downloaded to savepath
    '''
    # Get mails considering only one filter
    if Subject: result, msgnums = connection.search(None,f'(SUBJECT "{Subject}")')
    elif Date: result, msgnums = connection.search(None,f'(DATE "{Date}")')
    elif From: result, msgnums = connection.search(None,f'(FROM "{From}")')
    elif To: result, msgnums = connection.search(None,f'(TO "{To}")')    
    elif BCC: result, msgnums = connection.search(None,f'(BCC "{BCC}")')
    # Iterate through all mails returned from first search
    for msgnum in msgnums[0].split():
        # Get mail message
        result, raw_data = connection.fetch(msgnum, "(RFC822)")
        data = raw_data[0][1].decode('utf-8')
        message = email.message_from_string(data)
        # Apply all filters to skip unwanted messages
        if Date and (Date not in message.get("Date")): continue
        if Subject and (Subject not in message.get("Subject")): continue
        if From and (From not in message.get("From")): continue
        if To and (To not in message.get("To")): continue
        if BCC and (BCC not in message.get("BCC")): continue
        # Print data of filtered message
        print(f'Message number: {msgnum}')
        print(f'From: {message.get("From")}')
        print(f'To: {message.get("To")}')
        print(f'BCC: {message.get("BCC")}')
        print(f'Date: {message.get("Date")}')
        print(f'Subject: {message.get("Subject")}')
        print("Content:")
        # Create list that body content text/plain be appended
        body_content = []
        # Walk through the whole message
        for part in message.walk():
            # If this part of the message is plain text, append to body_content
            if part.get_content_type() == 'text/plain':
                body_content.append(part.get_payload(decode=False))
                print(part.get_payload(decode=False))
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
        return body_content

'''
USAGE EXAMPLE:
# Connects to imap
connection = createConnection(email_credential.email, email_credential.password)
# Search mail with multiple criteria (Can be Subject, Date, From, To and BBC) (if input a savepath, attachments will be downloaded)
mail = searchMail(connection, Subject='SPX EoD Tradefile - Fixed Income Offshore', Date='6 Oct 2022', savepath='H:/Teste/emailData')
print(mail)
'''