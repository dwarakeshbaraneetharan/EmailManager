import os
import pickle
# Gmail API utils
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
# for encoding/decoding messages in base64
from base64 import urlsafe_b64decode, urlsafe_b64encode
# for dealing with attachment MIME types
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from mimetypes import guess_type as guess_mime_type
import PySimpleGUI as sg
import random
import string
import webbrowser
import pdfkit

# Request all access (permission to read/send/receive emails, manage the inbox, and more)
SCOPES = ['https://mail.google.com/']

sg.theme('Material2')


def gmail_authenticate():
    creds = None
    # the file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first time
    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as token:
            creds = pickle.load(token)
    # if there are no (valid) credentials availablle, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # save the credentials for the next run
        with open("token.pickle", "wb") as token:
            pickle.dump(creds, token)
    return build('gmail', 'v1', credentials=creds)


def search_messages(API_service, query):
    result = API_service.users().messages().list(userId='me', q=query).execute()
    messages = []
    if 'messages' in result:
        messages.extend(result['messages'])
    while 'nextPageToken' in result:
        page_token = result['nextPageToken']
        result = API_service.users().messages().list(userId='me', q=query, pageToken=page_token).execute()
        if 'messages' in result:
            messages.extend(result['messages'])
    return messages


def parse_parts(API_service, parts, folder_name, message, isDownloading):
    """
    Utility function that parses the content of an email partition
    """
    totalSize = 0
    if isDownloading:
        if "/" not in folder_name:
            folder_name = os.path.join("Downloads", folder_name)

        if not os.path.isdir(folder_name):
            try:
                os.mkdir(folder_name)
            except:
                os.mkdir(folder_name.split("/")[0])
                os.mkdir(folder_name)
    if parts:
        for part in parts:
            filename = part.get("filename")
            mimeType = part.get("mimeType")
            body = part.get("body")
            data = body.get("data")
            file_size = body.get("size")
            totalSize += file_size
            part_headers = part.get("headers")
            if part.get("parts"):
                # recursively call this function when we see that a part
                # has parts inside
                parse_parts(API_service, part.get("parts"), folder_name, message, isDownloading)
            if mimeType == "text/plain":
                # if the email part is text plain
                if data:
                    text = urlsafe_b64decode(data).decode()
                    print(text)
            elif mimeType == "text/html":
                # if the email part is an HTML content
                # save the HTML file and optionally open it in the browser

                if not filename:
                    filename = "index.html"
                if isDownloading:
                    if os.path.isfile(os.path.join(folder_name, filename)):
                        with open(os.path.join(folder_name, filename), 'r') as file:
                            html_content = file.read()
                        if urlsafe_b64decode(data).decode().strip() == html_content.strip():
                            print("ALREADY EXISTS")
                            continue

                        folder_counter = 0
                        while os.path.isfile(os.path.join(folder_name, filename)):
                            folder_counter += 1
                            folder_name = "{} ({})".format(folder_name, folder_counter)

                        os.mkdir(folder_name)

                    filepath = os.path.join(folder_name, filename)
                    print("Saving HTML to", filepath)
                    with open(filepath, "wb") as f:
                        f.write(urlsafe_b64decode(data))
                    print(folder_name)
                    try:
                        pdfkit.from_file(filepath,
                                         output_path=os.path.join(folder_name, folder_name.split("/")[1] + ".pdf"),
                                         options={"enable-local-file-access": ""})
                    except Exception:
                        pass
            else:
                # attachment other than a plain text or HTML
                for part_header in part_headers:
                    part_header_name = part_header.get("name")
                    part_header_value = part_header.get("value")
                    if part_header_name == "Content-Disposition":
                        if "inline" or "attachment" in part_header_value:
                            # we get the attachment ID
                            # and make another request to get the attachment itself
                            if isDownloading:
                                print("Saving the file:", filename, "size:", get_size_format(file_size))
                                attachment_id = body.get("attachmentId")
                                attachment = API_service.users().messages() \
                                    .attachments().get(id=attachment_id, userId='me', messageId=message['id']).execute()
                                data = attachment.get("data")
                                filepath = os.path.join(folder_name, filename)
                                if data:
                                    with open(filepath, "wb") as f:
                                        f.write(urlsafe_b64decode(data))
    return totalSize


# utility functions
def get_size_format(b, factor=1024, suffix="B"):
    """
    Scale bytes to its proper byte format
    e.g:
        1253656 => '1.20MB'
        1253656678 => '1.17GB'
    """
    for unit in ["", "K", "M", "G", "T", "P", "E", "Z"]:
        if b < factor:
            return f"{b:.2f}{unit}{suffix}"
        b /= factor
    return f"{b:.2f}Y{suffix}"


def clean(text):
    # clean text for creating a folder
    return "".join(c if (c.isalnum() or c == " " or c == "(" or c == ")") else "_" for c in text)


def read_message(API_service, message):
    print(type(message))
    """
    This function takes Gmail API `service` and the given `message_id` and does the following:
        - Downloads the content of the email
        - Prints email basic information (To, From, Subject & Date) and plain/text parts
        - Creates a folder for each email based on the subject
        - Downloads text/html content (if available) and saves it under the folder created as index.html
        - Downloads any file that is attached to the email and saves it in the folder created
    """
    msg = API_service.users().messages().get(userId='me', id=message['id'], format='full').execute()
    # parts can be the message body, or attachments
    payload = msg['payload']
    headers = payload.get("headers")
    parts = payload.get("parts")
    emailInfo = {"id": message, "parts": parts}
    if headers:
        # this section prints email basic info & creates a folder for the email
        for header in headers:
            name = header.get("name")
            value = header.get("value")
            if name.lower() == 'from':
                emailInfo["from"] = value
                # we print the From address
            if name.lower() == "to":
                emailInfo["to"] = value
                # we print the To address
            if name.lower() == "subject":
                if value != "":
                    folder_name = clean(value)
                    emailInfo["subject"] = value
                    emailInfo["folder_name"] = folder_name

            if name.lower() == "date":
                emailInfo["date"] = value
                # we print the date when the message was sent
    if "folder_name" not in emailInfo:
        emailInfo["date"] = " ".join(emailInfo["date"].split(" ")[: -1])
        folder_name = clean(emailInfo["date"])
        print(folder_name)
        # if the email does not have a subject, then make a folder with "email" name
        # since folders are created based on subjects
        emailInfo["folder_name"] = folder_name
    return emailInfo


# get the Gmail API service
service = gmail_authenticate()
allEmailIds = search_messages(service, "")
allEmails = []

for emailId in allEmailIds:
    emailList = read_message(service, emailId)
    emailList["total_size"] = (get_size_format(parse_parts(service, emailList["parts"], '', emailId, False)))
    allEmails.append(emailList)
    print("#"*50)
    print(emailList["date"], emailList["total_size"])
    print("#"*50)

# Characters used for the checked and unchecked checkboxes.
BLANK_BOX = '☐'
CHECKED_BOX = '☑'

# ------ Make the Table Data ------
dataTable = [["checkbox", "sender", "subject", "size"]]

for emailDict in allEmails:
    dataTable.append([BLANK_BOX, emailDict["from"], emailDict["folder_name"], emailDict["total_size"]])

print(dataTable)
headings = [str(dataTable[0][x]) + ' ..' for x in range(len(dataTable[0]))]
headings[0] = ''
selected = [i for i, row in enumerate(dataTable[1:][:]) if row[0] == CHECKED_BOX]
# ------ Window Layout ------
layout = [[sg.Button("download", disabled=True)],
          [sg.Table(values=dataTable[1:][:], headings=headings, auto_size_columns=False,
                    col_widths=[5, 25, 25], font="Helvetica 14",
                    justification='center', num_rows=20, key='-TABLE-',
                    selected_row_colors='red on yellow',
                    vertical_scroll_only=False,
                    enable_click_events=True),
           sg.Sizegrip()]]

# ------ Create Window ------
window = sg.Window('Table with Checkbox', layout, resizable=True, finalize=True)

# Highlight the rows (select) that have checkboxes checked
window['-TABLE-'].update(values=dataTable[1:][:], select_rows=list(selected))
window['-TABLE-'].expand(True, True)
window['-TABLE-'].table_frame.pack(expand=True, fill='both')
window.maximize()
# ------ Event Loop ------
while True:
    event, values = window.read()
    print(event, values)
    print(type(event))
    if event == sg.WIN_CLOSED:
        break
    if event == 'download':
        print("Download pressed, and table values are", values["-TABLE-"])
        for v in values["-TABLE-"]:
            parse_parts(service, allEmails[v]["parts"], allEmails[v]["folder_name"], allEmails[v]["id"], True)
    elif event[0] == '-TABLE-' and event[2][0] not in (
            None, -1):  # if clicked a data row rather than header or outside table
        row = event[2][0] + 1
        if dataTable[row][0] == CHECKED_BOX:  # Going from Checked to Unchecked
            selected.remove(row - 1)
            dataTable[row][0] = BLANK_BOX
        else:  # Going from Unchecked to Checked
            selected.append(row - 1)
            dataTable[row][0] = CHECKED_BOX
        window['-TABLE-'].update(values=dataTable[1:][:],
                                 select_rows=list(selected))  # Update the table and the selected rows
        if len(selected) > 0:
            window['download'].update(disabled=False)
        else:
            window['download'].update(disabled=True)

window.close()
