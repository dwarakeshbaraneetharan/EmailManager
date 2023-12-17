import datetime
import os
import pickle
# Gmail API utils
import time

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
import subprocess
import pdfkit

# Request all access (permission to read/send/receive emails, manage the inbox, and more)
SCOPES = ['https://mail.google.com/', "https://www.googleapis.com/auth/drive.metadata.readonly"]

sg.theme('Material2')

totalStorageUsed = 0


def gmail_authenticate():
    global totalStorageUsed
    creds = None
    # the file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first time
    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as token:
            creds = pickle.load(token)
    # if there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # save the credentials for the next run
        with open("token.pickle", "wb") as token:
            pickle.dump(creds, token)
    result = build('drive', 'v3', credentials=creds).about().get(fields="*").execute()
    result = result.get("storageQuota", {})
    totalStorageUsed = int(result["usage"])
    print(get_size_format(int(result["usage"])), result["limit"])
    return build('gmail', 'v1', credentials=creds), build('drive', 'v3', credentials=creds)


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


def get_downloaded():
    htmlPaths = []
    downloadedIds = []
    for (dir_path, dir_names, file_names) in os.walk("Downloads"):
        if "index.html" in file_names:
            htmlPaths.append(dir_path + "/index.html")
    for path in htmlPaths:
        with open(path, 'r') as htmlReader:
            downloadedIds.append((htmlReader.readlines()[0].split("--")[1]))
    return downloadedIds


def parse_parts(API_service, parts, folder_name, message, isDownloading):
    """
    Utility function that parses the content of an email partition
    """
    if message["id"] in get_downloaded():
        isDownloading = False

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
                    htmldata = "<!--{}-->".format(message["id"]) + urlsafe_b64decode(data).decode()
                    with open(filepath, "w+") as f:
                        f.writelines(htmldata)
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


def delete_message(API_service, message):
    try:
        API_service.users().messages().delete(userId='me', id=message['id']).execute()
    except Exception:
        pass


def read_message(API_service, message):
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
        # if the email does not have a subject, then make a folder with "email" name
        # since folders are created based on subjects
        emailInfo["folder_name"] = folder_name
    return emailInfo


# get the Gmail API service
service, drive_service = gmail_authenticate()


def load_storage():
    result = drive_service.about().get(fields="*").execute()
    result = result.get("storageQuota", {})
    return int(result["usage"]), int(result["limit"])


def search_and_load(query, sorter):
    BLANK_BOX = '☐'

    allEmailIds = search_messages(service, query)
    for eId in allEmailIds:
        emailList = read_message(service, eId)
        totalSize = parse_parts(service, emailList["parts"], '', eId, False)
        if totalSize == 0:
            allEmailIds.remove(eId)
    allEmails = []

    rowColors = []
    for emailId in allEmailIds:
        emailList = read_message(service, emailId)
        totalSize = parse_parts(service, emailList["parts"], '', emailId, False)
        emailList["raw_size"] = totalSize
        emailList["total_size"] = (get_size_format(totalSize))
        if emailId["id"] in get_downloaded():
            emailList["isDownloaded"] = True
            rowColors.append((allEmailIds.index(emailId), "light gray"))
        else:
            emailList["isDownloaded"] = False
            rowColors.append((allEmailIds.index(emailId), "white"))
        allEmails.append(emailList)

    if sorter == "size":
        allSorted = []
        sortedColors = []
        while len(allEmails) > 0:
            largestSize = -1
            largestEmail = None
            for e in allEmails:
                if e["raw_size"] > largestSize:
                    largestSize = e["raw_size"]
                    largestEmail = e
            allSorted.append(largestEmail)
            sortedColors.append((len(sortedColors), rowColors.pop(allEmails.index(largestEmail))[1]))

            allEmails.remove(largestEmail)
        allEmails = allSorted[:]
        rowColors = sortedColors[:]

    # ------ Make the Table Data ------
    dataTable = [["checkbox", "sender", "subject", "size"]]

    for emailDict in allEmails:
        dataTable.append([BLANK_BOX, emailDict["from"], emailDict["folder_name"], emailDict["total_size"]])

    return allEmails, rowColors, dataTable


allEmails, rowColors, dataTable = search_and_load('', "default")

# Characters used for the checked and unchecked checkboxes.
BLANK_BOX = '☐'
CHECKED_BOX = '☑'

print(get_downloaded())

headings = [str(dataTable[0][x]) + ' ..' for x in range(len(dataTable[0]))]
headings[0] = ''
selected = []
# ------ Window Layout ------
layout = [[sg.Button("download", disabled=True, disabled_button_color='gainsboro', button_color='lime green',
                     font='Helvetica 14'),
           sg.Button("select all", font='Helvetica 14'), sg.Button("clear selected", font='Helvetica 14'),
           sg.Button("open saving directory", font='Helvetica 14'),
           sg.Button("delete", disabled=True, font='Helvetica 14', disabled_button_color='gainsboro',
                     button_color='indian red'),
           sg.Column(
               [[sg.Input(key="-SEARCH-", background_color='light gray', text_color='blue', font='Helvetica 14'),
                 sg.Button("search", font='Helvetica 14')]],
               element_justification="center",
               expand_x=True,
               key="c1",
               pad=(0, 0),
           ),
           sg.Column(
               [[sg.Text(
                   "{} of {} used".format(get_size_format(load_storage()[0]), get_size_format(load_storage()[1])),
                   key="-STORAGE DISPLAY-"),
                   sg.ProgressBar(key='-DRIVE-', max_value=load_storage()[1], orientation='h', size=(20, 20),
                                  bar_color=("deep sky blue", "light gray"))]],
               element_justification='right'
           )],
          [sg.Table(values=dataTable[1:][:], headings=headings, auto_size_columns=False,
                    col_widths=[5, 50, 50], font="Helvetica 18", row_colors=rowColors,
                    justification='center', num_rows=20, key='-TABLE-',
                    selected_row_colors='red on yellow',
                    # vertical_scroll_only=False,
                    enable_click_events=True),
           sg.Sizegrip()],
          [sg.ProgressBar(key='-DOWNLOAD BAR-', max_value=10, orientation='h', size=(40, 20),
                         bar_color=("deep sky blue", "light gray"), visible=False)]]

# ------ Create Window ------
window = sg.Window('Table with Checkbox', layout, resizable=True, finalize=True)

# Highlight the rows (select) that have checkboxes checked
window['-TABLE-'].update(values=dataTable[1:][:], select_rows=list(selected))
window['-TABLE-'].expand(True, True)
window['-TABLE-'].update(row_colors=rowColors)
window['-TABLE-'].table_frame.pack(expand=True, fill='both')
window['-DRIVE-'].UpdateBar(totalStorageUsed)
window.maximize()
firstClickOccurred = False
firstClickTime = 0
currentBar = 0
# ------ Event Loop ------
while True:
    event, values = window.read()
    print(event, values)
    if '+CLICKED+' in event and -1 in event[2] and 0 not in event[2]:
        if firstClickOccurred:
            timeDiff = int(
                (datetime.datetime.utcnow() - datetime.datetime(1970, 1, 1)).total_seconds() * 1000) - firstClickTime
            if timeDiff < 500:
                print(dataTable)
                if event[2][1] == 3:
                    allEmails, rowColors, dataTable = search_and_load(values['-SEARCH-'], "size")
                    window['-TABLE-'].update(values=dataTable[1:][:],
                                             select_rows=list(selected),
                                             row_colors=rowColors)  # Update the table and the selected rows
                firstClickOccurred = False
            else:
                firstClickOccurred = False
        else:
            firstClickOccurred = True
            firstClickTime = int(
                (datetime.datetime.utcnow() - datetime.datetime(1970, 1, 1)).total_seconds() * 1000)
    if event == sg.WIN_CLOSED:
        break
    if event == 'search':
        allEmails, rowColors, dataTable = search_and_load(values['-SEARCH-'], "default")
        window['-TABLE-'].update(values=dataTable[1:][:],
                                 select_rows=list(selected),
                                 row_colors=rowColors)  # Update the table and the selected rows
    if event == 'delete':
        if sg.Window("Confirm Deletion", [[sg.Text("Are you sure you want to delete these {} emails from your Google "
                                                   "Drive? This is final!".format(len(selected)), font="Helvetica 16")],
                                          [sg.Yes(), sg.No()]]).read(close=True)[0] == "Yes":
            selected.clear()
            for v in values["-TABLE-"]:
                delete_message(service, allEmails[v]["id"])
                totalStorageUsed -= allEmails[v]["raw_size"]
            window['-STORAGE DISPLAY-'].update(
                "{} of {} used".format(get_size_format(totalStorageUsed), get_size_format(load_storage()[1])))
            allEmails, rowColors, dataTable = search_and_load(values['-SEARCH-'], "default")
            window["-DRIVE-"].UpdateBar(totalStorageUsed)
            window['-TABLE-'].update(values=dataTable[1:][:],
                                     select_rows=list(selected),
                                     row_colors=rowColors)  # Update the table and the selected rows
    if event == 'download':
        print("Download pressed, and table values are", values["-TABLE-"])
        window['-DOWNLOAD BAR-'].update(visible=True, current_count=currentBar, max=len(values["-TABLE-"]))
        for v in values["-TABLE-"]:
            parse_parts(service, allEmails[v]["parts"], allEmails[v]["folder_name"], allEmails[v]["id"], True)
            window['-DOWNLOAD BAR-'].update(current_count=currentBar+1, max=len(values["-TABLE-"]))
            currentBar += 1
        allEmails, rowColors, dataTable = search_and_load(values['-SEARCH-'], "default")
        window['-TABLE-'].update(values=dataTable[1:][:],
                                 select_rows=list(selected),
                                 row_colors=rowColors)  # Update the table and the selected rows
        window["-DOWNLOAD BAR-"].update(visible=False)
        currentBar = 0
    if event == 'select all':
        for i in range(len(dataTable) - 1):
            dataTable[i + 1][0] = CHECKED_BOX
            selected.append(i)
        window['-TABLE-'].update(values=dataTable[1:][:],
                                 select_rows=list(selected),
                                 row_colors=rowColors)  # Update the table and the selected rows
        window['download'].update(disabled=False)
        window['delete'].update(disabled=False)
    if event == 'open saving directory':
        if not os.path.isdir('Downloads'):
            os.mkdir('Downloads')
        subprocess.Popen(['open', 'Downloads'])
    if event == 'clear selected':
        selected.clear()
        for i in range(len(dataTable) - 1):
            dataTable[i + 1][0] = BLANK_BOX

        window['-TABLE-'].update(values=dataTable[1:][:],
                                 select_rows=list(selected),
                                 row_colors=rowColors)  # Update the table and the selected rows
        window['download'].update(disabled=True)
        window['delete'].update(disabled=True)

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
                                 select_rows=list(selected),
                                 row_colors=rowColors)  # Update the table and the selected rows
        if len(selected) > 0:
            window['download'].update(disabled=False)
            window['delete'].update(disabled=False)
        else:
            window['download'].update(disabled=True)
            window['delete'].update(disabled=True)

window.close()
