# PyEmail Manager

This project uses Google's Gmail API to fetch, load, and display emails. Emails can then be downloaded to your physical computer and/or deleted from your Google Drive (while still being saved on your computer, if applicable). 

## Features

* fluid loading bars on startup and during operations (deleting, downloading)
![Screenshot 2024-04-06 at 11 58 05 AM](https://github.com/dwarakeshbaraneetharan/EmailManager/assets/55335467/0128b637-4ec8-498a-9e42-3554922fe4cb)
* Fully-featured, modern GUI
  
![Screenshot 2024-04-06 at 11 58 19 AM](https://github.com/dwarakeshbaraneetharan/EmailManager/assets/55335467/eaa0fd7b-6d0d-4b0c-8b50-8aac4264d932)

* storage space display
* searchbox that uses Google's native Gmail search for optimal results
* antimated buttons with colors that reflect state and function
* select all, clear selected, and range functions to make choosing emails easier
* shortcut to open directory where emails are saved

![Screenshot 2024-04-06 at 12 24 13 PM](https://github.com/dwarakeshbaraneetharan/EmailManager/assets/55335467/071cb3b8-3715-4e51-a8a5-7a59d76455d1)


## Getting Started

### Installing

* Signed into the Google Account you wish to use, create a new Google Cloud Project at: https://console.cloud.google.com/welcome
* Under "APIs & Services", choose "Enabled APIs & Services" and enable the Gmail API and Google Drive API
* Navigate to "Credentials", and create a new OAuth 2.0 Client ID with type "Desktop App"
* Download the OAuth client .json file for the newly created credential
![Screenshot 2024-04-06 at 12 21 29 PM](https://github.com/dwarakeshbaraneetharan/EmailManager/assets/55335467/9d211a41-944f-43c3-a71a-c3f17bdf82af)

* Clone this repo or download as .zip file
* Rename the previously downloaded credentials file to "credentials.json" and place it in the same folder as main.py

### Executing program

* run main.py
* in the window that opens, sign into your Google account
* click "Continue" after signing in
* give your Google Cloud project access to access your Gmail and Google Drive
* PyEmail Manager should open automatically
