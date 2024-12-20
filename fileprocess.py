import zipfile
import os 
from datetime import datetime,timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import smtplib
from pathlib import Path
import logging

#credential 
EMAIL = "support@glovicefx.com"
imap_server = "imap.gmail.com"
password = "bqlrklwltyjqsfcn"

#create message container
logging.basicConfig(filename="log.log", level=logging.INFO)

target_files = [
    "22040001_mail.htm",
    "23010001_mail.htm",
    "23020001_mail.htm",
    "23030001_mail.htm",
]
folder = os.getcwd()

logging.info(folder)
#unzipfile function
def unzip(zip_file, target_files):
    folder_name = os.path.splitext(zip_file)[0].split("/")[0]  # Extract the folder name correctly
    folder_save = os.path.join(folder, folder_name)
    os.makedirs(folder_save, exist_ok=True)  # Ensure the target folder exists

    with zipfile.ZipFile(zip_file, "r") as zip_ref:
        for target_file in target_files:
            if target_file in zip_ref.namelist():
                zip_ref.extract(target_file, folder_save)
                logging.info(f"Unzip file successfully {target_file}")
                # send_email(imap_server, EMAIL, password, target_file)
    return folder_save

#send email function
def send_email(server,email,password,file,name):
    msg = MIMEMultipart()
    msg['From'] = email
    msg['To'] = email
    msg['Subject'] = f'Daily Confirmation - {datetime.now().strftime("%Y-%m-%d")}'
    body = "Please find HTML attachment below."
    msg.attach(MIMEText(body, 'plain'))
    # file_path = f"../{file}"
    # for file in files:
    try:
        with open(file, 'rb') as file:  # Open the file in binary mode
            attachment = MIMEBase('application', 'octet-stream')
            attachment.set_payload(file.read())

        # Encode the file in base64
        encoders.encode_base64(attachment)

        # Add header for file attachment
        # attachment.add_header('Content-Disposition', f'attachment; filename={name}')
        attachment.add_header('Content-Disposition', f'attachment; filename={name}')


        # Attach the file to the message
        msg.attach(attachment)

    except FileNotFoundError:
        print(f"Error: The file {file} was not found.")
        logging.info(f"Error: The file {file} was not found.")

    # Setup the SMTP server
    with smtplib.SMTP(server, 587) as conn:
        conn.starttls()
        conn.login(email, password)
        # Send the email
        conn.sendmail(msg['From'], msg['To'], msg.as_string())
        print(f"Sent sucessfully {file}")
        logging.info(f"Sent sucessfully {file}")

# Find the folder with the date before
#ytd_date = (datetime.now() - timedelta(days=0)).strftime('%Y%m%d')
ytd_date = datetime.now().strftime('%Y%m%d')
print(ytd_date)
logging.info(f"Sending email for {ytd_date}")
target_folder = f"{ytd_date}.daily"
target_folder_path = os.path.join(folder, target_folder)
logging.info(f"Zip file: {target_folder} - {target_folder_path}")

if os.path.exists(target_folder_path):
    for item in os.listdir(target_folder_path):
        item_path = os.path.join(target_folder_path, item)
        logging.info(f"File path: {item_path}")
        if item.endswith(".zip"):
            folder_save = unzip(item_path, target_files)
            print(folder_save)
    for file in os.listdir(folder_save):
        file_path = os.path.join(folder_save, file)  # Construct absolute file path
        send_email(imap_server, EMAIL, password, file_path,file) 

# if os.path.exists(target_folder_path):
#     all_file_paths = []  # Collect all file paths to send in one email
#     for item in os.listdir(target_folder_path):
#         item_path = os.path.join(target_folder_path, item)
#         if item.endswith(".zip"):
#             folder_save = unzip(item_path, target_files)
#             print(folder_save)
#             for file in os.listdir(folder_save):
#                 file_path = os.path.join(folder_save, file)  # Construct absolute file path
#                 print(file_path)
#                 all_file_paths.append(file_path)  # Add file path to the list

    # Send all attachments in one email
    # send_email(imap_server, EMAIL, password, all_file_paths)  # P
                    