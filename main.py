import imaplib
import email
from email.header import decode_header
import os
from datetime import datetime,timedelta
import re
import logging
import requests



emaill = "support@glovicefx.com"
pasword = "bqlrklwltyjqsfcn"
MAS_sender = "reports@mas-markets.com"
EQUITI_sender = "support.sey@equiti.com"
BROC_sender = "dealing@broprimeasia.com"
imap_server = "imap.gmail.com"


logging.basicConfig(filename='app.log', level=logging.DEBUG)
current_directory = os.getcwd()
logging.debug(f'Current working directory: {current_directory}')


#filename processing
def sanitize_filename(filename):
    # Replace any invalid characters with an underscore
    return re.sub(r'[^a-zA-Z0-9_\-.]', '_', filename)


def saver(part,type,sender):
    acc_number = ""
    if type == "text/plain" or type == "text/html":
        body = part.get_payload(decode=True).decode()
        global format_date
        date_pattern = r"(\d{4} \b(?:January|February|March|April|May|June|July|August|September|October|November|December)\b \d{1,2})"
        pattern = r"A\/C No: <b>\d+<\/b>"
        # print(body)   
        date_match = re.search(date_pattern,body)
        acc_match = re.search(pattern,body)

        if acc_match:
            acc_number+=acc_match.group().split("<b>")[1].rstrip("</b>")
            # acc_number += acc_match.groups()[0]
            # print(f"Found daily statement for {acc_number}")
        if date_match:
            date = date_match.groups()[0].split(" ")
            email_year = date[0]
            email_month = date[1]
            email_date = date[2]
            date_str = f"{email_year} {email_month} {email_date}"
            format_date = datetime.strptime(date_str,"%Y %B %d").strftime("%Y%m%d")
            # print(date_match)

        if datetime.now().weekday() != 0:
            yesterday = datetime.now() - timedelta(days=1)
        else:
            yesterday = datetime.now() - timedelta(days=3)

        formatted_ytd = datetime.strftime(yesterday,"%Y%m%d")  

        if sender == EQUITI_sender:
            folder_name_equiti = current_directory+"EQUITI"
            if not os.path.isdir(folder_name_equiti):
                os.makedirs(folder_name_equiti) #make folder if it is not existed
                #make subfolder   
            sub_folder_205 = "205"
            sub_folder_011 = "011"
            folder_205 = os.path.join(folder_name_equiti,sub_folder_205)
            folder_011 = os.path.join(folder_name_equiti,sub_folder_011)
            if not os.path.isdir(folder_205) or not os.path.isdir(folder_011):
                os.makedirs(folder_205)
                os.makedirs(folder_011)  
        
        # if sender == BROC_sender:
        #     folder_name_broc = current_directory+"BROCTAGON"
        #     if not os.path.isdir(folder_name_broc):
        #         os.makedirs(folder_name_broc)

        if format_date == formatted_ytd:
            file_name = f"Daily Statement {format_date}.html"
            file_path = ""
            if acc_number == "3540205":
                file_path += os.path.join(folder_205,file_name)
                with open(file_path,"w",encoding="utf-8") as file:
                    file.write(body)
            elif acc_number == "3540011":
                file_path += os.path.join(folder_011,file_name)
                with open(file_path, "w", encoding="utf-8") as file:
                    file.write(body)
            
            print(f"Save complete for {acc_number} - {sender} in {file_path}")
            logging.debug(f"Save completed for {sender}")

        
def broctagon_saver(msg):
    if datetime.now().weekday() != 0:
        yesterday = datetime.now() - timedelta(days=1)
    else:
        yesterday = datetime.now() - timedelta(days=3)

    formatted_ytd = datetime.strftime(yesterday,"%d-%b-%Y") 
    charset = msg.get_content_charset()
    body = msg.get_payload(decode=True).decode(charset if charset else 'utf-8')
    date_pattern = r"\b\d{4}\.\d{2}\.\d{2}\b"
    acc_pattern = r"A\/C No: <b>\d+<\/b>"
    date = re.search(date_pattern,body)
    acc = re.search(acc_pattern,body)
    acc_num=""
    if acc:
        acc_num += acc.group().split("<b>")[1].rstrip("</b>")
        if date:
            date_str=date.group(0)
            date_format = datetime.strptime(date_str,"%Y.%m.%d").strftime("%d-%b-%Y")
                                    
            folder_name_broc = os.getcwd()+ "BROCTAGON"
            if not os.path.isdir(folder_name_broc):
                os.makedirs(folder_name_broc)
            sub_folder_821 = "101821"
            sub_folder_822 = "101822"

            folder_821 = os.path.join(folder_name_broc,sub_folder_821)
            folder_822 = os.path.join(folder_name_broc,sub_folder_822)
            if not os.path.isdir(folder_821) or not os.path.isdir(folder_822):
                os.makedirs(folder_821)
                os.makedirs(folder_822)  


            if date_format == formatted_ytd:
                file_name = f"Daily Statement {date_format}.html"
                filepath = ""
                if acc_num == "101821":
                    filepath = os.path.join(folder_821, file_name)
                    print(f"Filepath set to: {filepath}")  
                
                if acc_num == "101822":
                    filepath = os.path.join(folder_822, file_name)
                    print(f"Filepath set to: {filepath}")  
                
                if filepath: 
                    with open(filepath, "w", encoding="utf-8") as file:
                        file.write(body)
                        print(f"Save complete for 101821 - Broctagon: {filepath}")
                        logging.debug("Saved Broctagon")
                else:
                    print("Filepath not set. Check account number or other conditions.")


def mas_saver(content_disposition,part):
    #need to add code in here to convert pdf one again ensure it will be open in the micorosoft edge
    if content_disposition:
        dispositions = content_disposition.strip().split(";")
        if any(disposition.strip().lower() == "attachment" for disposition in dispositions):
            filename = part.get_filename()
            folder_name = current_directory + "MAS"
            if not os.path.isdir(folder_name):
                os.mkdir(folder_name)
            if filename and filename.lower().endswith('.pdf'):
                global formatted_filename
                formatted_filename = sanitize_filename(filename)  

                filepath = os.path.join(folder_name,formatted_filename)
                with open(filepath, 'wb') as f:
                    f.write(part.get_payload(decode=True))
                    print("Save compelete for MAS")
                    logging.debug(f"Save completed for MAS")

def save_VC(content_disposition,part):
    
    if content_disposition:
        dispositions = content_disposition.strip().split(";")
        if any(disposition.strip().lower() in ["attachment", "inline"] for disposition in dispositions):
            filename = part.get_filename()

            # Ensure filename exists and ends with .htm
            if filename and filename.lower().endswith('.htm'):
                # Create the folder where files will be saved
                folder_name = current_directory + "VC"
                if not os.path.isdir(folder_name):
                    os.mkdir(folder_name)

                ytd_date = (datetime.now() - timedelta(days=1)).strftime('%Y%m%d')
                target_folder = os.path.join(folder_name, ytd_date)
                os.makedirs(target_folder, exist_ok=True)

                # Create the full file path
                filepath = os.path.join(target_folder, filename)

                # Save the .htm file
                with open(filepath, 'wb') as f:
                    f.write(part.get_payload(decode=True))
                    print(f"Save complete for VC: {filepath}")

def log_in_server(mail,password,server,sender):
    with imaplib.IMAP4_SSL(server) as connection:
        connection.login(mail,password)
        status, result = connection.select("Inbox")
        if datetime.now().weekday() != 0:
            yesterday = datetime.now() - timedelta(days=1)
        else:
            yesterday = datetime.now() - timedelta(days=3)
        # print(yesterday)

        if sender == MAS_sender:
            format_yesterday = datetime.strftime(yesterday,"%Y-%m-%d")
            status, messages = connection.search(None, f'(FROM "{sender}" SUBJECT "Daily Statement KJ Glovice FX - {format_yesterday}")')
        elif sender == EQUITI_sender:
            format_yesterday = datetime.strftime(yesterday,"%d-%b-%Y")
            status, messages = connection.search(None, f'(FROM "{sender}" SUBJECT "Daily Confirmation" SINCE {format_yesterday})')
        elif sender == BROC_sender:
            format_yesterday = datetime.strftime(yesterday,"%d-%b-%Y")
            status, messages = connection.search(None, f'(FROM "{sender}" SUBJECT "Daily Confirmation" SINCE {format_yesterday})')
        elif sender == emaill:
            format_yesterday = datetime.strftime(yesterday,"%Y-%m-%d")
            status, messages = connection.search(None, f'(FROM "{sender}" SUBJECT "Daily Confirmation - {format_yesterday}")')

            
        if status!= "OK":
            print("Can not log in")
        else:
            print(f"Log in {mail}")
            print(f"Searching data from {sender}")
            email_ids = messages[0].split()
            # print(email_ids)
        for id in email_ids:
            status, msg_data = connection.fetch(id, '(RFC822)')
            if status!= "OK": 
                print ("Can not load email")
            else:
                for response in msg_data:
                    if isinstance(response,tuple): #resopnse is in tuple type
                        msg = email.message_from_bytes(response[1])
                        # print(msg)
                    # Decode the email subject
                        subject, encoding = decode_header(msg["Subject"])[0]
                        if isinstance(subject, bytes):
                            subject = subject.decode(encoding if encoding else 'utf-8')
                        # print(msg.is_multipart())

                        if msg.is_multipart():
                            for part in msg.walk():
                                content_disposition = part.get("Content-Disposition", None)
                                content_type = part.get_content_type()
                                # print(content_type)
                                if sender == EQUITI_sender:
                                    saver(part,content_type,sender)
                                elif sender == MAS_sender:
                                    mas_saver(content_disposition,part)
                                elif sender == emaill:
                                    save_VC(content_disposition,part)
                    
                        else:
                            content_disposition = msg.get("Content-Disposition", None)
                            content_type = msg.get_content_type()
                            # print(content_type)
                            if sender == EQUITI_sender:
                               saver(msg,content_type,sender)
                            elif sender == MAS_sender:
                                mas_saver(content_disposition,msg)
                            elif sender == BROC_sender:
                                broctagon_saver(msg)
                            elif sender == emaill:
                               save_VC(content_disposition,msg)

                       # log_in_server(emaill,pasword,imap_server,MAS_sender)



def save_swap_bix():
    print("Saving swap for Bidx")
    session = requests.Session()

    # Set custom headers to simulate a browser request (optional)
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36",
        "Accept": "application/json, text/javascript, */*; q=0.01"
    }
    url = "https://bidxmarkets.sharepoint.com/:x:/s/Operations/Ec68NemlfrlIrkhWwup5KecBlb6YcLjTyxfqaWTltCivOg?download=1"
    
    # Create the "SWAP MAS" folder if it doesn't exist
    folder_name = current_directory +  " SWAP MAS"
    if not os.path.isdir(folder_name):
        os.makedirs(folder_name)

    # Generate the save path using the current date
    today = datetime.now().strftime("%Y%m%d")
    save_path = f"{today} swap.xlsx"
    filepath = os.path.join(folder_name, save_path)

    try:
        # Make the request to download the file
        response = session.get(url, headers=headers, stream=True)

        # Check if the request is successful
        response.raise_for_status()

        # Save the file to disk
        with open(filepath, "wb") as file:
            for chunk in response.iter_content(chunk_size=8192):
                file.write(chunk)

        print(f"File downloaded successfully and saved to {filepath}")

    except requests.exceptions.RequestException as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    log_in_server(emaill,pasword,imap_server,EQUITI_sender)
    log_in_server(emaill,pasword,imap_server,MAS_sender)
    log_in_server(emaill,pasword,imap_server,BROC_sender)
    log_in_server(emaill,pasword, imap_server,emaill)
    save_swap_bix()
    input("Press Enter to exit...")