# pip install pywin32
import win32com.client
import os
import time
import re
import string

def find_pst_folder(OutlookObj, pst_filepath) :
    for Store in OutlookObj.Stores :
        if Store.IsDataFileStore and Store.FilePath == pst_filepath :
            return Store.GetRootFolder()
    return None

remove_list = [  "Classificado como Uso Interno."
               , "Classificado como Público."
                "Importance: High."]

def sanitize_folder_name(folder_name):
    valid_chars = "\-_.() %s%s" % (string.ascii_letters, string.digits)
    cleaned_folder_name = ''.join(c for c in folder_name if c in valid_chars)
    return cleaned_folder_name.replace(' ','_')
    
def sanitize_text(text, remove_list=None):
    if text is None:
        return ""
    if isinstance(text, bytes):
        text = text.decode('utf-8', errors='ignore')  
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = text.replace("\t", " ")

    if remove_list is not None:
        for phrase in remove_list:
            text = text.replace(phrase, '')

    text = re.sub(' +', ' ', text)
    text = re.sub('\n+', '\n', text)

    lines = text.split('\n')
    lines = [line if line.endswith('.') else f'{line}.' for line in lines]
    
    lines = text.split('\n')
    lines = [line.strip() if line.strip().endswith('.') else f'{line.strip()}.' for line in lines]
    
    lines = [line for line in lines if not re.fullmatch('\s*\.\s*', line)]
    text = '\n'.join(lines)

    text = re.sub('\.{2,}', '.', text)

    return text



def enumerate_folders(FolderObj, path) :
    for ChildFolder in FolderObj.Folders :
        new_path = os.path.join(path, sanitize_folder_name(ChildFolder.Name))
        if not os.path.exists(new_path):
            os.makedirs(new_path)
        enumerate_folders(ChildFolder, new_path)
    iterate_messages(FolderObj, path)

def iterate_messages(FolderObj, path):
    message_number = 0
    for item in FolderObj.Items:
        message_number += 1
        try:
            file = os.path.join(path, f'{message_number}.txt')
            if os.path.isfile(file):
                print("Skip .. " + file)
            else:
                header = ""
                header += "Sender Name: " + str(item.SenderName if hasattr(item, 'SenderName') else "") + "\n"
                header += "Sender Email: " + str(item.SenderEmailAddress if hasattr(item, 'SenderEmailAddress') else "") + "\n"
                header += "To: " + (str(item.To) if hasattr(item, 'To') else "") + "\n" + (str(item.CC) if hasattr(item, 'CC') else "")
                header += "Subject: " + (str(item.Subject) if hasattr(item, 'Subject') else "") + "\n"
                # loop através dos emails na caixa de entrada
                body_text = sanitize_text(item.Body)
                content = header + "\n" + body_text
                with open(file, 'w', encoding='utf-8') as f:
                    f.write(content)
                    f.close()
                print(file)
                try:
                    count_attachments = item.Attachments.Count
                    if count_attachments > 0 :
                        att_path = os.path.join(path, str(message_number))  # create attachment path
                        if not os.path.exists(att_path):
                            os.makedirs(att_path)  # create the path if not exists            
                        for att in range(count_attachments) :
                            att_file_path = os.path.join(att_path, sanitize_folder_name(item.Attachments.Item(att + 1).FileName))  # create attachment file path
                            item.Attachments.Item(att + 1).SaveAsFile(att_file_path)  # save the attachment
                except:
                   print("error in attachment name")
        except:
            print("error in attachment name")
    
    
Outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

pst = r"c:\pathto\yourpst\folder"
BASE_DIRECTORY = r"c:\whereexport\files"  # Chose path
Outlook.AddStore(pst)
PSTFolderObj = find_pst_folder(Outlook,pst)
try :
    enumerate_folders(PSTFolderObj, BASE_DIRECTORY)
except Exception as exc :
    print(exc)
finally :
    Outlook.RemoveStore(PSTFolderObj)
