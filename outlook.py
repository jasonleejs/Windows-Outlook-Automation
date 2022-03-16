import os
import datetime
import win32com.client

def save_attachments(save_path, subject=None, sender=None, date=None, save_types=None, avoid_types=None):
    download_count = 0
    for message in messages:
        attachments = message.Attachments
        num_attach = len([x for x in attachments])
        for x in range(1, num_attach):
            attachment = attachments.Item(x)
            try: 
                attachment.SaveASFile(os.path.join(save_path, attachment.FileName))
                print (f"Downloaded {attachment}")
                download_count += 1
            except:
                print(f"Failed to download {attachment} from {message.Subject}")
        message = messages.GetNext()
    print(f"{download_count} Total attachments downloaded")


if __name__ == '__main__':
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case the inbox. You can change that number to reference
    messages = inbox.Items

    save_path = r"C:\Users\jlee23\jlee23\Outlook Automations\attachments_test"
    save_path = input("Path for saving directory: ")

    save_attachments(save_path)
