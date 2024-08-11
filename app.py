import os
import re
import shutil
import ftplib
import qrcode
import subprocess
import threading
from docx import Document
from docxtpl import DocxTemplate
import pythoncom
import win32com.client
from win32com.shell import shell, shellcon
import tkinter as tk
from tkinter import ttk

## Config File Path
data_dir = os.environ["USERPROFILE"] + "/AppData/Local/Certificate Utility" # ~/AppData/Local/Certificate Utility
## Export File Path
base_dir = shell.SHGetFolderPath(0, shellcon.CSIDL_PERSONAL, None, 0) + "/" + "Certificates" # ~/Documents/Certificate Exports

try:
    os.mkdir(data_dir)
except FileExistsError:
    pass
try:
    os.mkdir(base_dir)
except FileExistsError:
    pass

def config():
    try:
        global ftp_host, ftp_port, ftp_path, ftp_username, ftp_password, url
        with open(f"{data_dir}/config.txt", "r") as file:
            data = file.read()
            ftp_host = data.split("host=")[1].split("\n")[0] if data.count("host=") > 0 else ""
            ftp_port = data.split("port=")[1].split("\n")[0] if data.count("port=") > 0 else ""
            ftp_path = data.split("path=")[1].split("\n")[0] if data.count("path=") > 0 else ""
            ftp_username = data.split("username=")[1].split("\n")[0] if data.count("username=") > 0 else ""
            ftp_password = data.split("password=")[1].split("\n")[0] if data.count("password=") > 0 else ""
            url = data.split("url=")[1].split("\n")[0] if data.count("url=") > 0 else ""
    except FileNotFoundError:
        open(f"{data_dir}/config.txt", "w")
        config()

config()

def replace_placeholders(doc, replacements):
    for paragraph in doc.paragraphs:
        for placeholder, replacement in replacements.items():
            if placeholder in paragraph.text:
                for run in paragraph.runs:
                    run.text = run.text.replace(placeholder, replacement)

def generate_qr(url, save_qr):
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=10,
        border=1,
    )
    qr.add_data(url)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    img.save(save_qr)

def qr_code(file, qr_code):
    tpl = DocxTemplate(file)
    tpl.replace_media("qrcode.png", qr_code)
    tpl.save(file)

def convert_pdf(in_file, out_file):
    wdFormatPDF = 17
    word = win32com.client.Dispatch('Word.Application', pythoncom.CoInitialize())
    doc = word.Documents.Open(in_file)
    doc.SaveAs(out_file, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()

def create_certificate():
    doc = Document("template.docx")
    name = entry1.get()
    date = entry2.get()
    place = entry3.get()
    certificate_no = entry4.get()
    idd = entry5.get()

    replacements = {
        '[NAME]': name,
        '[PLACE]': place,
        '[DATE]': date,
        '[CERT-NO]': certificate_no,
        '[ID]': idd,
    }
    if not re.match("^https?:\\/\\/(?:www\\.)?[-a-zA-Z0-9@:%._\\+~#=]{1,256}\\.[a-zA-Z0-9()]{1,6}\\b(?:[-a-zA-Z0-9()@:%_\\+.~#?&\\/=]*)/$", url):
        return 3
    certificate_name = certificate_no.replace("/", "-").replace("\\", "-")
    global certificate_dir
    certificate_dir = (base_dir + "/" + certificate_name + "/").replace("/", "\\")
    try:
        os.mkdir(certificate_dir)
    except FileExistsError:
        return 2
    try:
        replace_placeholders(doc, replacements)
        doc.save(certificate_dir + "certificate.docx")
        name_format = name.replace(" ", "_").replace("/", "").replace("\\", "").replace("'", "").replace('"', '')
        save_qr_code = certificate_dir + "qr-code.png"
        generate_qr(f"{url}{certificate_name}-{name_format}", save_qr_code)
        qr_code(certificate_dir + "certificate.docx", save_qr_code)
        with open(certificate_dir + "url.txt", "w") as path_file:
            path_file.write(f"{url}{certificate_name}-{name_format}")
        try:
            os.remove(save_qr_code)
        except FileNotFoundError or PermissionError:
            pass
        convert_pdf(certificate_dir + "certificate.docx", certificate_dir + "certificate.pdf")
        return 0
    except:
        return 1

def generate_certificate():
    context_text.unbind("<Button-1>")
    context_text['cursor'] = ""
    generate_button["state"] = "disabled"
    context_text.update_idletasks()
    generate_button.update_idletasks()
    if len(entry1.get()) > 0 and len(entry2.get()) > 0 and len(entry3.get()) > 0 and len(entry4.get()) > 0 and len(entry5.get()) > 0:
        context_text['text'] = "Generating Certificate..."
        context_text["foreground"] = "green"
        context_text.update_idletasks()
        res = create_certificate()
        if res == 1:
            context_text['text'] = "Failed to generate certificate"
            context_text["foreground"] = "red"
        elif res == 2:
            context_text['text'] = "Certificate already exists"
            context_text["foreground"] = "red"
            context_text.bind("<Button-1>", lambda event: open_explorer())
            context_text['cursor'] = "hand2"
        elif res == 3:
            context_text['text'] = "Invalid value for Upload URL"
            context_text["foreground"] = "red"
        else:
            context_text['text'] = "Click to open certificate"
            context_text["foreground"] = "green"
            context_text.bind("<Button-1>", lambda event: open_explorer())
            context_text['cursor'] = "hand2"
    else:
        context_text['text'] = "Empty values detected"
        context_text["foreground"] = "red"
    reset_button["state"] = "enabled"
    context_text.update_idletasks()
    reset_button.update_idletasks()

def upload_certificate():
    if upload_button["text"] == "Upload":
        upload_button["state"] = "disabled"
        upload_button.update_idletasks()
        if len(upload_entry.get()) > 0:
            cert_dir = base_dir + "/" + upload_entry.get().replace("/", "-").replace("\\", "-")
            try:
                open(cert_dir + "/certificate.pdf", "r").close()
                no_error = True
            except FileNotFoundError:
                upload_context_text['text'] = "Certificate not found on this computer"
                upload_context_text["foreground"] = "red"
                upload_context_text.update_idletasks()
                no_error = False
            if no_error:
                if ftp_host == "" or ftp_port == "" or ftp_username == "" or ftp_path == "":
                    upload_context_text['text'] = "  FTP Settings not configured correctly"
                    upload_context_text["foreground"] = "red"
                    upload_context_text.update_idletasks()
                    no_error = False
            if no_error:
                ip_pattern = r'\b(?:\d{1,3}\.){3}\d{1,3}\b'
                domain_pattern = r'\b(?:[a-zA-Z0-9](?:[-a-zA-Z0-9]*[a-zA-Z0-9])?\.)+[a-zA-Z]{2,}\b'
                ip_match = re.match(ip_pattern, ftp_host)
                domain_match = re.match(domain_pattern, ftp_host)
                if not (ip_match or domain_match):
                    upload_context_text['text'] = "FTP Host must be a valid ip or domain"
                    upload_context_text["foreground"] = "red"
                    upload_context_text.update_idletasks()
                    no_error = False
            if no_error:
                try:
                    int(ftp_port)
                    no_error = True
                except:
                    upload_context_text['text'] = "  FTP Port must only contain number"
                    upload_context_text["foreground"] = "red"
                    upload_context_text.update_idletasks()
                    no_error = False
            if no_error:
                try:
                    upload_context_text['text'] = "    Connecting to server... Please wait"
                    upload_context_text["foreground"] = "green"
                    upload_context_text.update_idletasks()
                    try:
                        ftp = ftplib.FTP_TLS()
                        ftp.connect(ftp_host, int(ftp_port))
                        ftp.auth()
                        ftp.prot_p()
                        ftp.login(ftp_username, ftp_password)
                    except:
                        ftp = ftplib.FTP()
                        ftp.connect(ftp_host, int(ftp_port))
                        ftp.login(ftp_username, ftp_password)
                    no_error = True
                except:
                    upload_context_text['text'] = "Unable to connect to remote ftp server"
                    upload_context_text["foreground"] = "red"
                    upload_context_text.update_idletasks()
                    no_error = False
            if no_error:
                try:
                    ftp.cwd(ftp_path)
                    no_error = True
                except:
                    upload_context_text['text'] = "Invalid path specified in ftp settings"
                    upload_context_text["foreground"] = "red"
                    upload_context_text.update_idletasks()
                    no_error = False
            if no_error:
                upload_context_text['text'] = "Uploading Ceritificate, Please wait..."
                upload_context_text["foreground"] = "green"
                upload_context_text.update_idletasks()
                try:
                    with open(cert_dir + "/url.txt") as cert_path:
                        cert_remote_dir = cert_path.read().split("/")[-1]
                        ftp.mkd(cert_remote_dir)
                    no_error = True
                except:
                    upload_context_text['text'] = "Error occurred while uploading, try again"
                    upload_context_text["foreground"] = "red"
                    no_error = False
                if no_error:
                    src_html_file = cert_dir + "/index.html"
                    dst_html_file = cert_remote_dir + "/index.html"
                    src_pdf_file = cert_dir + "/certificate.pdf"
                    dst_pdf_file = cert_remote_dir + "/certificate.pdf"
                    try:
                        ftp.storbinary(f"STOR {dst_html_file}", open(src_html_file,'rb'))
                        ftp.storbinary(f"STOR {dst_pdf_file}", open(src_pdf_file,'rb'))
                        upload_context_text['text'] = "    Certificate uploaded successfully"
                        upload_context_text["foreground"] = "green"
                    except:
                        upload_context_text['text'] = "Can't upload certificate on remote server"
                        upload_context_text["foreground"] = "red"
                upload_context_text.update_idletasks()
            try:
                ftp.quit()
            except:
                pass
        else:
            upload_context_text['text'] = "Please enter a valid certificate number"
            upload_context_text["foreground"] = "red"
            upload_context_text.update_idletasks()
        upload_button['text'] = "Reset"
        upload_button["state"] = "enabled"
        upload_button.update_idletasks()
    else:
        upload_entry.delete(0, "end")
        upload_entry.update_idletasks()
        upload_context_text['text'] = ""
        upload_context_text.update_idletasks()
        upload_button['text'] = "Upload"
        upload_button['state'] = "enabled"
        upload_button.update_idletasks()

def delete_certificate():
    if delete_button["text"] == "Delete":
        delete_button["state"] = "disabled"
        delete_button.update_idletasks()
        if len(delete_entry.get()) > 0:
            try:
                cert_dir = base_dir + "/" + delete_entry.get().replace("/", "-").replace("\\", "-")
                no_error = True
            except:
                delete_context_text['text'] = "Certificate not found on this computer"
                delete_context_text["foreground"] = "red"
                delete_context_text.update_idletasks()
                no_error = False
            if no_error:
                if ftp_host == "" or ftp_port == "" or ftp_username == "" or ftp_path == "":
                    delete_context_text['text'] = " FTP Settings not configured correctly"
                    delete_context_text["foreground"] = "red"
                    delete_context_text.update_idletasks()
                    no_error = False
            if no_error:
                ip_pattern = r'\b(?:\d{1,3}\.){3}\d{1,3}\b'
                domain_pattern = r'\b(?:[a-zA-Z0-9](?:[-a-zA-Z0-9]*[a-zA-Z0-9])?\.)+[a-zA-Z]{2,}\b'
                ip_match = re.match(ip_pattern, ftp_host)
                domain_match = re.match(domain_pattern, ftp_host)
                if not (ip_match or domain_match):
                    delete_context_text['text'] = "FTP Host must be a valid ip or domain"
                    delete_context_text["foreground"] = "red"
                    delete_context_text.update_idletasks()
                    no_error = False
            if no_error:
                try:
                    int(ftp_port)
                    no_error = True
                except:
                    delete_context_text['text'] = "  FTP Port must only contain number"
                    delete_context_text["foreground"] = "red"
                    delete_context_text.update_idletasks()
                    no_error = False
            if no_error:
                try:
                    delete_context_text['text'] = "    Connecting to server... Please wait"
                    delete_context_text["foreground"] = "green"
                    delete_context_text.update_idletasks()
                    try:
                        ftp = ftplib.FTP_TLS()
                        ftp.connect(ftp_host, int(ftp_port))
                        ftp.auth()
                        ftp.prot_p()
                        ftp.login(ftp_username, ftp_password)
                    except:
                        ftp = ftplib.FTP()
                        ftp.connect(ftp_host, int(ftp_port))
                        ftp.login(ftp_username, ftp_password)
                    no_error = True
                except:
                    delete_context_text['text'] = "Unable to connect to remote ftp server"
                    delete_context_text["foreground"] = "red"
                    delete_context_text.update_idletasks()
                    no_error = False
            if no_error:
                try:
                    ftp.cwd(ftp_path)
                    no_error = True
                except:
                    delete_context_text['text'] = "Invalid path specified in ftp settings"
                    delete_context_text["foreground"] = "red"
                    delete_context_text.update_idletasks()
                    no_error = False
            if no_error:
                try:
                    delete_context_text['text'] = "Deleting Ceritificate, Please wait..."
                    delete_context_text["foreground"] = "green"
                    delete_context_text.update_idletasks()
                    for i in ftp.nlst():
                        if i.startswith(delete_entry.get().replace("/", "-").replace("\\", "-")):
                            ftp.cwd(i)
                            for x in ftp.nlst():
                                try:
                                    try:
                                        ftp.delete(x)
                                    except:
                                        ftp.rmd(x)
                                except:
                                    pass
                            ftp.cwd(ftp_path)
                            ftp.rmd(i)
                            delete_context_text['text'] = "     Certificate deleted successfully"
                            delete_context_text["foreground"] = "green"
                            no_error = True
                            break
                    else:
                        delete_context_text['text'] = "Certificate not found on remote server"
                        delete_context_text["foreground"] = "red"
                        no_error = False
                    if no_error:
                        try:
                            shutil.rmtree(cert_dir)
                        except:
                            delete_context_text['text'] = "Failed to remove certificate on this PC"
                            delete_context_text["foreground"] = "red"
                except:
                    delete_context_text['text'] = "Error occurred while deleting, try again"
                    delete_context_text["foreground"] = "red"
                delete_context_text.update_idletasks()
            try:
                ftp.quit()
            except:
                pass
        else:
            delete_context_text['text'] = "Please enter a valid certificate number"
            delete_context_text["foreground"] = "red"
            delete_context_text.update_idletasks()
        delete_button['text'] = "Reset"
        delete_button["state"] = "enabled"
        delete_button.update_idletasks()
    else:
        delete_entry.delete(0, "end")
        delete_entry.update_idletasks()
        delete_context_text['text'] = ""
        delete_context_text.update_idletasks()
        delete_button['text'] = "Delete"
        delete_button['state'] = "enabled"
        delete_button.update_idletasks()

def open_explorer():
    CREATE_NO_WINDOW = 0x08000000
    subprocess.call(f"explorer {certificate_dir}", creationflags=CREATE_NO_WINDOW)

def reset_fields():
    context_text['text'] = ""
    reset_button["state"] = "disabled"
    entry1.delete(0, "end")
    entry2.delete(0, "end")
    entry3.delete(0, "end")
    entry4.delete(0, "end")
    entry5.delete(0, "end")
    generate_button["state"] = "enabled"
    context_text.update_idletasks()
    generate_button.update_idletasks()
    reset_button.update_idletasks()

def generate_certificate_thread():
    threading.Thread(target=generate_certificate).start()

def upload_certificate_thread():
    threading.Thread(target=upload_certificate).start()

def delete_certificate_thread():
    threading.Thread(target=delete_certificate).start()

window = tk.Tk()
window.title("Certificate Utility")
window.geometry("500x500")
window.resizable(False, False)
window.wm_iconphoto(False, tk.PhotoImage(file="icon.png"))

heading_label1 = ttk.Label(window, text="Create Certificate", font=("Helvetica", 16))
heading_label1.pack()
separator1 = ttk.Separator(window, orient="horizontal")
separator1.place(x=0, y=30, width=500)
sub_text = ttk.Label(window, text="Enter Details for Certificate", anchor='w', font=("Helvetica", 10)).pack(pady=30)

label1 = ttk.Label(window, text="Name :", anchor='w', font=("Helvetica", 10)).place(x=35, y=100)
label2 = ttk.Label(window, text="Date :", anchor='w', font=("Helvetica", 10)).place(x=35, y=130)
label3 = ttk.Label(window, text="Place :", anchor='w', font=("Helvetica", 10)).place(x=35, y=160)
label4 = ttk.Label(window, text="Certificate No :", anchor='w', font=("Helvetica", 10)).place(x=35, y=190)
label5 = ttk.Label(window, text="ID :", anchor='w', font=("Helvetica", 10)).place(x=35, y=220)
entry1 = ttk.Entry(window)
entry1.place(x=150, y=98, width=300, height=25)
entry2 = ttk.Entry(window)
entry2.place(x=150, y=128, width=300, height=25)
entry3 = ttk.Entry(window)
entry3.place(x=150, y=158, width=300, height=25)
entry4 = ttk.Entry(window)
entry4.place(x=150, y=188, width=300, height=25)
entry5 = ttk.Entry(window)
entry5.place(x=150, y=218, width=300, height=25)

reset_button = ttk.Button(window, text="Reset", command=reset_fields)
reset_button["state"] = "disabled"
reset_button.place(x=150, y=260)
generate_button = ttk.Button(window, text="Generate", command=generate_certificate_thread)
generate_button.place(x=240, y=260)
context_text = ttk.Label(window, text="")
context_text.place(x=330, y=262)

heading_label2 = ttk.Label(window, text="Upload Certificate", font=("Helvetica", 13))
heading_label2.place(x=40, y=310)
upload_label = ttk.Label(window, text="Certificate No:", anchor='w', font=("Helvetica", 9)).place(x=20, y=355)
upload_entry = ttk.Entry(window)
upload_entry.place(x=100, y=353, width=130, height=25)
upload_warning_text = ttk.Label(window, text="Uploads certificate to remote ftp server", font=("Helvetica", 8), foreground="grey")
upload_warning_text.place(x=25, y=385)
upload_button = ttk.Button(window, text="Upload", command=upload_certificate_thread)
upload_button.place(x=85, y=410)
upload_context_text = ttk.Label(window, text="")
upload_context_text.place(x=20, y=440)
heading_label3 = ttk.Label(window, text="Delete Certificate", font=("Helvetica", 13))
heading_label3.place(x=310, y=310)
delete_label = ttk.Label(window, text="Certificate No:", anchor='w', font=("Helvetica", 9)).place(x=270, y=355)
delete_entry = ttk.Entry(window)
delete_entry.place(x=350, y=353, width=130, height=25)
delete_warning_text = ttk.Label(window, text="Deletes certificate from this pc and server", font=("Helvetica", 8), foreground="grey")
delete_warning_text.place(x=270, y=385)
delete_button = ttk.Button(window, text="Delete", command=delete_certificate_thread)
delete_button.place(x=335, y=410)
delete_context_text = ttk.Label(window, text="")
delete_context_text.place(x=270, y=440)

separator2 = ttk.Separator(window, orient="horizontal")
separator2.place(x=0, y=300, width=500)
separator3 = ttk.Separator(window, orient="vertical")
separator3.place(x=250, y=300, height=180)
separator2 = ttk.Separator(window, orient="horizontal")
separator2.place(x=0, y=480, width=500)

def update_settings():
    global ftp_host, ftp_port, ftp_path, ftp_username, ftp_password, url
    ftp_host = host_entry.get()
    ftp_port = port_entry.get()
    ftp_path = path_entry.get()
    ftp_username = username_entry.get()
    ftp_password = password_entry.get()
    url = url_entry.get()
    with open(f"{data_dir}/config.txt", "w") as file:
        file.writelines([f"host={ftp_host}\n", f"port={ftp_port}\n", f"path={ftp_path}\n", f"username={ftp_username}\n", f"password={ftp_password}\n", f"url={url}"])
    update_button['state'] = "disabled"
    update_button["text"] = "Updated Successfully"
    update_button.place(x=95, y=310)
    update_button.update_idletasks()

def settings():
    global update_button, host_entry, port_entry, path_entry, username_entry, password_entry, url_entry
    settings_window = tk.Toplevel()
    settings_window.title("Settings")
    settings_window.geometry("300x360")
    settings_window.resizable(False, False)
    settings_window.wm_iconphoto(False, tk.PhotoImage(file="icon.png"))
    ftp_heading_label = ttk.Label(settings_window, text="Settings", font=("Helvetica", 13))
    ftp_heading_label.pack()
    ftp_separator = ttk.Separator(settings_window, orient="horizontal")
    ftp_separator.place(x=0, y=30, width=300)
    host_label = ttk.Label(settings_window, text="FTP Host :", anchor='w', font=("Helvetica", 10)).place(x=20, y=60)
    port_label = ttk.Label(settings_window, text="FTP Port :", anchor='w', font=("Helvetica", 10)).place(x=20, y=100)
    path_label = ttk.Label(settings_window, text="FTP Path :", anchor='w', font=("Helvetica", 10)).place(x=20, y=140)
    username_label = ttk.Label(settings_window, text="Username :", anchor='w', font=("Helvetica", 10)).place(x=20, y=180)
    password_label = ttk.Label(settings_window, text="Password :", anchor='w', font=("Helvetica", 10)).place(x=20, y=220)
    url_label = ttk.Label(settings_window, text="Upload URL :", anchor='w', font=("Helvetica", 10)).place(x=20, y=260)
    host_entry = ttk.Entry(settings_window)
    host_entry.insert(0, ftp_host)
    host_entry.place(x=120, y=60, width=150, height=25)
    port_entry = ttk.Entry(settings_window)
    port_entry.insert(0, ftp_port)
    port_entry.place(x=120, y=100, width=150, height=25)
    path_entry = ttk.Entry(settings_window)
    path_entry.insert(0, ftp_path)
    path_entry.place(x=120, y=140, width=150, height=25)
    username_entry = ttk.Entry(settings_window)
    username_entry.insert(0, ftp_username)
    username_entry.place(x=120, y=180, width=150, height=25)
    password_entry = ttk.Entry(settings_window)
    password_entry.insert(0, ftp_password)
    password_entry.place(x=120, y=220, width=150, height=25)
    url_entry = ttk.Entry(settings_window)
    url_entry.insert(0, url)
    url_entry.place(x=120, y=260, width=150, height=25)
    update_button = ttk.Button(settings_window, text="Update Settings", command=update_settings)
    update_button.place(x=100, y=310)

settings_window_btn = ttk.Label(window, text="Settings", font=("Helvetica", 8), foreground="blue", cursor="hand2")
settings_window_btn.bind("<Button-1>", lambda event: settings())
settings_window_btn.place(x=450, y=482)

window.mainloop()
