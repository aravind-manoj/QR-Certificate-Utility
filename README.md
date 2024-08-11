# Certificate Utility Tool
This Python application is a powerful tool designed to generate certificates with unique QR codes embedded in them. The QR codes are generated based on specific values such as the certificate number and other customizable fields. The tool features an intuitive graphical user interface (GUI) built using Tkinter and includes functionality to upload and manage certificates on a web server via FTP.

## Features
**User-Friendly GUI:** Built using Tkinter, making it easy to navigate and use.
**Dynamic QR Code Generation:** Each certificate includes a QR code that is dynamically generated based on the certificate number and other user-defined values.
**FTP Upload and Management:** Upload certificates directly to a web server via FTP. You can also remove certificates from the server when needed.
**PDF Export:** Certificates will be automatically exported to PDF format.

##Installation

Clone the repository:
```
git clone https://github.com/aravind-manoj/qr-certificate-utility.git
cd certificate-creation-tool
```

Install the required dependencies:
```
pip install -r requirements.txt
```
Run the application:
```
python certificate_creator.py
```
Requirements:
```
tkinter (for GUI)
qrcode (for QR code generation)
ftplib (for FTP functionality)
docx and docxtpl (for docx handling)
pywin32 (for accessing Windows API for converting docx to pdf)
```

## Usage

Launch the program using the command:
```
python app.py.
```
Fill in the necessary fields such as the recipient's name, place, date, certificate number and id.
Click on "Generate Certificate" to create the certificate with the embedded QR code.
If you are running it for first time, then you must configure settings first. Click on Settings on bottom of the app and fill the empty fields.
```
host - Host to connect (Required for FTP)
port - Port to connect (Required for FTP)
path - Path to save file on remote server (Required for FTP)
username - Username for FTP (Required for FTP)
password - Password for FTP (Required for FTP)
url - Base URL for creating QRcode (Required for Creating Certificate)
```
Use the FTP options to upload the certificate to your web server or remove an existing certificate from the server.
NOTE: Currenly this program works only on Windows operating systems. Also certificate generation will fails if you don't have MS Office installed on your PC.

## Screenshots
<img width="499" alt="shot1" src="https://github.com/user-attachments/assets/dcd7d2fe-789d-43a2-8852-3c97bd773931">
<img width="499" alt="shot2" src="https://github.com/user-attachments/assets/6e21e721-1097-48c1-aae1-f05f89c3024e">
<img width="499" alt="shot3" src="https://github.com/user-attachments/assets/446378ed-0297-4d82-9924-9affd3976918">


## Contributing

Contributions are welcome! Please open an issue or submit a pull request if you have any improvements or new features to suggest.

## License

This project is licensed under the MIT License.
