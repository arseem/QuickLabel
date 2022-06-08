# **QR Label**
Automated 3-field labels generator with QR and barcode for every field<br>

## About the project
This project was created for a specific use in the network equipment warehouse. Main purpose was to make the process of labelization of devices quicker. Every piece of equipment needed it's id (Numer Seryjny), description (Opis) and warehouse id (Skład). Application can generate multiple ready-to-print A4 sheets filled with generated labels.<br>
GUI provided by Kivy is is responsive and scalable. Preview of the output sheets is available all the time during the creation process. 
## Features
- Importing data from a spreadsheet
- Manual data input
- Adding elements using HID barcode scanner
- Exporting collected data to spreadsheet
- Exporting sheets with labels to pdf
- Direct print from app

## Examples
### Manual input
![manual input example gif](https://github.com/arseem/3OSCx1F/blob/master/example/Signal%20sum%20example.gif "Manual input example")

### Spreadsheet import
![spreadsheet import example gif](https://github.com/arseem/3OSCx1F/blob/master/example/Filters%20example.gif "Spreadsheet import example")

### Scanner mode
![scanner mode example gif](https://github.com/arseem/3OSCx1F/blob/master/example/Filters%20example.gif "Scanner mode example")

## Technologies in use
- Python
  - Kivy (GUI)
  - Pandas (spreadsheets handling)
  - PIL (image manipulation)
  - QR and barcode
## Requirements

<details>
  <summary>Click to expand</summary>
  <ul>
    barcode-generator==0.1rc15<br>
    certifi==2022.5.18.1<br>
    charset-normalizer==2.0.12<br>
    colorama==0.4.4<br>
    cycler==0.11.0<br>
    docutils==0.18.1<br>
    et-xmlfile==1.1.0<br>
    idna==3.3<br>
    Kivy==2.0.0<br>
    kivy-deps.angle==0.3.2<br>
    kivy-deps.glew==0.3.1<br>
    kivy-deps.sdl2==0.3.1<br>
    Kivy-Garden==0.1.5<br>
    kiwisolver==1.4.2<br>
    matplotlib==3.3.4<br>
    numpy==1.22.4<br>
    openpyxl==3.0.10<br>
    pandas==1.2.5<br>
    Pillow==9.1.1<br>
    Pygments==2.12.0<br>
    pyparsing==3.0.9<br>
    pypiwin32==223<br>
    python-barcode==0.14.0<br>
    python-dateutil==2.8.2<br>
    pytz==2022.1<br>
    pywin32==301<br>
    qrcode==6.1<br>
    requests==2.27.1<br>
    six==1.16.0<br>
    urllib3==1.26.9<br>
    xlrd==2.0.1<br>
    XlsxWriter==3.0.3<br>
  </ul>
</details>

## How to use
- Make sure you have Python and venv library installed and added to PATH
### Windows
- Run setup.ps1
### Other OS
- Create virtual environment in the base folder of an application and activate it using<br>
  > pip -m venv venv<br>
  > venv/Scripts/Activate.ps1<br>
- Make sure to have installed all of the depandancies from requirements.txt<br>
  > pip install -r requirements.txt
- Run main.py<br>
  > cd src<br>python main.py


### Alternatively (without virtual environment)
- Make sure to have installed all of the depandancies from requirements.txt<br>
  > pip install -r requirements.txt
- Run src/main.py (making sure that root folder is your base)<br><br>

