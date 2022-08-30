# RS-232 -> Excel
The objective of the project is to obtain information from a serial port and translate it into excel templates configured for voltage and current measurement purposes.

## Install
<b>Python 3.9</b><br>
<ul>
<li>git clone </li>
<li>Create a virtual environment in the project folder</li>
<li>(With the venv active) pip install -r requirements.txt</li>
</ul>

## Running the app
<ul>
<li>run: python app.py</li>
<li>Warning: It will collect data from a serie port until "\r\n\r\n\r\n" is recive with the specific data save in "DATA/"</li>
</ul>

## Pyinstaller 
I added a folder called "RS-232" which contains an .exe that allows running the program without the need for a Python interpreter. It was posible thank to a library call "Pyinstaller"
