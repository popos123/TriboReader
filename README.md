# TriboReader
This is a program for automation and processing of data from tribometers.

For now, supported tribometers is:
Anton Paar TRB3, Nano TRB, T11, Rtec

How to use:
Just put a .txt o .csv file to the folder where is a program and run it.

To compile a .py file to .exe, only You need to do a four steps:
1. Install python 3 from Microsoft Store or from https://www.python.org/downloads/
2. In the Terminal: pip install pandas xlsxwriter scipy pyinstaller
3. Next command: curl -O https://raw.githubusercontent.com/popos123/TriboReader/refs/heads/main/_TriboReader1.72.py
4. Final compilation: pyinstaller --onefile --noconsole --compress --name TriboReader _TriboReader1.72.py
