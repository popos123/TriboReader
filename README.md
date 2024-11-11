# TriboReader
This is a program for automation and processing of data from tribometers.<br>
For now, supported tribometers is: Anton Paar TRB3, Nano TRB, T11, Rtec.
<hr>
How to use:<br>
Just put a .txt o .csv file to the folder where is a program and run it.<br>
<br>
To run a code from .py file only You need to do a first two points from below.<br>
To compile a .py file to .exe, only You need to do a four steps:<br>
1. Install python 3 from Microsoft Store or from https://www.python.org/downloads/<br>
2. In the Terminal: <code>pip install pandas xlsxwriter scipy pyinstaller</code><br>
3. Next command: <code>curl -O https://raw.githubusercontent.com/popos123/TriboReader/refs/heads/main/_TriboReader1.72.py</code><br>
4. Final compilation: <code>pyinstaller --onefile _TriboReader1.72.py</code><br>
