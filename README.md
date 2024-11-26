# TriboReader
This is a program for automation and processing of data from tribometers.<br>
For now, supported tribometers is: Anton Paar TRB3, Nano TRB (NTR), T11, Rtec (HW ver. 2022).
<hr>
How to use:<br>
Just put a .txt or .csv file to the folder where is a program, edit config file (optional) and run it.<br>
<br>
How to run code:<br>
To run a code from .py file only You need to do a first two points from below.<br>
<br>
How to compile to .exe:<br>
To compile a .py file to .exe, only You need to do a four steps:<br>
1. Install python 3 from Microsoft Store or from https://www.python.org/downloads/<br>
2. In the Terminal: <code>pip install pandas xlsxwriter scipy pyinstaller</code><br>
3. Next command: <code>curl -O https://raw.githubusercontent.com/popos123/TriboReader/refs/heads/main/_TriboReader1.75.py</code><br>
4. Final compilation: <code>pyinstaller --onefile _TriboReader1.75.py</code><br>
<br>
How to compile to .deb:<br>
By using stdeb, tutorial: https://www.youtube.com/watch?v=nEg9VLK47xQ<br>
