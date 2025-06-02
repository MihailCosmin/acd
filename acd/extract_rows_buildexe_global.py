# Cosmin - 12.10.2021
from time import time
from time import sleep

from os import system
from os import remove
from os.path import isdir
from os.path import isfile

from shutil import rmtree
from shutil import copyfile

startTime = time()

if isdir("venv"):
    system("cmd /c venv\\Scripts\\activate")
    system("cmd /c pyinstaller --onefile --paths venv\\Lib\\site-packages extract_rows.py")
else:
    system("cmd /c virtualenv venv")
    system("cmd /c venv\\Scripts\\python.exe -m pip install --upgrade pip")
    system("cmd /c venv\\Scripts\\activate")
    sleep(1)
    system("cmd /c venv\\Scripts\\python.exe -m pip install --upgrade pip")
    system("cmd /c venv\\Scripts\\python.exe -m pip install -r requirements.txt --no-cache-dir")
    system("cmd /c pyinstaller --onefile --paths venv\\Lib\\site-packages extract_rows.py")

print(f"Building extract_rows.exe finished in {int(time() - startTime)} seconds")

if isdir("__pycache__"):
    rmtree("__pycache__")
if isdir("build"):
    rmtree("build")
if isfile("dist/extract_rows.exe"):
    copyfile("dist/extract_rows.exe", "extract_rows.exe")
if isdir("dist"):
    rmtree("dist")
if isfile("extract_rows.spec"):
    remove("extract_rows.spec")
if isdir("venv"):
    rmtree("venv")

sleep(600)
