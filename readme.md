# How to publish python code to app:

1) cd "C:\Users\HP\Documents\GitHub\DC-app"
*Location python file*

2) C:/Users/HP/AppData/Local/Programs/Python/Python312/python.exe -m PyInstaller --noconsole --onefile --collect-all customtkinter --collect-all tkcalendar --hidden-import psycopg2 --hidden-import psycopg2.extensions --hidden-import psycopg2._psycopg dcapp.py
*Build app*

3) C:\Users\HP\Documents\GitHub\DC-app\dist
*Open installed folder*