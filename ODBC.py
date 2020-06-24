import pyodbc
import sqlite3
from Google_Sheets import Google_Sheets


con = sqlite3.connect('test.sqlite3')
cur = con.cursor()
cur.execute("SELECT * FROM product")
rows = cur.fetchall()

BD = Google_Sheets("database")

start = ["productid", "name", "weight", "price", "pricewithdelivery", "packcount", "vat"]
BD.add_start(start)
BD.add_info_bd(rows)



