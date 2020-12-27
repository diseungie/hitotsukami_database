import mysql.connector

# ↓　passwordはMySQLをダウンロードするときに設定したやつ、他の人のパソコンでも同じPWで作動するかはよくわかんない
# ↓　databaseは作成したdatabase（今は俺がtestdatabaseというdatabaseを作った状態）
db = mysql.connector.connect(
    host="localhost",
    user="root",
    password="632091",
    database="testdatabase"
    )

mycursor = db.cursor()

# ↓　""の中にMySQLの言語を書く（CREATE　DATABASEとか）、そしてRunする
mycursor.execute("")


# ↓　"DESCRIBE table名"でそのテーブルの現在情報を表示する
# mycursor.execute("DESCRIBE Classes")

# ↓　"DESCRIBE table名"で表示されたテーブル情報をプリントする
# for x in mycursor:
#     print(x)
