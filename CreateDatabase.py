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
# mycursor.execute("INSERT INTO SUBJECTS VALUES('全学共通教育科目','春夏学期','月1','スポーツ方法（春夏）Ⅰ','平賀　慧','日','○')")
# mycursor.execute("INSERT INTO SUBJECTS VALUES('全学共通教育科目','春夏学期','月1','教養ゼミナール（小泉　順也）Ａ','小泉　順也','日','-')")


# ↓　"DESCRIBE table名"でそのテーブルの現在情報を表示する
mycursor.execute("SELECT * FROM SUBJECTS")

# ↓　"DESCRIBE table名"で表示されたテーブル情報をプリントする
for x in mycursor:
    print(x)
