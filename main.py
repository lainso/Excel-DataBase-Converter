import MySQLdb
from PySide2.QtGui import QIcon
from PySide2.QtUiTools import QUiLoader
from PySide2.QtWidgets import QMessageBox, QApplication
from sqlalchemy import create_engine
import pandas as pd

class DBExcel:

    def __init__(self):
        self.ui = QUiLoader().load("main.ui")
        self.ui.Connect.clicked.connect(self.Connect)
        self.ui.Import.clicked.connect(self.Import)
        self.ui.Export.clicked.connect(self.Export)

    def Connect(self):
        addr = self.ui.ipaddr.text()
        username = self.ui.DBUserName.text()
        passwd = self.ui.DBUserPasswd.text()
        DBname = self.ui.DBName.text()
        try:
            MySQLdb.connect(host=addr, user=username, passwd=passwd, db=DBname, charset="utf8")
            QMessageBox.about(self.ui, '连接情况', f'连接成功！')
        except Exception as e:
            QMessageBox.about(self.ui, '连接情况', f'连接失败！' + "\n" + str(e))

    def Import(self):
        # if the Connect button has not been clicked, then the program will not run
        if not self.ui.Connect.isEnabled():
            QMessageBox.about(self.ui, '连接情况', '请先连接数据库！')
            return
        addr = self.ui.ipaddr.text()
        port = self.ui.port.text()
        DBUserName = self.ui.DBUserName.text()
        DBUserPass = self.ui.DBUserPasswd.text()
        DBName = self.ui.DBName.text()
        tableName = self.ui.TableName.text()
        FileDir = self.ui.ImportDir.text()
        engine = create_engine('mysql+pymysql://'+DBUserName+':'+DBUserPass+'@'+ addr +':' + port+'/'+DBName)

        try:
            df = pd.read_excel(str(FileDir))
            df.to_sql(name=tableName, con=engine, index=False, if_exists='replace')
            QMessageBox.about(self.ui, '导入情况', f'导入成功！')
        except Exception as e1:
            QMessageBox.about(self.ui, '导入情况', f'导入失败！' + "\n" + str(e1))

    def Export(self):
        if not self.ui.Connect.isEnabled():
            QMessageBox.about(self.ui, '连接情况', '请先连接数据库！')
            return
        addr = self.ui.ipaddr.text()
        port = self.ui.port.text()
        DBUserName = self.ui.DBUserName.text()
        DBUserPass = self.ui.DBUserPasswd.text()
        DBName = self.ui.DBName.text()
        tableName2 = self.ui.tableName2.text()
        sql = f'SELECT * FROM ' + tableName2 + ';'
        exportDir = self.ui.exportDir.text()
        engine = create_engine('mysql+pymysql://' + DBUserName + ':' + DBUserPass + '@' + addr + ':' + port + '/' + DBName)

        try:
            db = pd.read_sql(sql=sql, con=engine)
            db.to_excel(exportDir)
            QMessageBox.about(self.ui, '导出情况', f'导出成功！')
        except Exception as e2:
            QMessageBox.about(self.ui, '导出情况', f'导出失败！' + "\n" + str(e2))

app = QApplication([])
app.setWindowIcon(QIcon('icon.ico'))
stats = DBExcel()
stats.ui.show()
app.exec_()
