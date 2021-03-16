import time
import sqlite3
import xlrd
import tablib

xlrdCtypeMap = ["#","text","number","datetime"]
openpyxlMap ={"s":"text","n":"number","d":"datetime"}
"""
def readXlsxFile(filePath):
    file = xlrd.open_workbook(filePath)
    table = file.sheets()[0]
    labelRow = table.row(0)
    firstRow = table.row(1)
    labels = []
    for cellIndex in range(len(labelRow)):
        ctpye = xlrdCtypeMap[firstRow[cellIndex].ctype]
        label = labelRow[cellIndex].value
        labels.append({
            "label":label,
            "type":ctpye
        })
    dataSet = []
    for rowIndex in range(1,table.nrows):
        nowRow = table.row(rowIndex)
        lineData=[]
        for cellIndex in range(len(labels)):
            if labels[cellIndex]["type"]=="datetime":
                dtValue = xlrd.xldate_as_datetime(nowRow[cellIndex].value,file.datemode)
                lineData.append(dtValue.strftime("%Y-%m-%d %H:%M:%S"))
            else:
                lineData.append(nowRow[cellIndex].value)
        dataSet.append(lineData)
    return labels,dataSet
"""

def readXlsxFile(filePath):
    file = openpyxl.load_workbook(filePath,data_only=True)
    table = file.active
    rows = [col for col in table.rows]
    labelRow = rows[0]
    firstRow = rows[1]
    labels = []
    for cellIndex in range(len(labelRow)):
        ctpye = openpyxlMap[firstRow[cellIndex].data_type]
        label = labelRow[cellIndex].value
        labels.append({
            "label": label,
            "type": ctpye
        })
    dataSet = []
    for rowIndex in range(1,table.max_row):
        nowRow = rows[rowIndex]
        lineData=[]
        for cellIndex in range(len(labels)):
            if labels[cellIndex]["type"]=="datetime":
                dtValue = nowRow[cellIndex].value
                lineData.append(dtValue.strftime("%Y-%m-%d %H:%M:%S"))
            else:
                lineData.append(nowRow[cellIndex].value)
        dataSet.append(lineData)
    return labels,dataSet

def genCreatTableSQL(labels,DBName):
    cols=["UID TEXT"]
    for index in range(1,len(labels)):
        if labels[index]["type"] in "text datetime":
            sqlType = "TEXT"
        else:
            sqlType = "REAL"
        cols.append("{}_{} {}".format(DBName,index,sqlType))
    sql = "CREATE TABLE {}(\n\t{}\n)".format(DBName,",\n\t".join(cols))
    return sql

def xlsx2SQLite(conn:sqlite3.Connection,cousor:sqlite3.Cursor,filePath,dbName):
    labels,dataSet = readXlsxFile(filePath)
    dbCraeteSQL = genCreatTableSQL(labels,dbName)
    cousor.execute(dbCraeteSQL)
    conn.commit()
    sql = "INSERT INTO {} VALUES ({})".format(dbName, ",".join(["?" for i in range(len(labels))]))
    for lineIndex in range(len(dataSet)):
        line = dataSet[lineIndex]
        cousor.execute(sql,line)
    conn.commit()
    return labels


class ExcelJoin:
    def __init__(self,debug=False,debugDBName=""):
        if debug:
            self.conn = sqlite3.connect(debugDBName)
        else:
            self.conn = sqlite3.connect(":memory:")
        self.couser = self.conn.cursor()
        self.labelsA =None
        self.labelsB =None
        self.isFileLoad = False

    def lodaXlsxFile(self,filepathA,filepathB):
        self.labelsA = xlsx2SQLite(self.conn,self.couser,filepathA,"dbA")
        self.labelsB = xlsx2SQLite(self.conn, self.couser, filepathB, "dbB")
        self.isFileLoad = True

    def joinA2B(self):
        if not self.isFileLoad:
            raise Exception("file not load")
        sql = "SELECT * FROM dbB as a LEFT JOIN dbA as b on a.UID=b.UID"
        self.couser.execute(sql)
        data = self.couser.fetchall()
        outlabel = []
        for label in self.labelsB:
            outlabel.append(label["label"])
        for label in self.labelsA:
            outlabel.append(label["label"])
        databook = tablib.Dataset(*data, headers=outlabel)
        t = time.strftime("%Y_%m_%d_%H_%M_%S")
        with open("outmin_{}.xlsx".format(t),"wb") as file:
            file.write(databook.xlsx)

    def joinB2A(self):
        if not self.isFileLoad:
            raise Exception("file not load")
        sql = "SELECT * FROM dbA as a LEFT JOIN dbB as b on a.UID=b.UID"
        self.couser.execute(sql)
        data = self.couser.fetchall()
        outlabel = []
        for label in self.labelsA:
            outlabel.append(label["label"])
        for label in self.labelsB:
            outlabel.append(label["label"])
        databook = tablib.Dataset(*data, headers=outlabel)
        t = time.strftime("%Y_%m_%d_%H_%M_%S")
        with open("outmax_{}.xlsx".format(t),"wb") as file:
            file.write(databook.xlsx)


    def findANotInB(self):
        if not self.isFileLoad:
            raise Exception("file not load")
        sql = "SELECT * FROM dbA as a WHERE a.UID NOT IN (SELECT UID FROM dbB as b)"
        self.couser.execute(sql)
        data = self.couser.fetchall()
        outlabel = []
        for label in self.labelsA:
            outlabel.append(label["label"])
        databook = tablib.Dataset(*data,headers=outlabel)
        t = time.strftime("%Y_%m_%d_%H_%M_%S")
        with open("outmiss_{}.xlsx".format(t), "wb") as file:
            file.write(databook.xlsx)

if __name__ == '__main__':
    excelJoin = ExcelJoin()
    excelJoin.lodaXlsxFile("test\\a.xlsx","test\\b.xlsx")
    excelJoin.joinA2B()
    excelJoin.joinB2A()
    excelJoin.findANotInB()

