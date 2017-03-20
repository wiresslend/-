# -*- coding:utf-8 -*-
import sys
import xlrd
import xlwt

reload(sys)
sys.setdefaultencoding('utf8')

class Lesson(object):
    def __init__(self, p_name, p_index):
        self.name = p_name
        self.index = p_index
    def get_name(self):
        return self.name
    def get_index(self):
        return self.index

def write_in_sheet(sheet_name,stu_name,stu_index,lesson_name,score,score1):
    cur_row = len(sheet_name.get_rows())
    sheet_name.write(cur_row+1, 0, stu_name)
    sheet_name.write(cur_row+1, 1, stu_index)
    sheet_name.write(cur_row+1, 2, lesson_name)
    sheet_name.write(cur_row+1, 3, score)
    sheet_name.write(cur_row+1, 4, score1)

data = xlrd.open_workbook('2013score.xls') #打开一张工作表
sheets = data.sheets()
new_workbook = xlwt.Workbook()
sheet1 = new_workbook.add_sheet('sheet1',cell_overwrite_ok=True)


for sheet in sheets:
    table = sheet  # 获取电子表格
    col = table.ncols #列
    row = table.nrows #行

    stu_name = ''   #学生姓名
    class_name = [] #课程名称
    result = {}     #生成的最终数据：{姓名：{科目：成绩}}


    print '当前表格有 %s 列' %col           #显示列数量
    print '当前表格有 %s 行' %row          #显示行数

    for i in range(2,col,2):
        A = Lesson(table.cell_value(2,i),[2,i])
        class_name.append(A)
    print class_name[1].name

    for i in class_name:
        for n in range(4,row):
            stu_name = table.cell_value(n,0)
            stu_index = table.cell_value(n,1)
            stu_score = table.cell_value(n,i.index[1])
            stu_score1 = table.cell_value(n,i.index[1]+1)
            write_in_sheet(sheet1,stu_name,stu_index,(i.name),stu_score,stu_score1)
new_workbook.save("test1.xls")


