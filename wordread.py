#作者：苏向阳
#平台：pycharm ，python3
#日期：2018.12.1
#功能：提取word表格内容到excel内。运行前需要先创建一个空的Excel文档，然后直接点击运行选择需要转换的word文档和需要保存到的Excel文件就行。
import win32com
from win32com.client import Dispatch, constants
from docx import Document
from tkinter.filedialog import askopenfilename
import tkinter.filedialog




def select_inv(con):#选择函数
    use = con.find("█"or" "and"√")
    unused = con.find("□")
    try:
        if unused > use :
            con = con [use+1:unused]
        else:
            con = con[use+1:]
    except:
        con = ""
    return (con)

def select_meet(con):#list1查找所有的会场，list2记录勾选的会场，list3记录勾选的会场在list1中的位置
    list1 = [i for i in range (len (con)) if
             (con[i:i + 1] == "□") or (con[i:i + 1] == "█")or(con[i:i + 1] == " ")or(con[i:i + 1] == "√")]
    list2 = [i for i in range (len (con)) if (con[i:i + 1] == "█")or(con[i:i + 1] == " ")or(con[i:i + 1] == "√")]
    list3 = [list1.index (list2[i]) for i in range (len (list2))]
    try:
        if len(list3)>2:#参加了一场还是两场
            con1 = con[list1[list3[0]] + 1:list1[list3[0] + 1]]
            if (list3[-1:][0] + 1) == len (list1):  # 判断最后一条是否为勾选项
                con2 = con[list1[list3[1]] + 1:]
            else:
                con2 = con[list1[list3[1]] + 1:list1[list3[1] + 1] - 1]
            con = con1 + con2
        else:
            con = con[list1[list3[0]] + 1:list1[list3[0] + 1]]
    except:
        con = ""

    return (con)

def select_type(con):#选择函数
    use = con.find("█普"and' 'and"√",2)
    unused = con.find("□")
    try:
        if unused > use :
            con = con [use+1:unused-1]
        else:
            con = con[use+1:]
    except:
        con =""
    return (con)


#选择需要打开的word文档
root = tkinter.Tk()
root.withdraw()#隐藏
path_word = askopenfilename(filetypes = (("docx", "*.docx*"),("all files", "*.*")))
root.destroy() #销毁
PATH=path_word
file = Document(PATH)
tables = file.tables
table = tables[0]

#通用信息
invoice_type2 = table.cell (2, 1).text + ""      #发票类型全部
invoice_type = select_type(invoice_type2)        #发票类型提取
company_name = table.cell (0, 3).text + ""
invoice_title = table.cell (1, 5).text + ""      #发票抬头
taxpayer_num = table.cell (2, 5).text + ""       #纳税人识别号
address_and_num = table.cell (3, 5).text + ""    #地址电话
bank = table.cell (4, 5).text + ""               #开户行
invoice_con2 = table.cell (5, 5).text + ""       #发票内容全部
invoice_con = select_inv(invoice_con2)           #发票内容提取

post_address = table.cell (6, 3).text + ""       #邮寄地址
title_speaker = table.cell (7, 3).text + ""      #报告题目及报告人
meeting_house2 = table.cell (8, 3).text + ""     #参加会场信息全部
meeting_house = select_meet(meeting_house2)      #参加会场信息提取

#参会人员信息
name1 = table.cell (11, 0).text + ""
gender1 = table.cell (11, 1).text + ""
pro_title1 = table.cell (11, 3).text + ""
tel_num1 = table.cell (11, 5).text + ""
postbox1 = table.cell (11, 6).text + ""
reg_data1 = table.cell (11, 8).text + ""

names = locals()

for i in range(5):
    if ((table.cell (12+i, 0).text + "")==""or (table.cell (12+i, 0).text + "")=='房 间 预 订'):
        break
    else:
        j=i+2

        names['name%s' % j] =table.cell (12+i, 0).text + ""
        names['gender%s' % j] = table.cell (12+i, 1).text + ""
        names['pro_title%s' % j] = table.cell (12+i, 3).text + ""
        names['tel_num%s' % j] = table.cell (12+i, 5).text + ""
        names['postbox%s' % j] = table.cell (12+i, 6).text + ""
        names['reg_data%s' % j] = table.cell (12+i, 9).text + ""

hotel1_prices1 = table.cell (12 + i + 3, 2).text + ""
hotel1_room_num = table.cell (12 + i + 3, 7).text + ""
hotel1_room_name = table.cell (12 + i + 3, 8).text + ""
for x in range (8):
    for y in range(i+1):
        if (table.cell (12 + y + 4, 7).text + "") != "__间":
            names['hotel1_con%s' % (y + 1)] = hotel1_room_num = (table.cell (12 + y + 4, 2).text + "") + (
                        table.cell (12 + i + 3, 8).text + "")
            names['hotel1_con%s' % (y + 1)] = hotel1_room_num = (table.cell (12 + y + 4, 2).text + "") + (
                        table.cell (12 + i + 3, 8).text + "")

#excel接口
w = win32com.client.Dispatch('Word.Application')
excel = win32com.client.Dispatch('Excel.Application')

#选择需要打开的excel文件
root = tkinter.Tk()
root.withdraw()#隐藏
path_excel = askopenfilename(filetypes = (("xlsx", "*.xlsx*"),("all files", "*.*")))
root.destroy() #销毁
workbook=excel.Workbooks.open(path_excel)
excel.Visible=False


first_sheet=workbook.Worksheets(1)
first_sheet.Cells(1,1).value="姓名"
first_sheet.Cells(1,2).value="性别"
first_sheet.Cells(1,3).value="单位"
first_sheet.Cells(1,4).value="职称"
first_sheet.Cells(1,5).value="电话"
first_sheet.Cells(1,6).value="邮箱"
first_sheet.Cells(1,7).value="报到日期"
first_sheet.Cells(1,8).value="发票抬头"
first_sheet.Cells(1,9).value="纳税人识别号"
first_sheet.Cells(1,10).value="地址电话"
first_sheet.Cells(1,11).value="开户行"
first_sheet.Cells(1,12).value="邮寄地址"
first_sheet.Cells(1,13).value="报告题目及报告人"

#有些是涂黑有些是打勾，识别起来会出错，第一个是自己提取的信息（有可能出差错），第二个是该栏的全部信息。
first_sheet.Cells(1,14).value="发票类型"
first_sheet.Cells(1,15).value="发票类型2"

first_sheet.Cells(1,16).value="发票内容"
first_sheet.Cells(1,17).value="发票内容2"

first_sheet.Cells(1,18).value="参加的会场"
first_sheet.Cells(1,19).value="参加的会场2"

for n in range(1,10000):
    if first_sheet.Cells(n,1).value==None:
        break
#添加信息
for h in range(i+1):
    first_sheet.Cells (h+n, 1).value = names['name%s' % (h+1)]
    first_sheet.Cells (h+n, 2).value = names['gender%s' % (h+1)]
    first_sheet.Cells (h+n, 3).value = company_name
    first_sheet.Cells (h+n, 4).value = names['pro_title%s' % (h+1)]
    first_sheet.Cells (h+n, 5).value = names['tel_num%s' % (h+1)]
    first_sheet.Cells (h+n, 6).value = names['postbox%s' % (h+1)]
    first_sheet.Cells (h+n, 7).value = names['reg_data%s' % (h+1)]

    first_sheet.Cells (h + n, 8).value = invoice_title
    first_sheet.Cells (h + n, 9).value = taxpayer_num
    first_sheet.Cells (h + n, 10).value = address_and_num
    first_sheet.Cells (h + n, 11).value = bank
    first_sheet.Cells (h + n, 12).value = post_address
    first_sheet.Cells (h + n, 13).value = title_speaker
    #可以选择是否输出这些信息
    # first_sheet.Cells (h + n, 14).value = invoice_type
    # first_sheet.Cells (h + n, 15).value = invoice_type2
    # first_sheet.Cells (h + n, 16).value = invoice_con
    # first_sheet.Cells (h + n, 17).value = invoice_con2
    # first_sheet.Cells (h + n, 18).value = meeting_house
    # first_sheet.Cells (h + n, 19).value = meeting_house2

workbook.Save()#保存到Excel












