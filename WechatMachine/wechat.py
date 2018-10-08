import itchat
import time
import xlrd
import xlsxwriter
from xlutils.copy import copy

send = 1
state = 0
exchangeHour = 0
exchangeStudents = []
studentsLabor,students = [],[]

chatRoom = u'数院2017级二班'

def read():
    global state
    file = xlrd.open_workbook('17级2班擦黑板.xls')
    table = file.sheet_by_index(0)
    data = []
    for i in range(1,55):
        number = int(table.cell_value(i,0))
        name = table.cell_value(i,1)
        turn = int(table.cell_value(i,2))
        date = []
        for j in range(table.ncols - 3):
            try:
                date.append(int(table.cell_value(i,j + 3)))
            except:
                pass
        temp = {'number':number,'name':name,'turn':turn,'date':date}
        data.append(temp)
    for i in range(0,53):
        if data[i]['turn'] > data[i + 1]['turn']:
            state = i + 1
    return data

def save(data,sheet=0,row=2,col=1):
    xls = xlrd.open_workbook('17级2班擦黑板.xls')
    xlsc = copy(xls)
    sheet1 = xlsc.get_sheet(0)
    for i in range(54):
        j = 0
        for t in data[i].values():
            if j == 3:
                for n in range(len(t)):
                    sheet1.write(i + 1, j + n, t[n])
            else:
                sheet1.write(i + 1, j, t)
                j += 1
    xlsc.save('17级2班擦黑板.xls')

def getChatroom(name):
    for room in itchat.get_chatrooms(update=True):
        if room['NickName'] == name:
            return room['UserName']

def getStudent(msg):
    userName = msg.get('ActualUserName',None)
    student = itchat.search_friends(userName=userName)   #找学生
    return student

@itchat.msg_register(itchat.content.TEXT,isGroupChat=True)
def LabourBot(msg):
    global send
    global state
    global exchangeHour
    global exchangeStudents
    global studentsLabor
    global students
    exchange = []

    data = None
    hour = time.localtime(time.time())[3]


    if hour >= 19 and send == 1:    #发通知

        data = read()
        studentsLabor = [data[state]['name'],data[(state + 1) % 54]['name']]
        students = [data[state]['name'],data[(state + 1) % 54]['name']]
        #print(students)
        message = students[0] + '和' + students[1] + '同学，明天将轮到你们擦黑板[爱心]。请回复“我会好好擦黑板”，否则将视作缺勤。如果需要更换人员，请按此格式回复（“张三换成李四”）。'
        #print(message)
        itchat.send_msg(msg=message,toUserName=getChatroom(chatRoom))

        #banzhang = itchat.search_friends(remarkName="张健")
        #message = '请您发送群公告提醒值日信息。'
        #itchat.send_msg(msg=message,toUserName=banzhang[0]['UserName'])

        for stu in data:
            if stu.get('name',None) in studentsLabor:
                stu['turn'] += 1
        print(data[state],data[(state + 1) % 54])

        send = 0
        state = (state + 2) % 54
    if hour < 19:
        send = 1
    #print(msg['Text'])

    if '我会好好擦黑板' in msg['Text']:        #处理回复
        data = read()

        student = getStudent(msg)

        print(studentsLabor)

        if student.get('RemarkName',None) in studentsLabor:
            for stu in data:
                if stu['name'] == student.get('RemarkName',None):
                    stu['date'].append(int(time.strftime("%m%d%Y", time.localtime())))
                    index = studentsLabor.index(stu['name'])
                    studentsLabor.pop(index)
                    message = stu['name'] + "已确认。"
                    #print(message)
                    itchat.send_msg(msg=message,toUserName=getChatroom(chatRoom))
        print(studentsLabor)
        for stu in data:
            if stu['name'] in students:
                print(stu)

    if '换成' in msg['Text']:          #处理换人

        flag = 0
        data = read()
        names = msg['Text'].split('换成')          # 0-本人 1-他人
        print(names)
        if len(names) == 2:

            student = getStudent(msg)

            for stu in data:
                if stu['name'] == names[1]:
                    flag += 1
                if stu['name'] == names[0]:
                    flag += 1

            if flag < 2:
                itchat.send_msg(msg="请输入正确名字",toUserName=getChatroom(chatRoom))

            if flag == 2 and student.get('RemarkName',None) in studentsLabor and student.get('RemarkName',None) == names[0]:
                exchangeStudents = names
                exchangeHour = time.localtime(time.time())[3]
                itchat.send_msg(msg="请" + names[1] + "同学在2小时内回复“确认”，否则视作换人无效。",toUserName=getChatroom(chatRoom))
                print(exchange)

    if "确认" in msg['Text']:         #确认换人

        data = read()
        student = getStudent(msg)
        if exchangeStudents != []:
            if student.get('RemarkName',None) == exchangeStudents[1] and time.localtime(time.time())[3] - exchangeHour < 3:
                for stu1 in data:
                    if stu1['name'] == exchangeStudents[0]:
                        for stu2 in data:
                            if stu2['name'] == exchangeStudents[1]:
                                index = students.index(stu1['name'])
                                students.pop(index)
                                students.append(stu2['name'])
                                index = studentsLabor.index(stu1['name'])
                                studentsLabor.pop(index)
                                studentsLabor.append(stu2['name'])
                                temp = stu1
                                stu1 = stu2
                                stu2 = temp
                                stu1['turn'] += 1
                                stu2['turn'] -= 1
                                message = "交换成功，" + students[0] + "和" + students[1] + '同学，明天将轮到你们擦黑板[爱心]。请回复“我会好好擦黑板”，否则将视作缺勤。'
                                #exchangeStudents = []
                                itchat.send_msg(msg=message,toUserName=getChatroom(chatRoom))
                                break
                        break

            for stu in data:
                if stu['name'] in exchangeStudents:
                    print(stu)
            exchangeStudents = []
            
    if "没做" in msg['Text']:
        
        data = read()
        student = getStudent(msg)
        if student.get('RemarkName',None) == "孙奕华":
            pass

    if data is not None:
        save(data)

itchat.auto_login(enableCmdQR=True, hotReload = True)

#itchat.auto_login(hotReload = True)
itchat.run()
