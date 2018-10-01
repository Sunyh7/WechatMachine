import itchat
import time

@itchat.msg_register(itchat.content.TEXT,isGroupChat=True)
def print_content(msg):
    print(msg['Text'])

@itchat.msg_register(itchat.content.PICTURE)
def print_image(msg):
    print(msg['Text'])


itchat.auto_login(hotReload=True)
#print(itchat.search_chatrooms(name="数院2017级二班")['UserName'])

for room in itchat.get_chatrooms(update=True):
    if room['NickName'] == "数院2017级二班":
        print(room['UserName'])
        #userName = room['UserName']

userName='@@0bb34b6af41a6ea8f5a1f5f632da2bec770e0a52616f13c4977d08c1d75f1ef7'   #5
erban = '@@eaf74e43a28d3079c64f0a71f0802063e43b0ba5ac9b979eb3b4eadb618580b0'    #二班班级群
itchat.send_msg(msg="",toUserName=userName)

itchat.run()
