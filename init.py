import xlwt
from wxpy import *

# 初始化人脉表格记录
wbk = xlwt.Workbook()
sheet = wbk.add_sheet('通讯录')
sheet.write(0, 0, '备注名')
sheet.write(0, 1, '称呼')
sheet.write(0, 2, '自称')
sheet.write(0, 3, '消息模板')
sheet.write(0, 4, '自定义消息')
sheet.write(0, 5, '是否发送(1是0否)')
sheet.write(0, 6, '上次发送时间')
sheet.write(0, 7, '上次发送内容')

message_sheet = wbk.add_sheet('消息模板')
message_sheet.write(0, 0, '模板名称')
message_sheet.write(0, 1, '候选消息1')
message_sheet.write(0, 2, '候选消息2')
message_sheet.write(0, 3, '候选消息3')
message_sheet.write(0, 4, '候选消息4')
message_sheet.write(0, 5, '候选消息5')


# 使用机器人, 遍历所有好友并写入Excel
bot = Bot()
friends = bot.friends()
for i, friend in enumerate(friends):
    # 生成第一列: 备注名
    sheet.write(i + 1, 0, friend.name)

    # 生成第五列: 是否发送(1发送, 0不发送)
    sheet.write(i + 1, 5, 0)

wbk.save('Connections.xls')
