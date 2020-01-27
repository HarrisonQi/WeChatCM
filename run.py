# 导入模块
import random

import xlrd
from wxpy import *

from errors import data_error


def init_workbook():
    # 读取人脉表格文件
    return xlrd.open_workbook('Connections.xls')


# 获取用户集合, 数据格式:[{备注:备注, 称呼:称呼, 自称:自称, 消息模板:消息模板, 自定义消息:自定义消息}]
def init_user_list():
    # 定位到用户sheet
    sheet_user = init_workbook().sheet_by_name('通讯录')
    user_list = []
    for i in range(1, sheet_user.nrows):
        # 判断是否设置发送
        if sheet_user.cell(i, 5).value == 1:
            user_list.append({
                '备注': sheet_user.cell(i, 0).value,
                '称呼': sheet_user.cell(i, 1).value,
                '自称': sheet_user.cell(i, 2).value,
                '消息模板': sheet_user.cell(i, 3).value,
                '自定义消息': sheet_user.cell(i, 4).value,
            })

    return user_list


def init_msg_template():
    # 存储消息模板为字典, 模板名为key, 候选集合作为value
    # 定位到消息模板sheet
    sheet_msg_template = init_workbook().sheet_by_name('消息模板')
    # 确定模板数量
    cols = sheet_msg_template.col(0)

    # 根据模板名集合遍历行, 存入字典
    msg_template_dic = {}
    for i, col in enumerate(cols):
        # 获取候选集合
        candidate_list = sheet_msg_template.row_values(i, 1, 5)
        # 去除空元素
        candidate_list = [i for i in candidate_list if (len(str(i)) != 0)]
        msg_template_dic[col.value] = candidate_list
    return msg_template_dic


def main():
    # 检查Excel文件
    # TODO 检查是否有重复消息模板名称

    # 初始化机器人，扫码登陆
    bot = Bot()

    # 读取人脉表格文件
    workbook = init_workbook()

    # 读取消息模板表
    msg_template = init_msg_template()

    # 读取用户表
    user_list = init_user_list()
    print('user_list', user_list)

    # 定位到通讯录sheet
    sheet_txl = workbook.sheet_by_name('通讯录')

    # 遍历用户集
    for user in user_list:
        # 获取目标备注名
        # remarks = '自己的'
        remarks = user['备注']

        # 获取称呼
        goal_call = user['称呼']
        # 获取自称
        self_call = user['自称']
        # 获取消息模板
        msg_template_name = user['消息模板']
        # 获取自定义消息
        customize_msg = user['自定义消息']

        # 最终发送信息
        if len(customize_msg.strip()) != 0:
            # 已定义自定义消息:
            final_msg = customize_msg
        elif len(msg_template_name.strip()) != 0:
            # 使用消息模板
            # 根据消息模板名获取候选集合
            candidate_list = msg_template[msg_template_name.strip()]
            # 判断是否未找到
            if len(candidate_list) == 0:
                raise data_error('未找到模板')
            # 随机使用一个候选
            final_msg = random.sample(candidate_list, 1)[0]

        else:
            # 未自定义消息且未定义消息模板
            raise data_error('未自定义消息且未定义消息模板')
            pass

        final_msg = final_msg.replace('${称呼}', goal_call).replace('${自称}', self_call)

        # 设置发送目标
        sending_goal = bot.friends().search(remarks)[0]

        # 发送文本给好友
        print('sending_goal %s' % sending_goal)
        print('final_msg %s' % final_msg)

        # 发送!
        sending_goal.send(final_msg)

    # 保持运行
    # embed()


main()
