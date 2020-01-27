# Wechat Connections Maintainer - 微信人脉维持机
---

**是否有过这样的困扰:**
1. 过年了, 想给大家拜个年, 我几百上千个微信好友, 群发又不合适...
2. 女朋友明天过生日, 但是明天还要上班, 凌晨给她发生日祝福太辛苦

**这个项目就是为了解决以上痛点诞生的!**
1. 只需要输入祝福语, 将带上昵称自动发送给每一个人! 支持昵称定制!
2. 定时自动发送消息

# 使用方法

1. 引入wxpy
    ```bash
    pip install -U wxpy
    ```
2. 引入openxl
    ```bash
    pip3 install openxl
    ```
3. 运行init.py, 进行通讯录初始化
   
   `你会在此得到一个新生成名为Connections的文件, 打开后可以看到两个sheet: 通讯录及消息模板`
4. 填写通讯录表格
    
   `最左边会把你所有的联系人生成, 参照下面的表格进行填写`

    |  备注名   | 你在微信给对方的备注  |
    |  ----  | ----  |
    | 称呼  | 私下里你如何称呼对方 |
    | 自称  | 私下里对方如何称呼你 |
    | 消息模板  | 发送的消息模板名 |
    | 自定义消息  | 不使用模板, 使用单独的自定义消息 |
    | 是否发送  | 1发送0不发送 |

5. 增加消息模板

    |  模板名   | 自定义模板名, 不可有重名模板  |
    |  ----  | ----  |
    | 候选消息1-5 | 你可以在一个模板下创建多个候选消息, 将会随机挑选一个 |
6. 消息模板可用的变量

    |  模板名   | 自定义模板名, 不可有重名模板  |
    |  ----  | ----  |
    |  ${称呼}  | 通讯录填写的称呼  |
    |  ${自称}  | 通讯录填写的自称  |
    ```
    例如: ${称呼}, 最近怎么样? 许久未见, 甚是想念! ${自称}在此给你拜个年! 
    ```
   
# 一次填写, N次使用! 
# 节省时间, 维持人脉关系!

**此项目基于:**

[ youfou / wxpy ](https://github.com/youfou/wxpy)

# 注意: 注册未满两年的账号无法使用web微信登录, 故无法使用wxpy, 也无法使用此项目, 坑爹的腾讯!