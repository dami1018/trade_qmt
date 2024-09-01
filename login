# 来源于：https://gitee.com/li-xingguo11111 小果量化

import time  # 导入时间模块，用于处理时间相关的操作
import pyautogui as pa  # 导入pyautogui模块，用于自动化GUI操作
import pywinauto as pw  # 导入pywinauto模块，用于自动化Windows应用程序
import schedule  # 导入schedule模块，用于定时任务调度
import yagmail  # 导入yagmail模块，用于发送邮件
import requests  # 导入requests模块，用于发送HTTP请求
import json  # 导入json模块，用于处理JSON数据
import random  # 导入random模块，用于生成随机数
from datetime import datetime  # 从datetime模块导入datetime类，用于处理日期时间

class qmt_auto_login:
    '''
    qmt自动登录
    '''
    def __init__(self,connect_path = r'D:\国金证券QMT交易端\bin.x64\XtItClient.exe',
                user = '',
                password = '',
                seed_type='qq',
                seed_qq='1752515969@qq.com',
                qq_paasword='jgyhavzupyglecaf',
                access_token_list=['1029762153@qq.com']):
        '''
        connect_path qmt安装路径
        user股票账户
        password密码
        seed_type发送方式 wx微信,qq qq,dd钉钉
        seed_qq发送qq
        qq_paasword QQ掩码
        access_token_list账户token
        特别在提醒，在服务器运行的qmt，退出服务器需要点击文件夹下面的退出保持链接bat退出服务器
        '''
        self.connect_path=connect_path  # 设置qmt安装路径
        self.user=user  # 设置股票账户
        self.password=password  # 设置密码
        self.app = None  # 初始化应用程序对象
        self.seed_type=seed_type  # 设置发送方式
        self.access_token_list=access_token_list  # 设置账户token列表
        self.seed_qq=seed_qq  # 设置发送qq
        self.qq_paasword=qq_paasword  # 设置QQ掩码

    def seed_dingding(self,msg='买卖交易成功,',access_token_list=['ab5d0a609429a786b9a849cefd5db60c0ef2f17f2ec877a60bea5f8480d86b1b']):
        '''
        发送钉钉
        '''
        access_token=random.choice(access_token_list)  # 随机选择一个access_token
        url='https://oapi.dingtalk.com/robot/send?access_token={}'.format(access_token)  # 构建请求URL
        headers = {'Content-Type': 'application/json;charset=utf-8'}  # 设置请求头
        data = {
            "msgtype": "text",  # 发送消息类型为文本
            "at": {
                #"atMobiles": reminders,
                "isAtAll": False,  # 不@所有人
            },
            "text": {
                "content": msg,  # 消息正文
            }
        }
        r = requests.post(url, data=json.dumps(data), headers=headers)  # 发送POST请求
        text=r.json()  # 获取响应的JSON数据
        errmsg=text['errmsg']  # 获取错误信息
        if errmsg=='ok':
                print('钉钉发生成功')  # 打印成功信息
                return text  # 返回响应数据
        else:
            print(text)  # 打印错误信息
            return text  # 返回响应数据

    def get_seed_qq_test(self,text='交易完成',seed_qq='1752515969@qq.com',qq_paasword='jgyhavzupyglecaf',re_qq='1029762153@qq.com'):
        '''
        发送qq，可以自己给自己发
        text发送内容，支持表格文字
        see_qq发送QQ
        qq_pasword发送qq的掩码
        re_qq接收信息QQ
        '''
        try:
            yag = yagmail.SMTP(user=seed_qq, password=qq_paasword, host='smtp.qq.com')  # 初始化SMTP对象
            text = text  # 设置发送内容
            yag.send(to=re_qq, contents=text, subject='邮件')  # 发送邮件
            print('qq发送完成')  # 打印成功信息
        except:
            print('qq发送失败')  # 打印失败信息

    def seed_wechat(self, msg='买卖交易成功,', access_token_list=[]):
        access_token=random.choice(access_token_list)  # 随机选择一个access_token
        url = 'https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=' + access_token  # 构建请求URL
        headers = {'Content-Type': 'application/json;charset=utf-8'}  # 设置请求头
        data = {
            "msgtype": "text",  # 发送消息类型为文本
            "at": {
                # "atMobiles": reminders,
                "isAtAll": False,  # 不@所有人
            },
            "text": {
                "content": msg,  # 消息正文
            }
        }
        r = requests.post(url, data=json.dumps(data), headers=headers)  # 发送POST请求
        text = r.json()  # 获取响应的JSON数据
        errmsg = text['errmsg']  # 获取错误信息
        if errmsg == 'ok':
            print('wechat发生成功')  # 打印成功信息
            return text  # 返回响应数据
        else:
            print(text)  # 打印错误信息
            return text  # 返回响应数据

    def seed_info(self,text=''):
        '''
        发送信息
        '''
        
        if self.seed_type=='qq':
            self.get_seed_qq_test(text=text,seed_qq=self.seed_qq,qq_paasword=self.qq_paasword,
                                  re_qq=self.access_token_list[-1])  # 根据发送方式选择发送qq
        elif self.seed_type=='wx':
            self.seed_wechat(msg=text,access_token_list=self.access_token_list)  # 根据发送方式选择发送微信
        elif self.seed_type=='dd':
            self.seed_dingding(msg=text,access_token_list=self.access_token_list)  # 根据发送方式选择发送钉钉
        else:
            self.get_seed_qq_test(text=text,seed_qq=self.seed_qq,qq_paasword=self.qq_paasword,
                                  re_qq=self.access_token_list[-1])  # 默认发送qq

    def send_vx_msg(self,msg):
        print(msg)  # 打印消息

    def login(self):
        test_app = pw.application.Application(backend="uia")  # 初始化应用程序对象
        try:
            # 获取test.exe的process id
            proc_id = pw.application.process_from_module("XtItClient.exe")  # 获取进程ID
            print('proc_id:', proc_id)  # 打印进程ID

            # 关联应用程序进程
            self.app = test_app.connect(process=proc_id)  # 连接到应用程序
            self.app.top_window().dump_tree()  # 打印窗口树
            self.app.kill()  # 关闭应用程序
        except Exception:
            pass
        # if app is None:
        self.app = pw.Application(backend='uia').start(self.connect_path, timeout=10)  # 启动应用程序
        time.sleep(5)  # 等待5秒
        self.app.top_window()  # 获取顶层窗口
        time.sleep(5)  # 等待5秒
        pa.typewrite(self.user)  # 输入用户名
        time.sleep(1)  # 等待1秒
        pa.hotkey('tab')  # 按下tab键
        time.sleep(1)  # 等待1秒
        pa.typewrite(self.password)  # 输入密码
        time.sleep(1)  # 等待1秒
        pa.hotkey('enter')  # 按下回车键
        time.sleep(3)  # 等待3秒
        # 判断是否成功 WindowSpecification
        login_window = self.app.window_(title="国金证券QMT交易端 1.0.0.29456", control_type="Pane")  # 获取登录窗口
        try:
            login_window.wait('visible', timeout=1)  # 等待窗口可见
            text='{} qmt登录失败'.format(datetime.now())  # 构建失败消息
            self.seed_info(text=text)  # 发送失败消息
            self.send_vx_msg('登录失败！')  # 打印登录失败消息
        except (pw.findwindows.ElementNotFoundError, pw.timings.TimeoutError):
            text='{} qmt登录成功'.format(datetime.now())  # 构建成功消息
            self.seed_info(text=text)  # 发送成功消息
            print('登录成功！')  # 打印登录成功消息

    def kill(self):
        '''
        退出程序
        '''
        self.app.kill()  # 关闭应用程序

if __name__=='__main__':
    models=qmt_auto_login(connect_path = r'D:\国金证券QMT交易端\bin.x64\XtItClient.exe',
                user = '',
                password = '',
                seed_type='qq',
                seed_qq='1752515969@qq.com',
                qq_paasword='jgyhavzupyglecaf',
                access_token_list=['1029762153@qq.com'])  # 初始化qmt_auto_login对象
    test='True'  # 设置测试标志
    if test=='True':
        print('测试登录')  # 打印测试登录消息
        models.login()  # 调用登录方法
        models.kill()  # 调用退出方法
    else:
        #定时登录
        schedule.every().day.at('{}'.format('09:10')).do(models.login)  # 设置定时登录任务
        #定时退出
        schedule.every().day.at('{}'.format('15:3')).do(models.kill)  # 设置定时退出任务
    while True:
        schedule.run_pending()  # 运行待处理的任务
        time.sleep(1)  # 等待1秒
