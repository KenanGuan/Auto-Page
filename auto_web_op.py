# encoding=utf8
'''
脚本功能：
1. 从给定格式的Excel表格智能提取数据，智能填补全国技术合同网站上的表格；
2. 将表格中没有的买方信息（如具体地址、法人、联系方式、邮编等）从互联网智能爬取，并填录到网站中。

遗留问题：
1. 有一些潜在的bug并未修复
2. 自动搜索功能并不可靠，可以增加备选网站的个数以增加成功率；
3. 能否取代selenium功能，而使用自动爬虫（必须研究各个网站的反爬手段，需要费一些时间）
4. 买方所在行政区划代码、合同项目技术领域、社会经济目标等表项无法自动选择，需要人工选录。（excel中领域为可选项，对应固定领域）

开发者：Kanan.Guan
日期：2021.03.12
'''
from selenium import webdriver
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.support.select import Select
from webdriver_manager.chrome import ChromeDriverManager
import time
import xlrd
import datetime
from xlrd import xldate_as_tuple
from bs4 import BeautifulSoup
import requests
import random
import sys, os
from tkinter import *
import tkinter.messagebox

class auto_op:
    def __init__(self):
        self.driver = None # chrome driver
        self.contract_info_list = [] # 保存合同信息
        self.root = Tk() # 首页图形化工具
        self.fill_table = None # 功能页面图形化工具
        self.txt_prompt = None # 提示框
        self.height_width_Entry = None # 合同索引输入框
        self.excel_path_entry = None # Excel路径输入框
        self.excel_path = None # 默认存储Excel路径
        self.first_flag = True # 是否第一次填写合同
        self.fail_flag = 0 # 登陆失败状态量（0: 首次登陆；1:网络异常；2:验证码不正确）
        self.CheckVar1 = None # 是否启用自动搜索
        self.use_auto_search = None
        self.dic_tech_area = {'城市建设与社会发展':0, '电子信息':1, '航空航天':2, '核应用':3, '环境保护与资源综合利用':4, '农业':5, '生物、医药和医疗器械':6, '先进制造':7, '现代交通':8, '新材料及其应用':9, '新能源与高效节能':10, '测绘遥感':11, '水利水电':12, '电力行业':13, '计算机相关（及智能产业）':14, '其他':15}

    '''
    UI 控制函数
    '''
    # 进入程序首页
    def firstPage(self):
        self.root.title('全国技术合同登记系统——自动填写小助手')
        self.root.geometry('360x240') # 这里的乘号不是 * ，而是小写英文字母 x
        # 登陆界面
        lb_login = Label(self.root,text="本程序仅供武大科发院内部使用",relief=GROOVE)
        lb_login.place(relx=0.1,rely=0.0,relheight=0.2,relwidth=0.85)
        lb_login = Label(self.root,text="点击“登陆”后自动登入网站（需手动输入验证码）",relief=GROOVE)
        lb_login.place(relx=0.1,rely=0.2,relheight=0.2,relwidth=0.85)
        lb_login_2 = Label(self.root,text="开发者：Cainan.Guan",relief=GROOVE)
        lb_login_2.place(relx=0.1,rely=0.4,relheight=0.2,relwidth=0.85)
        # Excel路径填写
        self.excel_path = self.get_excel_path()
        path_label = Label(self.root, text='请填入Excel表格绝对路径：')
        path_label.place(relx=0.1,rely=0.6,relheight=0.1,relwidth=0.85)
        self.excel_path_entry = Entry(self.root, width=10)
        self.excel_path_entry.insert('0', self.excel_path) # 默认路径
        self.excel_path_entry.place(relx=0.1,rely=0.7,relheight=0.1,relwidth=0.8)
        # 登陆按钮
        bnt_login = Button(self.root, text='登陆', command= self.logIn) # 登陆按钮
        bnt_login.place(relx=0.30,rely=0.85,height=30,width=60)
        # 退出程序
        bnt_out = Button(self.root, text='退出', command=self.exit) # 登陆按钮
        bnt_out.place(relx=0.60,rely=0.85,height=30,width=60)
        self.root.mainloop()
    # 点击登陆响应函数
    def logIn(self):
        # 读取excel表格数据
        self.excel_path = self.excel_path_entry.get() # 获取Excel表格路径
        if self.readExcel(): # 读取excel数据成功
            # 存储本次路径
            with open('userdata.txt', 'w') as f:
                f.write(self.excel_path)
                f.close()
            # 第一次打开浏览器
            if self.driver == None:
                # 打开Chrome
                option = webdriver.ChromeOptions()
                option.add_argument('disable-infobars')
                self.driver = webdriver.Chrome(ChromeDriverManager().install(), options=option)
                self.driver.implicitly_wait(4) # 设定隐式等待最大时长
            # 登陆操作
            try:
                self.operationAuth() # 如果登陆成功
                self.root.destroy() # 销毁首页
                self.fillTableUI() # 打开自动读入操作页面
            except:
                tkinter.messagebox.askokcancel('错误','请检查网络连接！')
                pass # 登陆失败
        else:
            pass # 读入数据失败
    # 登陆授权操作
    def operationAuth(self):
        url = "http://210.12.174.1:8084" # 合同录入网站
        self.driver.get(url) # 尝试连接网页
        try:
            self.driver.find_element_by_class_name('panel-tool-close').click() # panel-tool-close
        except:
            pass
        self.driver.find_element_by_class_name('big-user').click() # 点击买方登陆
        # time.sleep(0.5)
        self.driver.find_element_by_name("clientloginname").clear()
        self.driver.find_element_by_name("clientloginname").send_keys("*****")
        self.driver.find_element_by_name("clientcodeinput").clear()
        self.driver.find_element_by_name("clientcodeinput").send_keys("******")
        # 点击取得验证码，等待人工输入
        self.driver.find_element_by_id("checkCodeImg").click()
        # 循环检查页面是否跳转
        while True:
            try:
                # 找到登陆成功页面的元素
                self.driver.find_element_by_xpath('/html/body/div[1]/div[2]/div[2]/div/div/div[2]/a[1]')
                print('登陆成功！')
                return True
            except:
                print('请及时输入验证码')
                # tkinter.messagebox.askokcancel('警告','请在网页中及时输入验证码！')
                # self.fail_flag = 2
                
    # 自动读入操作页面UI
    def fillTableUI(self):
        self.fill_table = Tk()
        self.fill_table.title('全国技术合同登记系统——自动填写小助手')
        self.fill_table.geometry('360x300')
        # 合同索引
        height_width_label = Label(self.fill_table, text='当前处理合同的索引号：')
        self.height_width_Entry = Entry(self.fill_table, width=10)
        height_width_label.grid(row=0, column=0, sticky=E)
        self.height_width_Entry.grid(row=0, column=1, sticky=W)
        # 是否启用自动搜索
        self.CheckVar1 = IntVar()
        self.use_auto_search = Checkbutton(self.fill_table, text = "使用自动搜索功能", variable = self.CheckVar1, onvalue = 1, offvalue = 0, height=5, width = 10)
        self.use_auto_search.place(relx=0.05,rely=0.4,relheight=0.2, relwidth=0.4)
        # 按键
        bnt_exit = Button(self.fill_table, text='退出程序', command=self.exit) # 退出程序
        bnt_next = Button(self.fill_table, text='开始填入', command=self.fillTable)# 填写合同
        bnt_exit.grid(row=2, column=0, sticky=E)
        bnt_next.grid(row=2, column=1, sticky=W)
        # 提示框
        sb = Scrollbar(self.fill_table, orient=tkinter.VERTICAL)
        self.txt_prompt = Text(self.fill_table, wrap=tkinter.WORD, yscrollcommand=sb.set)
        self.txt_prompt.insert(END, '使用前请阅读！')
        self.txt_prompt.insert(END, '\n（1）首先输入将要处理的合同在excel文件中的索引号（第几行），然后单击“开始填入”按钮，开始填录当前合同。')
        self.txt_prompt.insert(END, '\n（2）自动搜索功能可以在网站自动获取公司及其法人信息，该功能不稳定，可以选择关闭/开启。')
        self.txt_prompt.insert(END, '\n（3）自动填录中，不要在当前谷歌浏览器上打开额外标签页，否则易引发异常。')
        self.txt_prompt.place(relx=0.05,rely=0.6,relheight=0.5, relwidth=0.9)
        sb.config(command=self.txt_prompt.yview)
        self.fill_table.mainloop()

    # 自动填写控制接口函数
    def fillTable(self):
        print(self.height_width_Entry.get())
        self.txt_prompt.delete(0.0, END) # 清空提示框
        # 检查是否输入正确
        try:
            index = int(self.height_width_Entry.get()) # 化为整数
        except:
            self.txt_prompt.insert(END, '\n【错误】请输入一个整数！')
            return
        # 填入表格
        if index < len(self.contract_info_list) and index > 1: # 不得大于excel总行数, 不得为负数
            if self.first_flag:
                self.first_flag = False # 第一次进入不用关闭Tab
            else:
                self.closeTab() # 关掉上一个合同页面
            try:
                self.autoFillWeb(index) # after submitting, close current page
                self.txt_prompt.insert(END, '【提示1】填写完毕！本次处理的合同为：{} {}'.format(index, self.contract_info_list[index - 2] ['project_name']) )
                self.txt_prompt.insert(END, '\n【提示2】技术领域为：{}'.format(self.contract_info_list[index - 2] ['tech_area']))
                self.txt_prompt.insert(END, '\n【提示3】请仔细检查表单中的缺失数据 ！检查完毕后，请手动点击“提交”')
                index += 1
                self.height_width_Entry.delete(0, END) # 清楚现有数据
                self.height_width_Entry.insert('0', str(index)) #将下一次的标号自动插入输入框
            except:
                tkinter.messagebox.askokcancel('错误','合同填写发生错误，请检查网络连接后再试一次！')

        else:
            self.txt_prompt.insert(END, '\n【错误】输入的整数必须在：2-表格最大行数 之间！')
            return
            
    # 退出程序
    def exit(self):
        try:
            self.root.destroy() # 关闭首页
        except:
            pass
        # 检查浏览器是否关闭
        try:
            self.driver.quit()
        except:
            pass
        try:
            self.fill_table.destroy() # 关闭UI界面
        except:
            pass
        try:
            quit() # 退出程序
        except:
            sys.exit(0)

    '''
    自动填写合同的函数, index代表excel表格中的行数
    '''
    # 自动填写控制
    def autoFillWeb(self, index):
        this_contract = self.contract_info_list[index - 2]
        corp_flag = False # 买方是否为企业
        # 填入项目名称/合作对象/金额总数
        self.driver.find_element_by_xpath('/html/body/div[1]/div[2]/div[2]/div/div/div[2]/a[1]/span/span[1]').click()
        self.driver.find_element_by_name("projectname").clear()
        self.driver.find_element_by_name("projectname").send_keys(this_contract['project_name'])
        self.driver.find_element_by_name("buyername").clear()
        self.driver.find_element_by_name("buyername").send_keys(this_contract['partner_name'])
        if float(this_contract['total_amount']) <= 0: # 检查金额有效性
            self.txt_prompt.insert(END, '【警告】合同金额为0，停止录入\n')
            return
        else:
            self.driver.find_element_by_name("totalamount").clear()
            self.driver.find_element_by_name("totalamount").send_keys(this_contract['total_amount'])
            self.driver.find_element_by_name("technicalamount").clear()
            self.driver.find_element_by_name("technicalamount").send_keys(this_contract['total_amount'])
        # 填入合同签署日期
        try:
            year = int(this_contract['stamp_date'][0])
            month = int(this_contract['stamp_date'][1])
            day = int(this_contract['stamp_date'][2])
            (lines, columns) = self.getDayPos(year, month, day)
            # 合同签订日期
            self.driver.find_element_by_name("signdate").click()
            self.driver.find_element_by_class_name('calendar-title').click()
            self.driver.find_element_by_class_name('calendar-menu-year').clear()
            self.driver.find_element_by_class_name('calendar-menu-year').send_keys(year) # 填入年份
            items = self.driver.find_elements_by_class_name('calendar-menu-month')
            items[month-1].click() # month
            items = self.driver.find_elements_by_class_name('calendar-day')
            items[(lines-1)*7 + columns - 1].click() # day
            # 合同起始日期
            self.driver.find_element_by_name("contractbegindate").click()
            self.driver.find_element_by_class_name('calendar-title').click()
            self.driver.find_element_by_class_name('calendar-menu-year').clear()
            self.driver.find_element_by_class_name('calendar-menu-year').send_keys(year) # 填入年份
            items = self.driver.find_elements_by_class_name('calendar-menu-month')
            items[month-1].click() # month
            items = self.driver.find_elements_by_class_name('calendar-day')
            items[(lines-1)*7 + columns - 1].click() # day
        except:
            self.txt_prompt.insert(END, '【警告】合同签订日期错误，请手动录入！\n')
            pass

        # 合同结束日期
        self.driver.find_element_by_name("contractenddate").click()
        self.driver.find_element_by_class_name('calendar-title').click()
        self.driver.find_element_by_class_name('calendar-menu-year').clear()
        self.driver.find_element_by_class_name('calendar-menu-year').send_keys('2021') # 填入年份(默认2021.12.31)
        items = self.driver.find_elements_by_class_name('calendar-menu-month')
        items[11].click() # month
        items = self.driver.find_elements_by_class_name('calendar-day')
        items[(5 - 1)*7 + 6 - 1].click() # day

        # 支付方式选择 （根据金额选择）
        self.driver.find_element_by_name('paymethod_view').click()
        items = self.driver.find_elements_by_class_name('combobox-item')
        if float(this_contract['total_amount']) >= 150000.0:
            items[2].click() # 分期支付
        else:
            items[1].click() # 一次支付

        # 关联交易（否） / 项目计划来源（计划外）
        self.driver.find_element_by_name('isrelated_view').click()
        items = self.driver.find_elements_by_class_name('combobox-item')
        items[1].click() # 否
        self.driver.find_element_by_name('projectplantype_view').click()
        items = self.driver.find_elements_by_class_name('tree-node')
        items[87].click() # 最后一项，计划外

        # 知识产权 （默认未涉及）
        self.driver.find_element_by_name('ipr_view').click()
        items = self.driver.find_elements_by_class_name('tree-node')
        items[12].click() # 未涉及知识产权

        # 合同类型 （根据表格选择）
        self.driver.find_element_by_name('contracttype_view').click()
        items = self.driver.find_elements_by_class_name('tree-node')
        if this_contract['contract_type'] == '技术开发' or this_contract['contract_type'] == '开发':
            items[0].click()
            items[2].click() # 委托开发
        elif this_contract['contract_type'] == '技术服务' or this_contract['contract_type'] == '服务':
            items[13].click()
            items[15].click() # 技术服务
        else:
            items[17].click() # 其他

        # 是否技术转移机构（否）
        self.driver.find_element_by_name('buyer_jszyjg_view').click()
        items = self.driver.find_elements_by_class_name('tree-node')
        items[1].click() # 否

        # 是否研发机构（是）
        self.driver.find_element_by_name('buyer_isyfjg_view').click()
        items = self.driver.find_elements_by_class_name('combobox-item')
        items[0].click() # 是

        # 买方性质 （根据表格选择）
        self.driver.find_element_by_name('buyertype_view').click()
        items = self.driver.find_elements_by_class_name('tree-node')
        if this_contract['partner_type'] == '企业':
            corp_flag = True
            items[3].click()
            time.sleep(0.5) # 强制等待页面变换
            items = self.driver.find_elements_by_class_name('tree-node')
            items[4].click()
            time.sleep(0.5)
            items = self.driver.find_elements_by_class_name('tree-node')
            items[9].click()
        elif this_contract['partner_type'] == '高校':
            items[1].click()
            time.sleep(0.5)
            items = self.driver.find_elements_by_class_name('tree-node')
            items[3].click()
        elif this_contract['partner_type'] == '其他事业单位' or this_contract['partner_type'] == '事业单位':
            items[1].click()
            time.sleep(0.5)
            items = self.driver.find_elements_by_class_name('tree-node')
            items[5].click()
        elif this_contract['partner_type'] == '政府部门':
            items[0].click()
            time.sleep(0.5)
            items = self.driver.find_elements_by_class_name('tree-node')
            items[4].click()
        else: # 其他
            items[5].click()
            time.sleep(0.5)
            items = self.driver.find_elements_by_class_name('tree-node')
            items[7].click()

        # 是否转制科研院所（否）
        self.driver.find_element_by_name('buyer_zzkyys_view').click()
        items = self.driver.find_elements_by_class_name('combobox-item')
        try:
            items[1].click() # 否
        except:
            pass

        # 如果是企业则额外填补三项
        if corp_flag:
            # 是否高新区内企业（否）
            self.driver.find_element_by_name('buyer_countrygxqqy_view').click()
            items = self.driver.find_elements_by_class_name('combobox-item')
            try:
                items[1].click() # 否
            except:
                pass
            # 上市公司 （否）
            self.driver.find_element_by_name('buyer_islist_view').click()
            items = self.driver.find_elements_by_class_name('combobox-item')
            try:
                items[1].click() # 否
            except:
                pass
            # 企业规模 （无标准）
            self.driver.find_element_by_name('buyer_scale_view').click()
            items = self.driver.find_elements_by_class_name('combobox-item')
            try:
                items[1].click() # 否
            except:
                pass
        # 机构从事的国民经济行业 （需手动选择）
    
        # 技术领域 -> 未使用
        try:
            flag_tech = 0
            tech_num = self.dic_tech_area[this_contract['tech_area']]
        except:
            flag_tech = 1
            self.txt_prompt.insert(END, "【警告】合同中的技术领域无效：{} \n".format(this_contract['tech_area']))
        if 0: # 技术领域有效
            self.driver.find_element_by_name('techarea_view').click()
            time.sleep(0.5) # 强制等待页面变换
            head_node = self.driver.find_element_by_class_name('tree') # 0 - 10
            items = head_node.find_elements_by_tag_name('li') # 第一层所有对象
            if tech_num == 0: # 城市建设
                items[10].click()
                time.sleep(0.5) # 强制等待页面变换
                items[10].find_element_by_xpath("./ul/li[4]/div").click()
            elif tech_num == 1: # 电子信息
                items[0].click()
                time.sleep(0.5) # 强制等待页面变换
                items[0].find_element_by_xpath("./ul/li[6]/div").click()
            elif tech_num == 2: # 航空航天
                self.txt_prompt.insert(END, "【警告】合同中的技术领域需要手动选择：{} \n".format(this_contract['tech_area']))
            elif tech_num == 3: # 核应用
                self.txt_prompt.insert(END, "【警告】合同中的技术领域需要手动选择：{} \n".format(this_contract['tech_area']))
            elif tech_num == 4: # 资源环境综合利用
                items[6].click()
                time.sleep(0.5) # 强制等待页面变换
                items[6].find_element_by_xpath("./ul/li[5]/div").click()
            elif tech_num == 5: # 农业
                items[8].click()
                time.sleep(0.5) # 强制等待页面变换
                items[8].find_element_by_xpath("./ul/li[3]/div").click()
            elif tech_num == 6: # 生物、医药和医疗器械
                items[3].click()
                time.sleep(0.5) # 强制等待页面变换
                items[3].find_element_by_xpath("./ul/li[6]/div").click() 
            elif tech_num == 7: # 先进制造
                items[2].click()
                time.sleep(0.5) # 强制等待页面变换
                items[2].find_element_by_xpath("./ul/li[3]/div").click()
            elif tech_num == 8:
                self.txt_prompt.insert(END, "【警告】合同中的技术领域需要手动选择：{} \n".format(this_contract['tech_area']))
            elif tech_num == 9:
                items[4].click()
                time.sleep(0.5) # 强制等待页面变换
                items[4].find_element_by_xpath("./ul/li[9]/div").click()
            elif tech_num == 10:
                items[5].click()
                time.sleep(0.5) # 强制等待页面变换
                items[5].find_element_by_xpath("./ul/li[6]/div").click()
            elif tech_num == 11:
                items[0].click()
                time.sleep(0.5) # 强制等待页面变换
                items[0].find_element_by_xpath("./ul/li[6]/div").click()
            elif tech_num == 12:
                items[5].click()
                time.sleep(0.5) # 强制等待页面变换
                items[5].find_element_by_xpath("./ul/li[7]/div").click()
            elif tech_num == 13:
                items[5].click()
                time.sleep(0.5) # 强制等待页面变换
                items[5].find_element_by_xpath("./ul/li[7]/div").click()
            elif tech_num == 14:
                self.txt_prompt.insert(END, "【警告】合同中的技术领域需要手动选择：{} \n".format(this_contract['tech_area']))
            else:
                self.txt_prompt.insert(END, "【警告】合同中的技术领域需要手动选择：{} \n".format(this_contract['tech_area']))
                    
        # 启用自动搜索
        if self.CheckVar1.get() == 1:
        # 得到公司具体信息
            try:
                (position, lp_name, lp_number) = self.getPartnerInfo(this_contract['partner_name'])
                self.driver.find_element_by_name('buyer_address').send_keys(position)
                self.driver.find_element_by_name('buyer_representative').send_keys(lp_name)
                self.driver.find_element_by_name('buyer_contact').send_keys(lp_name)
                self.driver.find_element_by_name('buyer_contacttel').send_keys(lp_number)
            except:
                windows = self.driver.window_handles
                self.driver.switch_to.window(windows[0]) # 切换初始页面
                self.txt_prompt.insert(END, "【警告】无法从互联网获得买家信息，请手动搜索！\n")
            # 得到邮编
            try:
                post_num = self.findPostNum(position)
                self.driver.find_element_by_name('buyer_zipcode').send_keys(post_num)
            except:
                self.txt_prompt.insert(END, "【警告】无法从互联网获得邮编，请手动搜索！\n")
                self.driver.find_element_by_name('buyer_zipcode').send_keys('000000')
    # 关闭当前合同信息
    def closeTab(self):
        try:
            self.driver.find_element_by_class_name('tabs-close').click()
        except:
            pass

    '''
    以下为工具函数
    '''
    # 获取存储的文件路径
    def get_excel_path(self):
        if os.path.exists('userdata.txt'):   
            with open('userdata.txt','r') as f:
                add = f.readline().rstrip()
                f.close()
                return add
        # 默认文件路径
        return '/Users/guanqianyun/Desktop/2021年-横向（进账）合同登记表.xls'
    # 读入Excel表格数据
    def readExcel(self):
        # 读入合同信息
        try:
            profile_data = xlrd.open_workbook(self.excel_path)
        except: # 找不到文件
            tkinter.messagebox.askokcancel('错误','请检查输入的Excel表格路径！')
            return False # 数据读入失败
        table1 = profile_data.sheets()[0] # 通过索引顺序获取table1
        # 写入合同信息
        for i in range(1, table1.nrows): # 跳过第一行
            row = table1.row_values(i)
            try: # 检查金额数据有效性
                amount = float(row[8])
            except:
                row[8] = -1
            # 写入数据
            dic = self.writeRow(row)
            self.contract_info_list.append(dic)
        # inspect result
        return True # 数据读入成功
    # 爬取邮编
    def findPostNum(self, position):
        flag1 = False
        flag2 = False
        headers = {
            "User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/68.0.3440.106 Safari/537.36"
        }
        # 选取地址的前十个字符
        keyword = position+ '%20邮编'
        print('keyword:', keyword)
        url = 'https://www.baidu.com/s?wd=' + keyword # 在百度中查找关键字
        response = requests.get(url, headers=headers, timeout=3) # 3秒超时
        response.encoding = 'utf-8'
        soup2 = BeautifulSoup(response.text, "html.parser")
        p_content2 = soup2.find_all('em') # 寻找标红字段
        for i in range(len(p_content2)):
            if p_content2[i].text == '邮编' or p_content2[i].text == '邮政编码':
                critical_content = p_content2[i].parent.text
                flag1 = True
                break

        # 寻找邮编
        if flag1: # 找到标红内容
            cnt = 0
            post_num = ''
            for i in range(len(critical_content)):
                if critical_content[i].isdigit():
                    cnt = cnt + 1
                    post_num += critical_content[i]
                    if cnt >= 6:
                        flag2 = True
                        break # 出现6次数字代表找到了邮编
                else:
                    cnt = 0
                    post_num = ''
            if not flag2: # 未能找到6位数字
                self.txt_prompt.insert(END, "【警告】无法从互联网获得邮编，请手动搜索！\n")
                post_num = '000000'
        else: # 未能找到标红内容
            self.txt_prompt.insert(END, "【警告】无法从互联网获得邮编，请手动搜索！\n")
            post_num = '000000'
        return post_num
    # 获取日历坐标
    def getDayPos(self, year, month, day):
        # 获取列标
        someday=datetime.date(year,month,day)
        columns = int(someday.strftime('%w')) % 7 + 1
        # 获取行标
        first_day = int(datetime.date(year, month, 1).strftime('%w')) % 7 + 1
        lines = int( (day + first_day - 2) / 7 )+ 1
        return (lines, columns)
    # 在企查查网站找到公司具体信息
    def getPartnerInfo(self, partner_name):
        js = 'window.open("https://aiqicha.baidu.com/");'
        self.driver.execute_script(js) # 打开新的tab
        windows = self.driver.window_handles
        self.driver.switch_to.window(windows[-1]) # 窗口跳转
        # 搜索公司
        #try:
        #    self.driver.find_element_by_class_name("close").click()
        #except:
        #    pass
        self.driver.find_element_by_id("aqc-search-input").send_keys(partner_name)
        self.driver.find_element_by_class_name("search-btn").click()
        # 获取首个搜索结果
        self.driver.find_elements_by_class_name("card")[0].find_element_by_tag_name("a").click()
        # 获取电话
        windows = self.driver.window_handles
        self.driver.switch_to.window(windows[-1]) # 窗口跳转
        self.driver.find_element_by_tag_name("table") # 确保页面跳转
        try:
            # 公司电话
            tel = self.driver.find_elements_by_class_name('content-info-child')[0].find_elements_by_tag_name('span')[1].text
        except:
            tel = '00000000000'
        try:
            # 公司地址
            addr = self.driver.find_elements_by_class_name('content-info-child')[1].find_element_by_class_name('content-info-child-right').find_element_by_class_name('child-data').text
        except:
            addr = '*'
        try:
            # 法人姓名
            legal_p = self.driver.find_element_by_class_name('portrait-text').find_elements_by_tag_name('a')[0].text
        except:
            legal_p = '*'
        # 关闭所有新窗口
        windows = self.driver.window_handles
        for i in range(len(windows) - 1):
            self.driver.switch_to.window(windows[i + 1])
            self.driver.close()
        # 返回原窗口
        windows = self.driver.window_handles
        self.driver.switch_to.window(windows[0])
        return (addr, legal_p, tel) # 返回（具体位置， 法人名字， 联系方式）
    # 写入dic
    def writeRow(self, row):
        dic = {
            'id': row[0],
            'department': row[1],
            'leading_people': row[2],
            'phone_num': row[3],
            'project_name': row[4],
            'partner_name': row[5],
            'partner_district': row[6],
            'partner_type': row[7],
            'total_amount': str(row[8] * 10000),
            'stamp_date': str(row[9]).split("."), # 分隔为年/月/日
            'effect_date': str(row[10]).split("."),
            'tech_area': row[11],
            'contract_type': row[12],
            'contract_reg_people': row[13],
        }
        return dic

if __name__ == '__main__':
    auto_util = auto_op() # 实例化
    auto_util.firstPage() # 进入程序首页
