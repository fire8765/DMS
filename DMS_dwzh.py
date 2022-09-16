#!/usr/bin/env python
# coding: utf-8

# In[1]:


import requests
import json
import re
import hashlib
import js2py
import time
import os
import pandas as pd
import glob
import xlrd
import shutil
import xlsxwriter
from ZZchaojiying import Chaojiying_Client
from PIL import Image
class Test(object):
    session = None

    def __init__(self,username,password1):
        self.username = username
        self.password1 = password1
#         self.password1 = password1
    
    def hex_pw(self):
        inputname = hashlib.md5()
        inputname.update(self.password1.encode("utf-8"))
        self.password=inputname.hexdigest()
        return self.password
    def yzm(self):
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64; Trident/7.0; rv:10.0) like Gecko',
            'Referer': 'http://dms.changan.com.cn/jc/common/UserManager/logout.do',
            'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',


        }
        js_func = """
                        now = new Date(); 
                        src = "http://dms.changan.com.cn/jc/image.jsp?code="+now.getTime(); 

              """

        codeurl = js2py.eval_js(js_func)        
    #         yzm="http://dms.changan.com.cn/jc/image.jsp"
        res=session.get(codeurl,headers=headers)
        res2=session.post(url='http://dms.changan.com.cn/jc/common/ValidateCodeAction/init.json',headers=headers)        
#         yzm="http://dms.changan.com.cn/jc/image.jsp"
        res=self.session.get(codeurl)
        with open('1.png','wb')as f:
            f.write(res.content)
        chaojiying = Chaojiying_Client('hongfa123', 'hongfa888', '6001')	#用户中心>>软件ID 生成一个替换 96001
        im = open('1.png', 'rb').read()		#本地图片文件路径 来替换 a.jpg 有时WIN系统须要//
        imgstr=chaojiying.PostPic(im, 6001).get("pic_str")
        return imgstr
    def yzminput(self):
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64; Trident/7.0; rv:10.0) like Gecko',
            'Referer': 'http://dms.changan.com.cn/jc/common/UserManager/logout.do',
            'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',


        }
        js_func = """
                        now = new Date(); 
                        src = "http://dms.changan.com.cn/jc/image.jsp?code="+now.getTime(); 

              """

        codeurl = js2py.eval_js(js_func)        
    #         yzm="http://dms.changan.com.cn/jc/image.jsp"
        res=session.get(codeurl,headers=headers)
        res2=session.post(url='http://dms.changan.com.cn/jc/common/ValidateCodeAction/init.json',headers=headers)  
#         res=requests.get(codeurl)

        with open('1.png','wb')as f:
            f.write(res.content)
        im = Image.open('1.png') 
        im.show()
        imgstr=input('请输入计算结果：')
        return imgstr
    def login(self,dian):
        #登录1
        self.session=requests.Session()
        login1url='http://dms.changan.com.cn/jc/index.jsp'
        lsp=self.session.get(login1url)
        login_url="http://dms.changan.com.cn/jc/common/UserManager/login.do"
    #     yzm="http://dms.changan.com.cn/jc/image.jsp"
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64; Trident/7.0; rv:10.0) like Gecko',
            'Referer': 'http://dms.changan.com.cn/jc/index.jsp',
            'Content-Type': 'application/x-www-form-urlencoded'
        }
        try:
            n=0
            while n < 5:
                data={

                   'password': self.hex_pw(),
                   'password1': self.password1,
                   'userName': self.username,
                   'validateCode': self.yzminput()
                }
                res=self.session.post(url=login_url,headers=headers,data=data)
                n+=1
                print('正在进行第'+str(n)+'次打码登录')
                logintext=res.text
    #             print(res.text)
    #             print(type(logintext))
    #             try:
                if '重复使用' in logintext:
                    print('需重新打码')
                    continue
                elif '验证码错误' in logintext:
                    print('需重新打码')
                    continue 
                elif 'document.createElement' in logintext:
                    print('打码登录成功')
                    break

                else:
                    continue
        except Exception as e:
            print(dian+'登录出错了：', e) 
        try:
#             req = html.cssselect('.TableList .normal_btn')[1]    
            resp = re.findall(r'职位选择.*',res.text)
        #     HtmlStr=etree.tostring(resp,encoding="utf-8").decode() 
        #     resp = HtmlStr.xpath("//input/@onclick")
#             print(resp)
            posdata=re.findall(r'\d+',resp[3])
#             print(posdata)
           
        except:
            da1=re.findall(r'\tgoTo.*',res.text)
            posdata=re.findall(r'\d+',str(da1))
#             print(posdata)
        try:
            data1={}
            data1['deptId']=posdata[0]
            data1['poseBusType']=posdata[3]
            data1['poseId']=posdata[1]
            data1['poseType']=posdata[2]
    #         print(data1)

            zyurl="http://dms.changan.com.cn/jc/common/MenuShow/menuDisplay.do"
            res2=self.session.post(url=zyurl,data=data1)
            time.sleep(1)
            res3=self.session.get("http://dms.changan.com.cn/jc/menu/leftMenu.jsp")
            print(dian+'登录成功')
#             print(res3.text)
        except Exception as e:
            print(dian+'登录失败：', e)    
            


    def load_cookies(self,diandata):
        with open("./COOKIE"+diandata+"cookie.txt", "r") as f:
            self.load_cookies=json.loads(f.read())
        return self.load_cookies
    def zhihuan_down(self,startDate, endDate, diandata,skcode,customer_name):
        zh_url = "http://dms.changan.com.cn/jc/sales/customerInfoManage/SalesReportExchangeCheck/queryReportInfo_Check2.json?COMMAND=1"  # 实销查询网址
#         zh_url='http://dms.changan.com.cn/jcx/sales/customerInfoManage/SalesReportBuyMoreCheck/queryBuyMoreList_Check2.json?COMMAND=1'
        headers1 = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64; Trident/7.0; rv:10.0) like Gecko',
            'Referer': 'http://dms.changan.com.cn/jc/sales/customerInfoManage/SalesReportExchangeCheck/queryInit2.do?g_webAppName=/jcx&isReport=null',
            'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
            # 'Host': 'dms.changan.com.cn',
            # 'Content-Length': '200',
            # 'Accept': 'text/javascript, text/html, application/xml, text/xml, */*',
            # 'Accept-Encoding': 'gzip, deflate',
            # 'Accept-Language': 'zh-Hans-CN, zh-Hans;q=0.5',
            # 'Cache-Control': 'no-cache',
            # 'Connection': 'Keep-Alive',
            # 'X-Prototype-Version': '1.6.0',
            # 'X-Requested-With': 'XMLHttpRequest'
        }
        data1 = {
            # 'curPage': '1',
            'status': skcode,#审核状态skcode='17061002' 001为待上传，004为驳回，003为不通过
            'startDate': startDate,
            'endDate': endDate,
            'customer_name': customer_name,
            'JCflag': 'JC'

        }
        req = self.session.post(url=zh_url, headers=headers1, data=data1)
#         print(self.session.cookies)
#         print(req.text)
        try:
#         print(req.text)
            d = json.loads(req.text)
            tp = d['ps']['totalPages']
            print(diandata +'获取到置换明细总页数为' + str(tp))
            print(diandata +'获取置换总数为：'+str(d['ps']['totalRecords']))
            for page in range(1, tp + 1):
                print('开始下载第'+str(page)+'页...')
                with open(r'D:\pyoutdata\待上传置换增购\json\\'+diandata + '置换明细' + str(page) + '.json', "wb")as f:
                    f.write(self.session.post(url=zh_url + "&curPage=" + str(page), headers=headers1, data=data1).content)
            print('置换数据下载成功！')
        except Exception as e:
            print('出错了：', e)


    def quit(self):
        self.session.get('http://dms.changan.com.cn/jc/common/UserManager/logout.do')
        self.session.close()
        print("本次下载结束！")
        
    def forcsv(self):
        print('开始转换数据！')
        lst = glob.glob(r"D:\pyoutdata\待上传置换增购\json\*.json")
        # print (lst)
        for i in lst:
            with open(i, 'r', encoding='utf-8', errors='ignore') as f:
                rows = json.load(f)
            # print(rows['ps']['records'][0])
            # 将json中的key作为header, 也可以自定义header（列名）
            header=tuple([ i for i in rows['ps']['records'][0].keys()])
            # print(header)
            data = []

            # 循环里面的字典，将value作为数据写入进去
            for row in rows['ps']['records']:
                body = []
                for v in row.values():
                    body.append(v)
                data.append((body))
            # 将含标题和内容的数据放到data里
            # print(data)
            newpd = pd.DataFrame(data)
            writer = pd.ExcelWriter(i + '.xlsx')
            newpd.to_excel(excel_writer=writer, sheet_name='sheet1', index=False, header=header)
            writer.save()
            # writer.close()
        lst1 = glob.glob(r"D:\pyoutdata\待上传置换增购\json\*.xlsx")
        # print(lst)
        for j in lst1:
            # print(j)
            # print(j[25:])
            # print(j[0:20])
            df = pd.read_excel(j)
            # print(df)
            # df1 = df.drop([3]) #删除第0行，inplace=True则原数据发生改变
            # df1 = df.drop(['SALES_EXCHANGE_ID'],axis=1) #删除列
            df1 = df.drop(df.columns[[6, 7, 9, 10, 11]], axis=1)
            # print(df1)
            df1.to_excel(j[0:20]+'\\cleandata\\'+ j[25:]+'newdata.xlsx', header=None, index=False)
        print('数据转换成功')
    def hebing(self,dian):
        file = glob.glob(r"D:\pyoutdata\待上传置换增购\cleandata\*.xlsx")
        target_xls = r"D:\pyoutdata\待上传置换增购\最终明细表\置换\最终"+dian+"合并明细表.xlsx"
        # 读取数据
        data = []
        for f in file:
            wb = xlrd.open_workbook(f)
            for sheet in wb.sheets():
                for rownum in range(sheet.nrows):
                    data.append(sheet.row_values(rownum))
        # print(data)
        # 写入数据
        workbook = xlsxwriter.Workbook(target_xls)
        worksheet = workbook.add_worksheet()
        font = workbook.add_format({"font_size": 10})
        for i in range(len(data)):
            for j in range(len(data[i])):
                worksheet.write(i, j, data[i][j], font)
        # 关闭文件流
        workbook.close()
        print('数据合并成功！')

        #Python简单删除目录下文件以及文件夹
        filelist1=[]
        filelist2=[]
        rootdir1=r"D:\pyoutdata\待上传置换增购\json"
        rootdir2=r"D:\pyoutdata\待上传置换增购\cleandata" #选取删除文件夹的路径,最终结果删除img文件夹
        filelist1=os.listdir(rootdir1)
        filelist2=os.listdir(rootdir2)
        for k in filelist1:
            filepath = os.path.join(rootdir1,k)#将文件名映射成绝对路劲
            if os.path.isfile(filepath):            #判断该文件是否为文件或者文件夹
                os.remove(filepath)                 #若为文件，则直接删除
                print(str(filepath)+" removed!")
            elif os.path.isdir(filepath):
                shutil.rmtree(filepath,True)        #若为文件夹，则删除该文件夹及文件夹内所有文件
                print("dir "+str(filepath)+" removed!")    
        for s in filelist2:
            filepath = os.path.join(rootdir2,s)#将文件名映射成绝对路劲
            if os.path.isfile(filepath):            #判断该文件是否为文件或者文件夹
                os.remove(filepath)                 #若为文件，则直接删除
                print(str(filepath)+" removed!")
            elif os.path.isdir(filepath):
                shutil.rmtree(filepath,True)        #若为文件夹，则删除该文件夹及文件夹内所有文件
                print("dir "+str(filepath)+" removed!")    
        print("缓存删除成功")
def test1():
    data = "信息表4.xlsx"
    name1 = xlrd.open_workbook(data)
    sheet = name1.sheet_by_name("Sheet1")#取帐户信息
    sheet2 = name1.sheet_by_name("Sheet2")#取帐户信息
    ddd=1  # 此处改用户名-0宏发，1宏卓，2汇源，3宏诚，4宏哲，5宏之睿，6宏大
    n = sheet.col_values(0)[ddd] # 此处改用户名-0宏发，1汇源，2宏大，3宏诚，4宏哲，5宏之睿，6宏卓
    p = sheet.col_values(1)[ddd]  # 此处改密码
    dian=sheet.col_values(3)[ddd]#店名代码
    d=sheet.col_values(4)[ddd]#店名代码
    startDate='2021-12-01' # 开始日期,注意格式
    endDate='2022-02-28'#结束日期,注意格式
    skcode='17061001' # 001待上传，002审核通过，004驳回，003不通过
    customer_name='付永群'
    customer_names=sheet2.col_values(0)
    tmp = Test(n,p)
    tmp.hex_pw()
    tmp.login(dian)
    time.sleep(3)
    for name in customer_names:
        tmp.zhihuan_down(startDate,endDate,dian,skcode,name)#置换明细
        tmp.forcsv()
        tmp.hebing(name)
#         time.sleep(2)
#         tmp.shixiao_dw(startDate,endDate,dian,str(d))#实销明细下载
#         time.sleep(1)
    tmp.quit()
    print('All down!')
def test():
    data = "信息表4.xlsx"
    name1 = xlrd.open_workbook(data)
    sheet = name1.sheet_by_name("Sheet1")#取帐户信息
    username = sheet.col_values(0) # 此处改用户名-0宏发，1汇源，2宏大，4宏诚，5宏哲，11宏之睿，12宏卓
    password1 = sheet.col_values(1)  # 此处改密码
    diandata=sheet.col_values(3)#店名代码
    dc=sheet.col_values(4)#店名代码
    startDate='2022-01-01' # 开始日期,注意格式
    endDate='2022-03-31'#结束日期,注意格式
    skcode='17061001' # 001待上传，002审核通过，004驳回，003不通过
    customer_name=''
    for n,p,dian,d in zip(username,password1,diandata,dc):
        tmp = Test(n,p)
        tmp.hex_pw()
        tmp.login(dian)
#         time.sleep(3)
#         tmp.dakehu_dw(startDate,endDate,dian,str(d))#大客户
#         time.sleep(2)
#         tmp.zherang_dw(s,e,dian,str(d))#折让明细
        time.sleep(3)
        tmp.zhihuan_down(startDate,endDate,dian,skcode,customer_name)#置换明细
        tmp.forcsv()
        tmp.hebing(dian)
#         time.sleep(2)
#         tmp.shixiao_dw(s,e,dian,str(d))#实销明细下载
#         time.sleep(1)
        tmp.quit()
    print('所有下载完毕')

if __name__ == "__main__":
    test()#批量
#     test1()#单店


# In[ ]:





# In[ ]:




