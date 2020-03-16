import os
import smtplib
import sys
from email.mime.text import MIMEText
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import random
import schedule
from time import sleep
import time

# 注：需将个人教务系统账号密码以及相关邮箱信息分别存放在同目录下的myInfo.txt和emailInfo.txt文件中
with open('myInfo.txt', 'r') as f:
    # 个人账号信息封装在myInfo.txt文件中，其中依次存放账号、密码信息（用空格分开）
    myInfo = f.read()
    myInfo = myInfo.split()
    usr_name = myInfo[0]
    usr_password = myInfo[1]


# 打卡
def checkIn():
    driver = webdriver.Chrome()
    driver.get("http://ehall.jit.edu.cn/new/index.html")
    print("加载“我的金科院”页面成功！")
    # 打开Chrome浏览器并进入我的金科院首页  注：需下载webdriver（Chrom版）放在Chrom根目录下，MacOS放在Python根目录下

    now_handle = driver.current_window_handle
    driver.switch_to.window(now_handle)
    try:
        WebDriverWait(driver, 1800).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".amp-no-login-zh")))
    finally:
        sleep(1)
    driver.find_element_by_css_selector(".amp-no-login-zh").click()
    # 等待首页的登录按钮加载完成后点击登录按钮

    now_handle = driver.current_window_handle
    print("加载登录页面成功!")
    driver.switch_to.window(now_handle)

    driver.find_element_by_id("username").clear()
    driver.find_element_by_id("password").clear()
    driver.find_element_by_id("username").send_keys(usr_name)
    sleep(1)
    driver.find_element_by_id("password").send_keys(usr_password)
    sleep(1)
    driver.find_element_by_css_selector(".ipt_btn_dl").click()
    # 进入登录界面后填写用户名和密码并点击登录

    now_handle = driver.current_window_handle
    print("加载“学生桌面”页面成功!")
    sleep(1)
    driver.switch_to.window(now_handle)
    try:
        WebDriverWait(driver, 1800).until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, ".card-myFavorite-content-item > .style-scope:nth-child(4) .widget-title")))
    finally:
        sleep(1)

    driver.find_element_by_css_selector(
        ".card-myFavorite-content-item > .style-scope:nth-child(4) .widget-title").click()
    # 成功进入学生桌面后等待“健康信息填报系统”按钮加载完成，加载完成后点击它进入打卡页面

    windos = driver.window_handles
    driver.switch_to.window(windos[-1])
    # 切换到新打开的打卡页面窗口
    print("加载“信息填报”页面成功!")
    sleep(1)
    now_handle = driver.current_window_handle
    try:
        WebDriverWait(driver, 1800).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".bh-mb-16 > .bh-btn-primary")))
    finally:
        sleep(1)
    driver.find_element_by_xpath("//div[@data-action='add']").click()
    # 等待“新增”按钮加载完成，加载完成后点击“新增”按钮
    print("点击“新增”按钮")

    now_handle = driver.current_window_handle
    driver.switch_to.window(now_handle)

    try:
        WebDriverWait(driver, 1800).until(
            EC.presence_of_element_located((By.NAME, "DZ_DZBZ")))
    finally:
        sleep(1)

    js = "var q=document.documentElement.scrollTop=100"
    driver.execute_script(js)
    # 将页面滚动
    sleep(3)

    driver.find_element_by_id("save").click()
    print("点击“保存”按钮")
    # 等待“保存”按钮加载完成，加载完成后点击“保存”按钮

    now_handle = driver.current_window_handle
    driver.switch_to.window(now_handle)

    try:
        WebDriverWait(driver, 1800).until(
            EC.presence_of_element_located((By.XPATH, "//label[text()='否']")))
    finally:
        sleep(1)
    driver.find_element_by_xpath("//label[text()='否']").click()
    # 等待“是否”对话框加载完成，加载完成后点击“否”单选框
    print("选择“否”")
    sleep(10)
    # 等待10秒
    driver.find_element_by_xpath("//button[contains(@class,'bh-btn bh-btn-primary')]").click()
    # 点击提交后打卡成功
    print("打卡成功！")
    sleep(3)
    driver.quit()
    sentEmail()
    # 发送邮件通知打卡成功
    sleep(60)
    shutdown()
    # 关机


# 发邮件
def sentEmail():
    with open('emailInfo.txt', 'r') as f1:
        # 电子邮箱信息封装在emailInfo.txt文件中，其中依次存放SMTP服务器地址、发送者邮箱账号、发送者QQ邮箱授权码、接收者邮箱账号信息（用空格分开）
        # 注：需进QQ邮箱设置中开启POP3/SMTP服务（开户后会显示授权码，务必记下，登录邮箱时要用到）、IMAP/SMTP服务，
        # 并勾选“收取我的文件夹”、“SMTP发信后保存到服务器”，也可使用其他邮箱，使用方法见百度
        emailInfo = f1.read()
        emailInfo = emailInfo.split()
        host = emailInfo[0]
        sender = emailInfo[1]
        qqCode = emailInfo[2]
        receiver = emailInfo[3]
    port = 465
    body = '<h1>你已成功打卡</h1>'
    msg = MIMEText(body, 'html')
    msg['subject'] = '打卡通知'
    msg['from'] = sender
    msg['to'] = receiver

    s = smtplib.SMTP_SSL(host, port)
    s.login(sender, qqCode)
    # sender就是要登录的邮箱的账号，qqCode就是QQ邮箱的授权码
    s.sendmail(sender, receiver, msg.as_string())
    print("成功发送邮件！")
    # 发送邮件的函数，打卡完成后发送一封email到接收者邮箱


# 关机
def shutdown():
    os.system("shutdown -s -t  60 ")
    print("一分钟后关机！！！")


if __name__ == "__main__":
    # checkIn() # 测试用，立即执行checkIn函数
    schedule.every().day.at("00:00").do(checkIn)
    # 晚上12点时执行打卡的函数，成功后会发送邮件提示并自动关机

    while True:
        schedule.run_pending()
        time.sleep(1)
