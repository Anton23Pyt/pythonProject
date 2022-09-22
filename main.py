# -*- coding: cp1251 -*-
from selenium import webdriver
from webdriver_manager.firefox import GeckoDriverManager
import time as ti
import openpyxl as op
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.service import Service


def delete1(driver):
    d = driver.find_element_by_css_selector(".Button div.ripple-container").click()
    ti.sleep(1)
    d = driver.find_element_by_css_selector(".top div:nth-of-type(3)").click()
    ti.sleep(4)
    x = 1
    try:
        while x < 2:
            d5 = driver.find_element_by_css_selector(
                ".chat-list > div.ListItem:nth-of-type(1) div.ripple-container").click()
            ti.sleep(1)
            d = driver.find_element_by_css_selector(".chat-info-wrapper h3").click()
            ti.sleep(1)
            d = driver.find_element_by_css_selector("i.icon-edit").click()
            ti.sleep(1)
            d = driver.find_element_by_css_selector(".destructive div.ripple-container").click()
            ti.sleep(1)
            d = driver.find_element_by_css_selector("button.danger").click()
            ti.sleep(4)
    except:
        try:
            d5 = driver.find_element_by_css_selector(".chat-list > div.ListItem div.ripple-container").click()
            ti.sleep(1)
            d = driver.find_element_by_css_selector(".chat-info-wrapper h3").click()
            ti.sleep(1)
            d = driver.find_element_by_css_selector("i.icon-edit").click()
            ti.sleep(1)
            d = driver.find_element_by_css_selector(".destructive div.ripple-container").click()
            ti.sleep(1)
            d = driver.find_element_by_css_selector("button.danger").click()
        except:
            o = 0



e=input("название файла:")
q=input("Введите ячейку(с большой буквы латиницей):")
e=str(e)+".xlsx"
wb=op.reader.excel.load_workbook(filename=e)

wb.active=0
sheet=wb.active
r=[]
s1=sheet.max_row
for i in range(1,s1+1):
    d=str(q)+str(i)
    if str(sheet[d].value)!="f" and str(sheet[d].value)!="s" and str(sheet[d].value)!=None:
        r.append("+" + str(sheet[d].value))
driver = webdriver.Firefox(service=Service(GeckoDriverManager().install()))
driver.get("https://web.telegram.org/z/")
f2=input()
d = driver.find_element_by_css_selector(".Button div.ripple-container").click()
ti.sleep(1)
d = driver.find_element_by_css_selector(".top div:nth-of-type(3)").click()
ti.sleep(1)
d=driver.find_element_by_css_selector("input.form-control").click()
d=driver.find_element_by_css_selector("input.form-control")
d.send_keys("k")
ti.sleep(10)
x = 1
i=1
fi=[]
df=[]
try:
    while x < 2:
        e='.chat-list > div.ListItem:nth-of-type('+str(i)+') div.ripple-container'
        d5 = driver.find_element_by_css_selector(e).click()
        ti.sleep(1)
        d = driver.find_element_by_css_selector(".chat-info-wrapper h3").click()
        sd = driver.find_element_by_css_selector(".chat-info-wrapper h3")
        sd=sd.text
        ti.sleep(1)
        d = driver.find_element_by_class_name("multiline-item")
        z1 = d.text
        z1 = z1.replace("Phone", "")
        z1 = z1.replace(" ", "")
        z1 = z1.replace("\n", "")
        po=0
        for j in [r]:
            if z1 in j:
                 i = i + 1
                 try:
                     df.append(int(sd.replace("k", "")))
                 except:
                     gpe=0
                 fi.append(str(z1))
                 po=1
        if po==0:
            d = driver.find_element_by_css_selector("i.icon-edit").click()
            ti.sleep(1)
            d = driver.find_element_by_css_selector(".destructive div.ripple-container").click()
            ti.sleep(1)
            d = driver.find_element_by_css_selector("button.danger").click()
            ti.sleep(4)
        else:
            continue
except:
    ferrf=0
ti.sleep(1)
d=driver.find_element_by_css_selector(".DropdownMenu button").click()
ti.sleep(1)
d=driver.find_element_by_css_selector(".Button div.ripple-container").click()
ti.sleep(1)
d=driver.find_element_by_css_selector("a:nth-of-type(2)").click()
ti.sleep(1)
if len(df)>0:
    y=max(df)+1
else:
    y=0
w=driver.find_element_by_css_selector(".btn-icon.btn-menu-toggle > div.c-ripple").click()
ti.sleep(1)
w=driver.find_element_by_css_selector(".tgico-user div").click()
ti.sleep(1)
re=1
f2=[]
f3=[]
for i in range(0,len(r)):
    for j in [fi]:
        if r[i] in j:
            continue
    re=re+1
    o = 0
    s=y
    y=y+1
    e=0
    try:
        d = driver.find_element_by_css_selector(".btn-circle.is-visible div").click()
    except:
        ti.sleep(1)
        driver.find_element_by_css_selector("span.btn-icon").click()
        ti.sleep(10)
        d = driver.find_element_by_css_selector(".btn-circle.is-visible div").click()
    d2 = driver.find_element_by_css_selector("div.input-field:nth-of-type(1) div.input-field-input")
    y1="k"+str(y)
    d2.send_keys(y1)
    d3 = driver.find_element_by_css_selector(".input-field-phone div.input-field-input")
    ti.sleep(0.5)
    d3.clear()
    d3.send_keys(r[i])
    ti.sleep(3)
    d4 = driver.find_element_by_css_selector(".btn-primary div.c-ripple").click()
    try:
        d4 = driver.find_element_by_css_selector(".btn-primary div.c-ripple").click()
        o = 1
    except:
        rtrttrtrg = 2
    if o == 1:
        f2.append(str(r[i]).replace("+",""))
        y=s
        d5 = driver.find_element_by_css_selector("span.btn-icon").click()
    if o==0:
        f3.append(str(r[i]).replace("+", ""))
    ti.sleep(1)
ti.sleep(1)
d=driver.find_element_by_css_selector(".active button.btn-icon").click()
ti.sleep(1)
d=driver.find_element_by_css_selector(".btn-icon.btn-menu-toggle > div.c-ripple").click()
ti.sleep(1)
d=driver.find_element_by_css_selector(".z div").click()
ti.sleep(1)
d5=driver.find_element_by_css_selector(".DropdownMenu div.ripple-container").click()
ti.sleep(1)
d5=driver.find_element_by_css_selector(".top div:nth-of-type(3)").click()
x=1
f=[]
f1=[]
f4=[]
ti.sleep(1)
d=driver.find_element_by_css_selector("input.form-control").click()
d=driver.find_element_by_css_selector("input.form-control")
d.send_keys("k")
ti.sleep(10)
d5 = driver.find_element_by_css_selector(".chat-list > div.ListItem:nth-of-type(1) div.ripple-container").click()
s=driver.current_url
s=s.split("#")
f.append(s[1])
d=driver.find_element_by_css_selector(".chat-info-wrapper h3")
f1.append(d.text)
d = driver.find_element_by_css_selector(".chat-info-wrapper h3").click()
ti.sleep(1)
d = driver.find_element_by_class_name("multiline-item")
z1 = d.text
z1 = z1.replace("Phone", "")
z1 = z1.replace(" ", "")
z1 = z1.replace("+", "")
z1 = z1.replace("\n", "")
f4.append(str(z1))
ti.sleep(1)
i=2
while x<2:
    r=".chat-list > div:nth-of-type("
    r1=") div.ripple-container"
    t=r+str(i)+r1
    try:
        d5 = driver.find_element_by_css_selector(t).click()
    except:
        break
    s = driver.current_url
    s = s.split("#")
    f.append(s[1])
    d = driver.find_element_by_css_selector(".chat-info-wrapper h3")
    f1.append(d.text)
    ti.sleep(1)
    d = driver.find_element_by_class_name("multiline-item")
    z1 = d.text
    z1 = z1.replace("Phone", "")
    z1 = z1.replace(" ", "")
    z1 = z1.replace("+", "")
    z1 = z1.replace("\n", "")
    f4.append(str(z1))
    i = i + 1
    ti.sleep(1)
for i in range(1,s1+1):
    d = str(q) + str(i)
    z=str(sheet[d].value)
    for j in [f2]:
        if z in j:
            sheet[d] = "f"
            wb.save(e)
assert len(f)==len(f1)==len(f4)
ti.sleep(1)
d=driver.find_element_by_css_selector(".DropdownMenu button").click()
ti.sleep(1)
d=driver.find_element_by_css_selector(".DropdownMenu div.ripple-container").click()
ti.sleep(1)
d=driver.find_element_by_css_selector("a:nth-of-type(2)").click()
ti.sleep(2)
for i in range(0,len(f)):
    try:
        d = driver.find_element_by_xpath('//*[@id="folders-container"]/div/div[1]/ul/li[1]/div[1]').click()
    except:
        try:
            driver.find_element_by_css_selector("#folders-container > div > div.chatlist-top.has-contacts > ul > li:nth-child(1) > div.c-ripple").click()
        except:
            continue
    ti.sleep(1)
    d = driver.find_element_by_css_selector("div.inner").click()
    ti.sleep(1)
    d = driver.find_element_by_css_selector(".selection-container-forward div").click()
    c1 = ".no-shadow li"
    ti.sleep(1)
    try:
        d = driver.find_element_by_css_selector("input.selector-search-input")
        d.send_keys(str(f1[i]))
        ti.sleep(1)
        try:
            d = driver.find_element_by_css_selector(c1).click()
        except:
            try:
                c = ".no-shadow li[data-peer-id="
                c1 = c + "'" + str(f[i]) + "'" + "]"
                d = driver.find_element_by_css_selector(c1).click()
            except:
                c = "li[data-peer-id="
                c1 = c + "'" + str(f[i]) + "'" + "]"
                d = driver.find_element_by_css_selector(c1).click()
    except:
        d = driver.find_element_by_css_selector("span.btn-icon")
        ti.sleep(1)
        d = driver.find_element_by_css_selector(".selection-container-left button")
        continue
    d4=driver.find_element_by_css_selector(".active div.no-scrollbar")
    d4.send_keys(Keys.ENTER)
    for j in range(1, s1 + 1):
        d = str(q) + str(j)
        z = str(sheet[d].value)
        if z==f4[i]:
            sheet[d] = "s"
            wb.save(e)
            break
    ti.sleep(1)
ti.sleep(4)
try:
    driver.find_element_by_css_selector("span.btn-icon").click()
except:
    po=0
d=driver.find_element_by_css_selector(".sidebar-tools-button > div.c-ripple").click()
ti.sleep(1)
d=driver.find_element_by_css_selector(".z div").click()
ti.sleep(1)
delete1(driver)
ti.sleep(1)
driver.close()

