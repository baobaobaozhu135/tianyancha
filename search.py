# _*_coding:utf-8_*_

import xlrd
import xlwt
from selenium import webdriver

def infomation_find(driver,text_keys):
    search_target = driver.find_element_by_id("header-company-search")
    search_target.clear()
    search_target.send_keys(text_keys)
    searchs = driver.find_elements_by_css_selector(".search>div>[class='input-group-btn btn -sm btn-primary']")
    for ele in searchs:
        if ele.text == u"天眼一下":
            ele.click()
            break

    driver.implicitly_wait(10)
    # text_header = driver.find_elements_by_css_selector(".result-list>.search-result-single >.content>.header>"\
    # +"a[tyc-event-ch='CompanySearch.Company']>text>em")[0]
    # （猜想）click事件是基于当前页面位置点击的，当滚动条滚到看不见元素的地方，则点击不到
    text_href = driver.find_elements_by_css_selector(".result-list>.search-result-single >.content>.header>" \
                                                     + "a[tyc-event-ch='CompanySearch.Company']")[0].get_attribute(
        "href")

    if text_href != None:
        print text_href
        driver.get(text_href)
    driver.implicitly_wait(10)
    html_name = driver.find_element_by_css_selector("#company_web_top>.box>.content>.header>h1.name").text
    if html_name == text_keys:
        ls = []
        fddbr = driver.find_element_by_css_selector(
            ".humancompany>.name>.link-click").text  # 法定代表人
        zczb = driver.find_element_by_css_selector(
            "table.table>tbody>tr:nth-child(1)>td:nth-child(2)>div:nth-child(2)").get_attribute("title")  # 注册资本
        zcsj = driver.find_element_by_css_selector(
            "table.table>tbody>tr:nth-child(2)>td>div:nth-child(2)>text").text  # 注册时间
        gszt = driver.find_element_by_css_selector(
            "table.table>tbody>tr:nth-child(3)>td>.num-opening").text  # 公司状态

        gszch = driver.find_element_by_css_selector(
            "table[class='table -striped-col -border-top-none']>tbody>tr:nth-child(1)>td:nth-child(2)").text  # 工商注册号
        zzjgdm = driver.find_element_by_css_selector(
            "table[class='table -striped-col -border-top-none']>tbody>tr:nth-child(1)>td:nth-child(4)").text  # 组织机构代码

        tyxydm = driver.find_element_by_css_selector(
            "table[class='table -striped-col -border-top-none']>tbody>tr:nth-child(2)>td:nth-child(2)").text  # 统一信用代码
        gslx = driver.find_element_by_css_selector(
            "table[class='table -striped-col -border-top-none']>tbody>tr:nth-child(2)>td:nth-child(4)").text  # 公司类型

        hy = driver.find_element_by_css_selector(
            "table[class='table -striped-col -border-top-none']>tbody>tr:nth-child(3)>td:nth-child(4)").text  # 行业

        rygm = driver.find_element_by_css_selector(
            "table[class='table -striped-col -border-top-none']>tbody>tr:nth-child(5)>td:nth-child(4)").text  # 人员规模

        zcdz = driver.find_element_by_css_selector(
            "table[class='table -striped-col -border-top-none']>tbody>tr:nth-child(8)>td:nth-child(2)").text  # 注册地址

        # jyfw = driver.find_element_by_css_selector("table[class='table -striped-col -border-top-none']>tbody>option:nth-child(8)>"\
        # + "option:nth-child(1) .tyc-num") # 经营范围

        ls.append(fddbr)
        ls.append(zczb)
        ls.append(zcsj)
        ls.append(gszt)
        ls.append(gszch)
        ls.append(zzjgdm)
        ls.append(tyxydm)
        ls.append(gslx)
        ls.append(hy)
        ls.append(rygm)
        ls.append(zcdz)
        return ls

#先进入进行查询，以进入查询页面
url = "https://www.tianyancha.com/search"
excel_read_url = ur"F:\所有企业名单.xlsx" # excel读取地址
#excel_write_url = ur"F:\所有企业信息.xlsx" # excel写入地址

#读取excel
r_workbook = xlrd.open_workbook(excel_read_url)
w_workbook = xlwt.Workbook(encoding="utf-8")
w_sheet = w_workbook.add_sheet("myWorkSheet")


# 读和写两个excel
r_sheet = r_workbook.sheet_by_index(0)
#w_sheet = r_workbook.sheet_by_index(0)
# 启动火狐并进入查询页面
driver = webdriver.Firefox()
driver.get(url)

row_count = r_sheet.nrows
for i in range(1, row_count):
    try:
        cell_content = r_sheet.cell(i,0).value
        if cell_content != None:
            ls = infomation_find(driver, cell_content)
            w_sheet.write(i, 0 , cell_content)
            w_sheet.write(i, 1, ls[0])
            w_sheet.write(i, 2, ls[1])
            w_sheet.write(i, 3, ls[2])
            w_sheet.write(i, 4, ls[3])
            w_sheet.write(i, 5, ls[4])
            w_sheet.write(i, 6, ls[5])
            w_sheet.write(i, 7, ls[6])
            w_sheet.write(i, 8, ls[7])
            w_sheet.write(i, 9, ls[8])
            w_sheet.write(i, 10, ls[9])
            w_sheet.write(i, 11, ls[10])

    except:
        continue
w_workbook.save("allInfomation.xls")


