# -*- coding:utf-8 -*-
from django.shortcuts import render
from Connor import models
from django.http import HttpResponse
import time
import json
import os
import re
import xlrd
import sqlite3

# Create your views here.

# 登陆界面控制器
def login(request):
    return render(request, "login.html")
# 主界面框架控制器
def index(request):
    if request.method == "POST":
        username = request.POST.get("_ctl0:txtusername", None)
        password = request.POST.get("_ctl0:txtpassword", None)
        if not models.UserInfo.objects.filter(user="nlp", pwd="nlp503"):
            models.UserInfo.objects.create(user="nlp", pwd="nlp503")
        info = models.UserInfo.objects.filter(user=username, pwd=password)
        if info:
            return render(request, "index.html")
        else:
            return render(request,"login.html",{"message":"用户不存在或密码错误"})
    else:

        return render(request, "login.html", {"message": "走正门"})
# 主界面顶部控制器
def topFrame(request):
    return render(request,"topFrame.html")
# 主界面顶部第二栏控制器
def colFrame(request):
    import time
    data = time.strftime('%Y-%m-%d', time.localtime(time.time()))
    return render(request, "colFrame.html",{"data" : data})
# 主界面左侧菜单控制器
def MenuFrame(request):
    return render(request,"MenuFrame.html")
# 主界面左侧展开收起控制器
def pushRLFrame(request):
    return render(request,"pushRLFrame.html")
# 主界面默认内容控制器
def PageFrame(request):
    import tools
    import time
    startdata, enddata = tools.daterange()
    #读取配置文件location.conf
    #这种方法慢了
    # if not models.locationconf.objects.count():
    #     models.locationconf.objects.create(user="nlp",page="1",li="0")
    info = models.locationconf.objects.filter(id ="1")
    if not info:
        models.locationconf.objects.create(user="nlp",page="1",li="0",time=time.strftime('%Y-%m-%d', time.localtime(time.time())))
    user = info.values()[0]["user"]
    page = info.values()[0]["page"]
    li = info.values()[0]["li"]
    saved = int(page) * int(li)
    time = info.values()[0]["time"]
    #如果收到后台请求
    if request.method == "POST":
        import sys
        sys.path.append("..")
        #爬虫
        from spider.crawl_list import ESIspider
        from Connor.models import EsiDissertation
        es = ESIspider()
        title, context, author, publication = es.get_SCIurl()
        EsiDissertation.objects.create(title=title, author=author, context=context, publication=publication)
        startdate = request.POST.get('startdate')
        enddate = request.POST.get('enddate')
        btn = "暂停"
        status = 11
        result = "Error!"
        return HttpResponse(json.dumps({
            "status": status,
            "result": result,
            "btn": btn
        }))

    return render(request,"PageFrame.html",{"startdata" : startdata, "enddata" : enddata, "user" : user, "saved" : str(saved),  "time" : time})
#论文统计控制器
def Page_lwtj(request):
    return render(request,"Page_lwtj.html")
#论文统计控制器
def spiderSen(request):
    return render(request,"PageFrame.html")

#年度论文图表
def Page_paperofYears(request):
    cur_year = int(time.strftime('%Y', time.localtime(time.time())))

    ref_count = {}
    total_count = {}
    esi_category = {'Computer Science': 0, 'Engineering': 0, 'Materials Sciences': 0, 'Biology & Biochemistry': 0,
           'Environment & Ecology': 0, 'Microbiology': 0, 'Molecular Biology & Genetics': 0, 'Social Sciences': 0,
           'Economics & Business': 0, 'Chemistry': 0, 'Geosciences': 0, 'Mathematics': 0, 'Physics': 0, 'Space Science': 0,
            'Agricultural Sciences': 0, 'Plant & Animal Science': 0, 'Clinical Medicine': 0, 'Immunology': 0,
           'Neuroscience & Behavior': 0, 'Pharmacology & Toxicology': 0, 'Psychology & Psychiatry': 0,
           'Multidisciplinary': 0, 'Non-ESI': 0}

    esi_statistics = {}

    for year in range(cur_year - 10, cur_year + 1):
        year_ref_count = 0
        year_total_count = 0
        esi_statistics[year] = esi_category.copy()
        paper_data = models.Dissertation.objects.filter(DATE__contains=year)
        for paper in paper_data:
            year_ref_count += paper.REFERCOUNT
            year_total_count += 1
            if ';' in paper.RESEARCHDIR:
                esi_statistics[year]['Multidisciplinary'] += 1
            else:
                if paper.RESEARCHDIR not in esi_statistics[year]:
                    esi_statistics[year]['Non-ESI'] += 1
                else:
                    esi_statistics[year][paper.RESEARCHDIR] += 1

        ref_count[year] = year_ref_count * 1
        # times -1 to show the data on the left in the chart
        total_count[year] = year_total_count * -1

    return render(request, "Page_paperofYears.html", {
                      'refcount': json.dumps(ref_count),
                      'totalcount': json.dumps(total_count),
                      'esi': json.dumps(esi_statistics)
                  })

#上传Excel文件并保存至static/journalsExcelFolder
def Page_journalsImport(request):

    if request.method == "POST":  # 请求方法为POST时，进行处理
        files = request.FILES.getlist("excels", None)
        if not files:
            return HttpResponse("no files for upload!")

        for f in files:
            destination = open(os.path.join(".\static\journalsExcelFolder", f.name), 'wb+')
            for chunk in f.chunks():
                destination.write(chunk)
            destination.close()

        JournalsDBAppend()
        return HttpResponse("上传成功")

    return render(request,"Page_journalsImport.html")

#解析Excel数据存入数据库
def JournalsDBAppend():

    excelfolderpath = ".\static\journalsExcelFolder\\"

    conn = sqlite3.connect('.\db.sqlite3')
    c = conn.cursor()

    deleteSql = """delete from Connor_journals"""
    c.execute(deleteSql)

    pathDir = os.listdir(excelfolderpath)

    for allDir in pathDir:
        child = os.path.join(allDir)
        excelpath = excelfolderpath+child
        workbook = xlrd.open_workbook(excelpath)
        booksheet = workbook.sheet_by_index(0)

        for row in range(1,booksheet.nrows):
            row_data = []
            for col in range(booksheet.ncols):
                cel = booksheet.cell(row, col)
                val = cel.value
                val = str(val)
                row_data.append(val)
            if booksheet.nrows == 5:
                title = row_data[0]
                title29 = row_data[0]
                title20 = row_data[1]
                cate = row_data[4]
            else:
                title = row_data[0]
                title29 = row_data[1]
                title20 = row_data[2]
                cate = row_data[5]
            c.execute("insert into Connor_journals (TITLE,TITLE29,TITLE20,CATE) values (?, ?, ?, ?)",
                      (title, title29, title20, cate))
            conn.commit()

    conn.close()
