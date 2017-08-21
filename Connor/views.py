# -*- coding:utf-8 -*-
from django.shortcuts import render
from Connor import models
from django.http import HttpResponse
import time
import json
import os
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
            request.session['username'] = username
            request.session['password'] = password
            return render(request, "index.html")
        else:
            return render(request, "login.html", {"message": "用户不存在或密码错误"})
    elif request.session.get('username', None):
        return render(request, "index.html")
    else:
        return render(request, "login.html", {"message": "走正门"})

# 主界面顶部控制器
def topFrame(request):
    if request.session.get('username', None):
        return render(request, "topFrame.html")
    else:
        return render(request, "login.html", {"message": "走正门"})

# 主界面顶部第二栏控制器
def colFrame(request):
    if request.session.get('username', None):
        import time
        data = time.strftime('%Y-%m-%d', time.localtime(time.time()))
        return render(request, "colFrame.html", {"data": data})
    else:
        return render(request, "login.html", {"message": "走正门"})

# 主界面左侧菜单控制器
def MenuFrame(request):
    if request.session.get('username', None):
        return render(request, "MenuFrame.html")
    else:
        return render(request, "login.html", {"message": "走正门"})

# 主界面左侧展开收起控制器
def pushRLFrame(request):
    if request.session.get('username', None):
        return render(request, "pushRLFrame.html")
    else:
        return render(request, "login.html", {"message": "走正门"})

# 主界面默认内容控制器
def PageFrame(request):
    if request.session.get('username', None):
        from Connor import tools
        import time
        startdata, enddata = tools.daterange()
        # 读取配置文件location.conf
        # 这种方法慢了
        # if not models.locationconf.objects.count():
        #     models.locationconf.objects.create(user="nlp",page="1",li="0")
        info = models.locationconf.objects.filter(id="1")
        if not info:
            models.locationconf.objects.create(user="nlp", page="1", li="0",
                                               time=time.strftime('%Y-%m-%d', time.localtime(time.time())))
        user = info.values()[0]["user"]
        page = info.values()[0]["page"]
        li = info.values()[0]["li"]
        saved = int(page) * int(li)
        time = info.values()[0]["time"]
        # 如果收到后台请求
        if request.method == "POST":
            import sys
            sys.path.append("..")
            # 爬虫
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

        return render(request, "PageFrame.html",
                      {"startdata": startdata, "enddata": enddata, "user": user, "saved": str(saved), "time": time})
    else:
        return render(request, "login.html", {"message": "走正门"})

#论文统计控制器
def Page_lwzl(request):
  #如果接收到了POST请求
  if request.method == "POST":
      title = request.POST.get('title', None)
      #在Dissertation数据库中模糊查询
      paper_list = models.Dissertation.objects.filter(TITLE__contains=title)
      #将变量paper_list放在data中并且返回到Page_lwzl.html中
      print(title)
      return render(request, "Page_lwzl.html", {"data": paper_list})
  #没收到POST请求则返回初始界面
  print(1)
  return render(request, "Page_lwzl.html")

#被引频次统计
def Page_lwyp(request):
    if request.session.get('username', None):
        # 在Dissertation数据库中统计各区间被引频次
        refercount_150 = models.Dissertation.objects.filter(REFERCOUNT__gte=150).count()
        refercount_101 = models.Dissertation.objects.filter(REFERCOUNT__range=(101, 150)).count()
        refercount_71 = models.Dissertation.objects.filter(REFERCOUNT__range=(71, 100)).count()
        refercount_51 = models.Dissertation.objects.filter(REFERCOUNT__range=(51, 70)).count()
        refercount_31 = models.Dissertation.objects.filter(REFERCOUNT__range=(31, 50)).count()
        refercount_21 = models.Dissertation.objects.filter(REFERCOUNT__range=(21, 30)).count()
        refercount_11 = models.Dissertation.objects.filter(REFERCOUNT__range=(11, 20)).count()
        refercount_5 = models.Dissertation.objects.filter(REFERCOUNT__range=(5, 10)).count()
        refercount_4 = models.Dissertation.objects.filter(REFERCOUNT__exact=4).count()
        refercount_3 = models.Dissertation.objects.filter(REFERCOUNT__exact=3).count()
        refercount_2 = models.Dissertation.objects.filter(REFERCOUNT__exact=2).count()
        refercount_1 = models.Dissertation.objects.filter(REFERCOUNT__exact=1).count()
        refercount_0 = models.Dissertation.objects.filter(REFERCOUNT__exact=0).count()
        dict = {">=150": refercount_150, "101~150": refercount_101, "71~100": refercount_71,
                "51~70": refercount_51, "31~50": refercount_31, "21~30": refercount_21,
                "11~20": refercount_11, "5~10": refercount_5, "4": refercount_4,
                "3": refercount_3, "2": refercount_2, "1": refercount_1,
                "0": refercount_0}

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
        for year in range(cur_year - 12, cur_year + 1):
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
        list = ['0', '1', '2', '3', '4', '5~10', '11~20', '21~30', '31~50', '51~70', '71~100', '101~150', '>=150']

        return render(request, "Page_citationFrequency.html", {
            'refcount': json.dumps(ref_count),
            'totalcount': json.dumps(dict),
            'esi': json.dumps(esi_statistics),
            'list': list,
        })
    else:
        return render(request, "login.html", {"message": "走正门"})


def Page_lwtj(request):
    if request.session.get('username', None):
        return render(request, "Page_lwtj.html")
    else:
        return render(request, "login.html", {"message": "走正门"})

#论文统计控制器
def spiderSen(request):
    if request.session.get('username', None):
        return render(request, "spider.html")
    else:
        return render(request, "login.html", {"message": "走正门"})

#年度论文图表
def Page_paperofYears(request):
    if request.session.get('username', None):

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
    else:
        return render(request, "login.html", {"message": "走正门"})


def Page_lwzl(request):
    if request.method == "POST":
        title = request.POST.get('title', None)
        paper_list = models.Dissertation.objects.filter(TITLE__contains=title)
        return render(request, "Page_lwzl.html", {"data": paper_list})
    return render(request, "Page_lwzl.html")


def Page_yygx(request):
    import time
    thisyear = int(time.strftime('%Y', time.localtime(time.time())))
    year = [x for x in range(thisyear - 10, thisyear + 1)]
    paper_pair = []
    if request.method == "POST":
        searchYear = request.POST.get('selyear', None)
    else:
        searchYear = thisyear-10
    paper_list = models.Dissertation.objects.values("TITLE").filter(DATE__contains=searchYear)[:10]
    for paper in paper_list:
        refer = models.refer.objects.filter(TITLE__contains=paper['TITLE'])
        paper_pair.append(refer)
    return render(request, "Page_yygx.html", {"year": year, "paper_pair": paper_pair})

def Page_ComputerScience(request):
    if request.session.get('username', None):
        #在Dissertation数据库中统计各区间被引频次
        #过滤得到以Computer Science 或者 Science & Technology 开头的单位 ，并统计数量
        from django.db.models import Q
        refercount_150 = models.Dissertation.objects.filter(Q(RESEARCHDIR__startswith='Computer Science') | Q(RESEARCHDIR__startswith='Science & Technology'), REFERCOUNT__gte=150).count()
        refercount_101 = models.Dissertation.objects.filter(Q(RESEARCHDIR__startswith='Computer Science') | Q(RESEARCHDIR__startswith='Science & Technology'), REFERCOUNT__range=(101, 150)).count()
        refercount_71 = models.Dissertation.objects.filter(Q(RESEARCHDIR__startswith='Computer Science') | Q(RESEARCHDIR__startswith='Science & Technology'), REFERCOUNT__range=(71, 100)).count()
        refercount_51 = models.Dissertation.objects.filter(Q(RESEARCHDIR__startswith='Computer Science') | Q(RESEARCHDIR__startswith='Science & Technology'), REFERCOUNT__range=(51, 70)).count()
        refercount_31 = models.Dissertation.objects.filter(Q(RESEARCHDIR__startswith='Computer Science') | Q(RESEARCHDIR__startswith='Science & Technology'), REFERCOUNT__range=(31, 50)).count()
        refercount_21 = models.Dissertation.objects.filter(Q(RESEARCHDIR__startswith='Computer Science') | Q(RESEARCHDIR__startswith='Science & Technology'), REFERCOUNT__range=(21, 30)).count()
        refercount_11 = models.Dissertation.objects.filter(Q(RESEARCHDIR__startswith='Computer Science') | Q(RESEARCHDIR__startswith='Science & Technology'), REFERCOUNT__range=(11, 20)).count()
        refercount_5 = models.Dissertation.objects.filter(Q(RESEARCHDIR__startswith='Computer Science') | Q(RESEARCHDIR__startswith='Science & Technology'), REFERCOUNT__range=(5, 10)).count()
        refercount_4 = models.Dissertation.objects.filter(Q(RESEARCHDIR__startswith='Computer Science') | Q(RESEARCHDIR__startswith='Science & Technology'), REFERCOUNT__exact=4).count()
        refercount_3 = models.Dissertation.objects.filter(Q(RESEARCHDIR__startswith='Computer Science') | Q(RESEARCHDIR__startswith='Science & Technology'), REFERCOUNT__exact=3).count()
        refercount_2 = models.Dissertation.objects.filter(Q(RESEARCHDIR__startswith='Computer Science') | Q(RESEARCHDIR__startswith='Science & Technology'), REFERCOUNT__exact=2).count()
        refercount_1 = models.Dissertation.objects.filter(Q(RESEARCHDIR__startswith='Computer Science') | Q(RESEARCHDIR__startswith='Science & Technology'), REFERCOUNT__exact=1).count()
        refercount_0 = models.Dissertation.objects.filter(Q(RESEARCHDIR__startswith='Computer Science') | Q(RESEARCHDIR__startswith='Science & Technology'), REFERCOUNT__exact=0).count()
        dict = {"refercount_150": refercount_150, "refercount_101": refercount_101, "refercount_71": refercount_71,"refercount_51": refercount_51,"refercount_31": refercount_31,"refercount_21": refercount_21,"refercount_11": refercount_11,"refercount_5": refercount_5,"refercount_4": refercount_4,"refercount_3": refercount_3,"refercount_2": refercount_2,"refercount_1": refercount_1,"refercount_0": refercount_0}
        return render(request, "Page_ComputerScience.html", dict)
    else:
        return render(request, "login.html", {"message": "走正门"})


#上传期刊Excel文件并保存至static/journalsExcelFolder
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
#上传职工Excel文件并保存到static/staffsExcelFolder
def Page_staffsImport(request):

    # if request.method == "POST":  # 请求方法为POST时，进行处理
    #     files = request.FILES.getlist("excels", None)
    #     if not files:
    #         return HttpResponse("no files for upload!")
    #
    #     for f in files:
    #         destination = open(os.path.join(".\static\staffsExcelFolder", f.name), 'wb+')
    #         for chunk in f.chunks():
    #             destination.write(chunk)
    #         destination.close()
    #
    #     staffsDBAppend()
    #     return HttpResponse("上传成功")

    return render(request,"Page_staffsImport.html")

#解析期刊Excel数据存入数据库
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

        for row in range(1, booksheet.nrows):
            row_data = []
            for col in range(booksheet.ncols):
                cel = booksheet.cell(row, col)
                val = cel.value
                val = str(val)
                row_data.append(val)
            if booksheet.ncols == 5:
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

#解析员工Excel导入数据库
def staffsDBAppend():
    excelfolderpath = ".\static\staffsExcelFolder\\"

    conn = sqlite3.connect('.\db.sqlite3')
    c = conn.cursor()

    deleteSql = """delete from Connor_staffs"""
    c.execute(deleteSql)

    pathDir = os.listdir(excelfolderpath)

    for allDir in pathDir:
        child = os.path.join(allDir)
        excelpath = excelfolderpath + child
        workbook = xlrd.open_workbook(excelpath)
        booksheet = workbook.sheet_by_index(0)

        for row in range(1, booksheet.nrows):
            row_data = []
            for col in range(booksheet.ncols):
                cel = booksheet.cell(row, col)
                val = cel.value
                val = str(val)
                row_data.append(val)
            institution = row_data[0]
            staffname_cn = row_data[1]
            staffname_en = row_data[2]

            c.execute("insert into Connor_staffs (INSTITUTION, STAFFNAME_CN, STAFFNAME_EN) values (?, ?, ?)",
                      (institution, staffname_cn, staffname_en))
            conn.commit()

    conn.close()

