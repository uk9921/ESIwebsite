# -*- coding:utf-8 -*-
from __future__ import print_function
import re
from urllib2 import urlopen,Request
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import traceback
import time
# Create your spiders here
import os
os.environ.setdefault("DJANGO_SETTINGS_MODULE", "ESIwebsite.settings")
import django
django.setup()
from Connor.models import Dissertation,refer
# http://www.lib.wust.edu.cn/resource/Purchased.aspx
# 进入图书馆校内访问SCI数据库连接http://www.lib.wust.edu.cn/resource/temp.aspx?type=p&id=33&url=http://www.webofscience.com/
# 这个是整个爬虫的入口，如果此链接改变需要在此处修改

class ESIspider():
    def __init__(self):
        '''
        INIT ESIspider
        初始化，注册并获取SID
        '''
        ESIspider.req_header = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/53.0.2785.143 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'Accept-Encoding': 'gzip, deflate, sdch',
            'Connection': 'keep-alive',
            'Host': 'apps.webofknowledge.com',
            'Upgrade-Insecure-Requests': '1',
        }
        try:
            print ('Initializing the ghost driver...')
            ESIspider.LibDbUrl = "http://www.lib.wust.edu.cn/resource/temp.aspx?type=p&id=33&url=http://www.webofscience.com/"
            driver = webdriver.PhantomJS(executable_path='C://Python27//Scripts//phantomjs-2.1.1-windows//bin//phantomjs')
            driver.maximize_window()
            driver.get(ESIspider.LibDbUrl)
            try:
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, '客户反馈和技术支持')))
            finally:
                # 数据库选择WOS集
                driver.find_element_by_css_selector("[role='presentation']").click()
                driver.find_element_by_xpath('//*[@id="select2-select-results"]/li[2]').click()
                try:
                    print('Filling in the form...')
                    WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.ID, 'WOS_GeneralSearch_input_form_sb')))
                finally:
                    pass
                driver.save_screenshot('seleniumlog/WOSdb.png')
                # 选择机构扩展
                driver.find_elements_by_class_name('select2-selection__arrow')[1].click()
                inputbox = driver.find_element_by_class_name('select2-search__field')
                inputbox.send_keys(u'机构扩展')
                inputbox.send_keys(Keys.ENTER)
                print ('Sleeping 1s...')
                time.sleep(1)
                driver.save_screenshot('seleniumlog/fixinput.png')
                # 填写武汉科技大学
                input1 = driver.find_element_by_id('value(input1)')
                input1.send_keys(Keys.TAB)
                input1.send_keys('wuhan university of science technology')
                print('Sleeping 1s...')
                time.sleep(1)
                print ('Requests will be submitted... ')
                driver.save_screenshot('seleniumlog/fixFrom.png')
                # 提交
                driver.find_element_by_id('WOS_GeneralSearch_input_form_sb').click()
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'selectSortBy_.top')))
                    driver.save_screenshot('seleniumlog/firstSub.png')
                finally:
                    '''
                   处理二级页面，左侧选择articl、review类型、精炼
                   上侧选择被引频次倒序，下侧选择每页50记录
                   '''
                    print ('Searching article and review...')
                    driver.find_element_by_id('DocumentType_1').click()
                    driver.find_element_by_id('DocumentType_3').click()
                    # 第二位为精炼
                    driver.find_element_by_xpath(' // *[ @ id = "DocumentType_tr"] / div[3] / a').click()
                    try:
                        WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.ID, 'select2-selectSortBy_top-container')))
                        driver.save_screenshot('seleniumlog/article_review.png')
                    finally:
                        pass
                    #记录数
                    ESIspider.hitCount = int(filter(str.isdigit, driver.find_element_by_id('hitCount.top').text.encode('utf-8')))
                    print(ESIspider.hitCount)
                    # 排序
                    print ('Sorting the list...')
                    driver.find_element_by_id('select2-selectSortBy_top-container').click()
                    driver.find_element_by_xpath('//*[@id="select2-selectSortBy_top-results"]/li[4]').click()
                    try:
                        elment = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.ID, 'select2-selectSortBy_top-container')))
                        driver.save_screenshot('seleniumlog/sort.png')
                    finally:
                        pass
                    print ('50 pages per page are being selected...')
                    # 每页50条
                    driver.find_element_by_id('select2-selectPageSize_bottom-container').click()
                    driver.find_element_by_xpath('//*[@id="select2-selectPageSize_bottom-results"]/li[3]').click()
                    try:
                        WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.CLASS_NAME, 'smallV110')))
                        driver.save_screenshot('seleniumlog/50_perpage.png')
                    finally:
                        ESIspider.URL = driver.current_url
                        ESIspider.SID = re.findall(r"SID=(.+?)&", ESIspider.URL)[0]
                        print ('SID is:'+ESIspider.SID)
                        print ('Initialize successfully!')
        except:
            driver.quit()
            print ('The connection is closed and ghost drive is now been restarted ...')
            traceback.print_exc()
            ESIspider.__init__(self)

    def restart(self):
        '''
        restart the ghost driver
        :return:
        '''
        print('The connection is closed and ghost drive is now been restarted ...')
        ESIspider.__init__(self)
        ESIspider.get_item(self)

    def get_page(self, pageindex):
        '''
        pageindex = ESIspider.hitCount/50 if ESIspider.hitCount%50==0 else ESIspider.hitCount/50+1
        :param pageindex:
        :return:
        '''
        page = 'page='+str(pageindex)
        url = ESIspider.URL.replace('page=1', 'page=')
        url = url.replace('page=', page)
        req = Request(url, None, headers=ESIspider.req_header)
        trycount = 0
        while True:
            try:
                print (' [open '+page+']', end='')
                resp = urlopen(req, timeout=5)
                html = resp.read()
                return html
            except:
                trycount += 1
                print (' get_page'+str(pageindex)+'failed, retrying...')
                traceback.print_exc()
                time.sleep(1)
            if trycount>=10:
                ESIspider.restart(self)
                return
        return False

    def get_bibliography(self, bibliographylink, title, referpage_index = 1):
        '''
        get the bibliography and save
        :param bibliographylink:
        :param title:
        :param referpage_index:
        :return:
        '''
        req = Request(bibliographylink, None, ESIspider.req_header)
        trycount = 0
        while True:
            try:
                response = urlopen(req,timeout=5)
                bibliographypage = response.read()
                print(' [open bibliographylink] ', end='')
                break
            except:
                trycount += 1
                time.sleep(1)
                print(' urlopen bibliographylink failed, retrying...' + str(trycount))
                traceback.print_exc()
            if trycount >= 10:
                ESIspider.restart(self)
                return

        bibliographypage = BeautifulSoup(bibliographypage, 'html.parser')
        index = (referpage_index-1)*30+1
        for bid in xrange(index, index+30):
            id = 'RECORD_'+str(bid)
            bibliographya = bibliographypage.find('div', id=id)
            if bibliographya:
                refer_title = bibliographya.find('value', attrs={'lang_id': True})
                if refer_title:
                    refer_title = refer_title.get_text()
                else:
                    refer_title = 'null'
                print (' [saveing bibliography...] ', end='')
                refer.objects.create(TITLE=title, REFERENCE_TITLE = refer_title)
            else:
                break
        print(' [saved'+str(bid)+'] ', end='')
        referpage_index += 1
        nextpagebtn = bibliographypage.find('a', attrs={'title': 'Next Page'})
        if nextpagebtn:
            nextpagelink = nextpagebtn['href']
            ESIspider.get_bibliography(self, nextpagelink, title, referpage_index)

    def get_item(self):
        '''
        The business part of the crawler
        :return:
        '''
        with open('breakpoint.log') as f:
            log = f.read()
            log = eval(log)
            pageindex = int(log['page'])
            itemindex = int(log['item'])
        html = ESIspider.get_page(self, pageindex)
        html = BeautifulSoup(html,"html.parser")
        if itemindex > ESIspider.hitCount:
            print('\nJob done, ' + log['item'] + ' records are saved.')
            exit()
        for i in range(itemindex, (pageindex-1)*50+51):
            id  = "RECORD_"+ str(i)
            print (id+' ', end='')
            item = html.find('div', id=id)
            #不存在时id为空
            if item:
                #title,date,aulist,firstau,publication,category,wos,woscate,researchdir,refcount,mechanism
                #标题，日期，作者列表，武科大第一作者，发表物，文章类型，wos号，wos类别，研究方向，被引用次数,机构
                (title,date,authorlist,first_author,publication,category,
                 wos,woscate,researchdir,refcount,mechanism)=['null']*11
                titlevalue = item.find("value", attrs={"lang_id": ''})
                if titlevalue:
                    title = titlevalue.string
                print (' [title] ', end='')
                itemurl = item.find("a", class_="smallV110")
                # 获得发表期刊
                publicationa = item.find('a', title=re.compile('View journal information'))
                if publicationa:
                    publication = publicationa.get_text().replace('\n', '')
                print ( ' [publication] ', end='')
                #被引用次数
                ref = item.find("a", attrs={"title": re.compile("View all of the articles that cite this one")})
                if ref:
                    # 引用次数链接
                    paper_reference = "http://apps.webofknowledge.com/" + ref["href"] + "&cacheurlFromRightClick=no"
                    req = Request(paper_reference, None, ESIspider.req_header)
                    trycount = 0
                    while True:
                        try:
                            #引入次数页面
                            response = urlopen(req,timeout=5)
                            soup_paper_reference = response.read()
                            print (' [open paper_reference] ', end='')
                            break
                        except:
                            print (' urlopen('+paper_reference+') error, retrying...')
                            traceback.print_exc()
                            trycount += 1
                            time.sleep(1)
                        if trycount >= 10:
                            ESIspider.restart(self)
                            return
                    soup_paper_reference = BeautifulSoup(soup_paper_reference,'html.parser')
                    refsp = soup_paper_reference.find("span", id="CAScorecard_count_WOSCLASSIC")
                    if refsp:
                        num_of_refer = refsp.get_text()
                        refcount = num_of_refer.replace('\t','').replace('\n','').replace(' ','')
                        refcount = int(refcount)
                    else:
                        refcount = 0
                else:
                    refcount = 0
                print (' [ref] ', end='')
                #文章链接
                if itemurl:
                    itemurl = "http://apps.webofknowledge.com" + itemurl['href']
                # 获得点击标题进入后的详细页面
                req = Request(itemurl, None, ESIspider.req_header)
                trycount = 0
                while True:
                    try:
                        response = urlopen(req,timeout=10)
                        detailpage = response.read()
                        print(' [open detailpage] ', end='')
                        break
                    except:
                        trycount += 1
                        time.sleep(1)
                        traceback.print_exc()
                        print (itemurl)
                        print(' urlopen itemurl failed, retrying...' + str(trycount))
                    if trycount>=10:
                        ESIspider.restart(self)
                        return

                detailpage = BeautifulSoup(detailpage,'html.parser')

                #提取内容

                #wos号
                wosstring = detailpage.find(string = re.compile("WOS:[0-9]+"))
                if wosstring:
                    wos = wosstring.replace('\t','').replace('\n','').replace(' ','')
                print (' [wos] ', end='')
                #日期
                spans = detailpage.find_all('span')
                for span in spans:
                    if span.find(text=re.compile('Published:')):
                        date = span.find_next_sibling('value').string
                        break
                print (' [date] ', end='')
                infodiv = detailpage.find('div',class_="block-record-info")
                if infodiv:
                    authorlist = infodiv.get_text()
                    p = re.compile(ur'[By]|:')
                    authorlist = p.sub('',authorlist)
                    authorlist = authorlist.replace(' ','').replace('\n','').encode('utf-8')
                    list_author = authorlist.split(';')
                else:
                    continue
                print (' [authorlist] ', end='')
                #获得机构列表 此处有问题
                for j in ['1','2','3','4','5','6']:
                    univerminid = 'addressWOS:000238230900018-'+j
                    univermin = detailpage.find('a', id=univerminid)
                    if univermin:
                        mechanism += univermin.get_text().replace('\n', '').replace('  ',' ')
                    else:
                        break
                print (' [mechanism] ', end='')
                #获得文献类型 article review
                catesp = detailpage.find('span', text=re.compile('Document Type:'))
                if catesp:
                    category = catesp.find_parent(True).text
                    category = category[category.find(':')+1:]
                print (' [catesp] ', end='')
                #获得研究方向
                researchdirsp = detailpage.find('span', text=re.compile('Research Areas:'))
                if researchdirsp:
                    researchdir = researchdirsp.find_parent(True).text
                    researchdir = researchdir[researchdir.find(':')+1:]
                print (' [researchdir] ', end='')
                #获得WOS类别
                woscatesp = detailpage.find('span', text=re.compile('Web of Science Categories:'))
                if woscatesp:
                    woscatep = woscatesp.find_parent(True).text
                    woscatep = woscatep[woscatep.find(':')+1:]
                print (' [woscatesp] ', end='')
                #获得第一位武科大作者
                for author in list_author:
                    if '[' in author:
                        author_index = re.findall(ur'(?<=\[).*(?=\])', author)[0]
                        if not author_index:
                            continue
                        if ',' in author_index:
                            author_index = author_index.split(',')
                            for index in author_index:
                                uniid = 'research_pref_org_exp_link_' + index
                                university = detailpage.find('span', id=uniid)
                                if university:
                                    university = university.find('span', class_='hitHilite')
                                if university:
                                    first_author = author[:-1].split('[')[0]
                                    break
                        else:
                            uniid = 'research_pref_org_exp_link_' + author_index
                            university = detailpage.find('span', id=uniid)
                            if university:
                                university = university.find('span', class_='hitHilite')
                            if university:
                                first_author = author[:-1].split('[')[0]
                                break
                    else:
                        pass
                print (' [first_author] ', end='')
                # 参考文献入库
                sidebardiv = detailpage.find('div', class_='l-column-sidebar2')
                if sidebardiv:
                    bibliographya = sidebardiv.find('a', attrs={'title': 'View this record’s bibliography'})
                    if bibliographya:
                        bibliographylink = 'http://apps.webofknowledge.com/' + bibliographya['href']
                        ESIspider.get_bibliography(self, bibliographylink, title)
                #入库
                print (' [saving Dissertation] ')
                Dissertation.objects.create(TITLE=title, DATE=date, AULIST=authorlist, AUWUST=first_author,PUBLICATION=publication,
                                            CATECORY=category, WOSID=wos, WOSCATE=woscatep, RESEARCHDIR=researchdir, REFERCOUNT=refcount)
                itemindex = int(i) + 1
                with file('breakpoint.log', 'w') as f:
                    log = {
                        'page': pageindex,
                        'item': itemindex
                    }
                    print('[',log['page'], log['item'],']', end='')
                    f.write(str(log))
                    print(' [Progress saved.] ')
            else:
                if i > ESIspider.hitCount:
                    print ('\nJob done, '+ str(i) +' records are saved.')
                    exit()
                else:
                    ESIspider.restart(self)
                    return

        with file('breakpoint.log', 'w') as f:
            log = {
                'page': pageindex+1,
                'item': itemindex
            }
            f.write(str(log))
            print(' [Progress saved.] ')

        esi.get_item()
esi = ESIspider()
# esi.get_item()
print ('Done')
