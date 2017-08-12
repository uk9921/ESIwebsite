from urllib2 import urlopen,Request
from bs4 import BeautifulSoup
import threading
from time import ctime,sleep
headers = {
    "User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36",
    'X-Forwarded-For':'61.103.192.15',
}
# url1 = "http://ir.wust.edu.cn/AchivementDetails/Index/92630"
# url2 = "http://ir.wust.edu.cn/AchivementDetails/Index/92629"
# def click1(c):
#     for i in xrange(c):
#         req = Request(url1, headers=headers)
#         urlopen(req, timeout=5)
# def click2(c):
#     for i in xrange(c):
#         req = Request(url2, headers=headers)
#         urlopen(req, timeout=5)
# threads = []
#
# threads.append(threading.Thread(target=click1,args=(10,)))
# threads.append(threading.Thread(target=click2,args=(10,)))
# if __name__ == '__main__':
#     for t in threads:
#         t.setDaemon(True)
#         t.start()
#     for t in threads:
#         t.join()
#     print "all over %s" %ctime()

print BeautifulSoup(urlopen(Request('https://www.zhihu.com/question/37046157', headers={"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36"})),'html.parser').find('code', class_='language-python').text

