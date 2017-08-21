# from __future__ import unicode_literals
# -*- coding:utf-8 -*-
from django.db import models

# Create your models here.
class UserInfo(models.Model):
    user = models.CharField(max_length = 32)
    pwd = models.CharField(max_length = 32)

class EsiDissertation(models.Model):
    index = models.CharField(max_length = 10, null=True)
    title = models.CharField(max_length = 500, null=True)
    author = models.CharField(max_length = 500, null=True)
    context = models.TextField(max_length = 500, null=True)
    publication = models.CharField(max_length = 255, null=True)

class locationconf(models.Model):
    user = models.CharField(max_length = 32, null=True)
    page = models.CharField(max_length = 5, null=True)
    li = models.CharField(max_length=5, null=True)
    time = models.DateField(null=True)

class Dissertation(models.Model):
    TITLE = models.TextField(max_length=500, null=True)#论文标题
    DATE = models.CharField(max_length=100, null=True)#出版日期
    AULIST = models.TextField(max_length=500, null=True)#作者列表
    AUWUST = models.CharField(max_length=100, null=True)#第一位武科大作者
    PUBLICATION = models.CharField(max_length=100, null=True)#出版物
    CATECORY = models.CharField(max_length=100, null=True)#类别
    WOSID = models.CharField(max_length=100, null=True)#wos号
    WOSCATE = models.CharField(max_length=500, null=True)#wos类别
    RESEARCHDIR = models.CharField(max_length=500, null=True)#研究方向
    REFERCOUNT = models.IntegerField(null=True)#被引频次（详细页面）
    MECHANISM = models.TextField(max_length=500, null=True)#作者机构
    TOTALREFCOUNT = models.IntegerField(null=True)#总被引频次
    HOT = models.BooleanField(default=False)#是否为热点论文
    HIGHTREF = models.BooleanField(default=False)#是否为高被引论文


class refer(models.Model):
    TITLE = models.CharField(max_length=500, null=True)#主动引用的论文标题（一定为本校）
    REFERENCE_TITLE = models.CharField(max_length=500, null=True)#被参考的论文（可能为外校）

class refercount(models.Model):
    TITLE = models.CharField(max_length=500, null=True)
    RC_YS = models.IntegerField(null=True)  # 年初被引频次（详细页面）
    TRC_YS = models.IntegerField(null=True)  # 年末总被引频次
    #被引频次分两种类型保存在以下两个字段，中间用英文逗号‘,’分割。
    #YYYY-MMM-DD RC1（Refer Count）,YYYY-MMM-DD RC2
    HISTORYRC = models.IntegerField(null=True)   # 被引频次（详细页面）
    HISTORYTRC = models.IntegerField(null=True)   # 总被引频次

class journals(models.Model):
    TITLE = models.CharField(max_length=500, null=True)
    TITLE29 = models.CharField(max_length=300, null=True)
    TITLE20 = models.CharField(max_length=300, null=True)
    CATE = models.CharField(max_length=500, null=True)

class staffs(models.Model):
    INSTITUTION = models.CharField(max_length=300, null=True)
    STAFFNAME_CN = models.CharField(max_length=100, null=True)
    STAFFNAME_EN = models.CharField(max_length=300, null=True)