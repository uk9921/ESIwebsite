# from __future__ import unicode_literals

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
    TITLE = models.TextField(max_length=500, null=True)
    DATE = models.CharField(max_length=100, null=True)
    AULIST = models.TextField(max_length=500, null=True)
    AUWUST = models.CharField(max_length=100, null=True)
    PUBLICATION = models.CharField(max_length=100, null=True)
    CATECORY = models.CharField(max_length=100, null=True)
    WOSID = models.CharField(max_length=100, null=True)
    WOSCATE = models.CharField(max_length=500, null=True)
    RESEARCHDIR = models.CharField(max_length=500, null=True)
    REFERCOUNT = models.IntegerField(null=True)
    MECHANISM = models.TextField(max_length=500, null=True)

class refer(models.Model):
    TITLE = models.CharField(max_length=500, null=True)
    REFERENCE_TITLE = models.CharField(max_length=500, null=True)

class journals(models.Model):
    TITLE = models.CharField(max_length=500, null=True)
    TITLE29 = models.CharField(max_length=300, null=True)
    TITLE20 = models.CharField(max_length=300, null=True)
    CATE = models.CharField(max_length=500, null=True)