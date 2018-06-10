# -*- coding: UTF-8 -*-
# Filename:SQL_insert_template.py


class SQL_insert_template(object):
    def __init__(self):
        self.BASE_SQL = {}
        self.VALUE_SQL = {}

    def getSQL(self):
        return [self.BASE_SQL, self.VALUE_SQL]

    pass


class SQL_insert_test(SQL_insert_template):
    '''
    科技统计外部数据导入脚本模版
    导入表 tesc
    '''

    def __init__(self):
        self.BASE_SQL = {
            'Sheet1':
            'test1'
        }
        self.VALUE_SQL = {
            'Sheet1':
            'test2'
        }

    pass


class SQL_insert_TESC(SQL_insert_template):
    '''
    科技统计外部数据导入脚本模版
    导入表 tesc
    '''

    def __init__(self):
        self.BASE_SQL = {
            'Sheet1':
            'INSERT INTO CommonContactSource (contactsinfoid,contactsname,contactsemail,contactsmobile,BelongArea,orgInfoId,status,Tesc,SGST,YPKW,CXQY,createdtime,createdby,modifiedtime,modifiedby,SZMT,PDCX,PDJX,PDYFJG,PDOTHER) VALUES '
        }
        self.VALUE_SQL = {
            'Sheet1':
            '(newid(),\'{0[1]}\',\'{0[2]}\',\'{0[3]}\',\'{0[4]}\',\'{0[5]}\',\'1\',\'1\',\'0\',\'0\',\'0\',GETDATE(),\'新科技企业统计系统数据导入\',GETDATE(),\'新科技企业统计系统数据导入\',\'0\',\'0\',\'0\',\'0\',\'0\');'
        }

    pass
