# -*- coding: UTF-8 -*-
# Filename:SQL_insert_template.py


class SQL_base_template(object):
    def __init__(self):
        self.BASE_SQL = None
        self.VALUE_SQL = None

    def getSQL(self):
        return [self.BASE_SQL, self.VALUE_SQL]

    pass


class SQL_insert_test(SQL_base_template):
    '''
    科技统计外部数据导入脚本模版
    导入表 tesc
    '''

    def __init__(self):
        self.BASE_SQL = 'test_base'
        self.VALUE_SQL = 'test_value'

    pass


class SQL_insert_TESC(SQL_base_template):
    '''
    科技统计外部数据导入脚本模版
    导入表 tesc
    '''

    def __init__(self):
        self.BASE_SQL = 'INSERT INTO CommonContactSource (contactsinfoid,contactsname,contactsemail,contactsmobile,BelongArea,orgInfoId,status,Tesc,SGST,YPKW,CXQY,createdtime,createdby,modifiedtime,modifiedby,SZMT,PDCX,PDJX,PDYFJG,PDOTHER) VALUES '
        self.VALUE_SQL = '(newid(),\'{0[1]}\',\'{0[2]}\',\'{0[3]}\',\'{0[4]}\',\'{0[5]}\',\'1\',\'1\',\'0\',\'0\',\'0\',GETDATE(),\'新科技企业统计系统数据导入\',GETDATE(),\'新科技企业统计系统数据导入\',\'0\',\'0\',\'0\',\'0\',\'0\');'

    pass


class SQL_UPDATE_EXPERT(SQL_base_template):
    '''
    专家库银行信息更新
    导入表 tesc
    '''

    def __init__(self):
        self.BASE_SQL = 'update tab_expert_base '
        self.VALUE_SQL = 'set bankBranch=\'{0[1]}\' where id_number=\'{0[2]}\';'

    pass
