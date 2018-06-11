# -*- coding: UTF-8 -*-
import configobj


class ConfigUtil:
    def __init__(self):
        conf_ini = "./test.ini"
        self.__configObj = configobj.ConfigObj(
            'D:/developer/python_workspace/python_Excel2SQL/e2s_config.ini', encoding='UTF8')
        pass

    def getConfigString(self, key):
        #return self.__configObj['MAIN'][key]dsad
        return self.__configObj[key]
