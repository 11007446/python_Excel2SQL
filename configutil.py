# -*- coding: UTF-8 -*-
from configobj import ConfigObj


class ConfigUtil:
    def __init__(self):
        conf_ini = "./test.ini"
        self.__configObj = ConfigObj('e2s_config.ini', encoding='UTF8')
        pass

    def getConfigString(self, key, section='PATH'):
        return self.__configObj[section][key]
