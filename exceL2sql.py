# -*- coding: UTF-8 -*-
import datetime
import openpyxl
import os
import shutil
import sql_insert_template
import configutil


def getSqlTemplate(sql_template_name):
    return ""
    pass


def loadExcel():
    '''
    读取配置文件中待解析excel文件\文件列表,以及输出sql文件路径.
    解析Excel内容,生成Sql Insert语句执行,并保存到sql文件留档
    '''
    config = configutil.ConfigUtil()
    filepaths, outputpath_base, thesqltemplate = config.getConfigString(
        'EXCELINPUTPATH'), config.getConfigString('SQLOUTPUTPATH'), config.getConfigString('SQLTEMPLATE')
    print('解析EXCEL文件源: %s' % (filepaths))
    sqltemplate = getSqlTemplate(thesqltemplate)
    if type(filepaths) is list:
        for filepath in filepaths:
            (filename, ext) = os.path.splitext(os.path.basename(filepath))
            outputpath_folder = '%s_%s' % (
                filename, str(datetime.datetime.now().strftime('%Y%m%d')))
            outputpath = outputpath_base + outputpath_folder + '/'
            e2s = Excel2sql(sqltemplate, [filepath, outputpath])
            e2s.parseWorkBook()
            pass
        pass
    elif type(filepaths) is str:
        (filename, ext) = os.path.splitext(os.path.basename(filepaths))
        outputpath_folder = '%s_%s生成' % (
            filename, str(datetime.datetime.now().strftime('%Y%m%d')))
        outputpath = outputpath_base + outputpath_folder + '/'
        e2s = Excel2sql(sqltemplate, [filepaths, outputpath])
        e2s.parseWorkBook()
        pass
    print('解析完毕', end='\n\n')


class Excel2sql(object):
    def __init__(self, sqltemplate, config):
        self.sqltemplate = sqltemplate
        self.config = config
        pass

    def parseWorkBook(self):
        filepath, outputpath = self.config

        print('     解析EXCEL文件: %s' % (filepath))
        if os.path.exists(outputpath):
            shutil.rmtree(outputpath)
        os.makedirs(outputpath)

        wb = openpyxl.load_workbook(filepath)

        for sheet in wb.sheetnames:
            print('         解析表%s开始' % (sheet))
            self.parseSheetData2SQL(wb[sheet], outputpath, 2)
            print('         解析表%s完成' % (sheet), end='\n\n')
            pass
        print('     解析完毕', end='\n\n')

    def parseSheetData2SQL(self, sheet, output_path, startRow=1, pagemaxrow=5000):

        sheet_title = sheet.title
        columnpart, valuepart = self.sqltemplate.getSQL()
        if(sheet_title not in columnpart.keys()):
            return

        maxRow, maxCol = sheet.max_row, sheet.max_column

        output_filename = sheet_title + '_' + str(
            datetime.datetime.now().strftime('%Y%m%d%H%M'))
        # with open(output_path + output_filename, 'w', encoding='utf-8') as fw:

        fw = open(output_path + output_filename +
                  '.sql', 'w', encoding='utf-8')
        for loopindex, row in enumerate(sheet.iter_rows(
                row_offset=startRow, max_row=maxRow - startRow)):
            rowlist = []
            for cell in row:
                cellvalue = cell.value
                if cellvalue is None:
                    rowlist.append('')
                else:
                    rowlist.append(cellvalue)

            mainsql = columnpart[sheet_title] + valuepart[sheet_title].format(
                rowlist)

            if (loopindex > 0 and loopindex % pagemaxrow == 0):
                fw.close()
                fw = open(
                    output_path + output_filename +
                    '_' + str(loopindex) + '.sql',
                    'w',
                    encoding='utf-8')

            fw.write(mainsql + '\n')
        fw.write('--%s数据共计%d条\n\n' % (sheet_title, loopindex))
        fw.close()
        # pass

    pass


if __name__ == "__main__":
    loadExcel(sql_insert_template.SQL_insert_TESC())
    pass
