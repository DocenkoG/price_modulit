# -*- coding: UTF-8 -*-
import os
import os.path
import logging
import logging.config
import sys
import configparser
import time
# import elittech_downloader
import shutil
import csv
import requests, lxml.html



def convert2csv( cfg ):
    filename_in  = cfg.get('basic','filename_in')
    filename_out = cfg.get('basic','filename_out')

    file_in  = open( filename_in, 'r', newline='', encoding='cp1251')
    file_out = open( filename_out,'w', newline='')
    csvReader = csv.DictReader(file_in)
    csvWriter = csv.DictWriter(file_out, fieldnames=[
        'бренд',
        'группа',
        'подгруппа',
        'код',
        'код производителя',
        'наименование',
        'описание',
        'закупка',
        'продажа',
        'валюта',
        'наличие',
        '?'])

    for k in range (0, len(csvReader.fieldnames)):
        csvReader.fieldnames[k] = csvReader.fieldnames[k].lower()
    print(csvReader.fieldnames)
    csvWriter.writeheader()
    recOut = {}
    for recIn in csvReader:
        recOut['бренд']        = recIn['бренд']
        recOut['группа']       = recIn['группа']
        recOut['подгруппа']    = recIn['подгруппа']
        recOut['код']          = recIn['наименование']
        recOut['код производителя'] = recIn['код производителя']
        recOut['наименование'] = recIn['бренд']+' '+recIn['наименование']+' '+recIn['описание']
        recOut['описание']     = recIn['бренд']+' '+recIn['наименование']+' '+recIn['описание']+' Код продавца: '+recIn['артикул']+' код производителя: '+recIn['код производителя']
        recOut['продажа']      = recIn['розничная']
        try:
            recOut['закупка']  = float(recIn['розничная']) * 0.7
        except:
            recOut['закупка']  = 0.1
        #
        recOut['валюта']       = recIn['валюта']
        recOut['наличие']      = recIn['наличие']
        recOut['?']            = '?'
        #print(recOut)
        csvWriter.writerow(recOut)
    log.info('Обработано '+ str(csvReader.line_num) +'строк.')
    file_in.close()
    file_out.close()



def download( cfg ):
    from selenium import webdriver
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.remote.remote_connection import LOGGER
    LOGGER.setLevel(logging.WARNING)
    
    retCode     = False
    filename_new= cfg.get('download','filename_new')
    filename_old= cfg.get('download','filename_old')
    login       = cfg.get('download','login'    )
    password    = cfg.get('download','password' )
    url_lk      = cfg.get('download','url_lk'   )
    url_file    = cfg.get('download','url_file' )

    download_path= os.path.join(os.getcwd(), 'tmp')
    if not os.path.exists(download_path):
        os.mkdir(download_path)

    for fName in os.listdir(download_path) :
        os.remove( os.path.join(download_path, fName))
    dir_befo_download = set(os.listdir(download_path))
        
    if os.path.exists('geckodriver.log') : os.remove('geckodriver.log')
    try:
        ffprofile = webdriver.FirefoxProfile()
        ffprofile.set_preference("browser.download.dir", download_path)
        ffprofile.set_preference("browser.download.folderList",2);
        ffprofile.set_preference("browser.helperApps.neverAsk.saveToDisk", 
                ",application/octet-stream" + 
                ",application/vnd.ms-excel" + 
                ",application/vnd.msexcel" + 
                ",application/x-excel" + 
                ",application/x-msexcel" + 
                ",application/zip" + 
                ",application/xls" + 
                ",application/vnd.ms-excel" +
                ",application/vnd.ms-excel.addin.macroenabled.12" +
                ",application/vnd.ms-excel.sheet.macroenabled.12" +
                ",application/vnd.ms-excel.template.macroenabled.12" +
                ",application/vnd.ms-excelsheet.binary.macroenabled.12" +
                ",application/vnd.ms-fontobject" +
                ",application/vnd.ms-htmlhelp" +
                ",application/vnd.ms-ims" +
                ",application/vnd.ms-lrm" +
                ",application/vnd.ms-officetheme" +
                ",application/vnd.ms-pki.seccat" +
                ",application/vnd.ms-pki.stl" +
                ",application/vnd.ms-word.document.macroenabled.12" +
                ",application/vnd.ms-word.template.macroenabed.12" +
                ",application/vnd.ms-works" +
                ",application/vnd.ms-wpl" +
                ",application/vnd.ms-xpsdocument" +
                ",application/vnd.openofficeorg.extension" +
                ",application/vnd.openxmformats-officedocument.wordprocessingml.document" +
                ",application/vnd.openxmlformats-officedocument.presentationml.presentation" +
                ",application/vnd.openxmlformats-officedocument.presentationml.slide" +
                ",application/vnd.openxmlformats-officedocument.presentationml.slideshw" +
                ",application/vnd.openxmlformats-officedocument.presentationml.template" +
                ",application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" +
                ",application/vnd.openxmlformats-officedocument.spreadsheetml.template" +
                ",application/vnd.openxmlformats-officedocument.wordprocessingml.template" +
                ",application/x-ms-application" +
                ",application/x-ms-wmd" +
                ",application/x-ms-wmz" +
                ",application/x-ms-xbap" +
                ",application/x-msaccess" +
                ",application/x-msbinder" +
                ",application/x-mscardfile" +
                ",application/x-msclip" +
                ",application/x-msdownload" +
                ",application/x-msmediaview" +
                ",application/x-msmetafile" +
                ",application/x-mspublisher" +
                ",application/x-msschedule" +
                ",application/x-msterminal" +
                ",application/x-mswrite" +
                ",application/xml" +
                ",application/xml-dtd" +
                ",application/xop+xml" +
                ",application/xslt+xml" +
                ",application/xspf+xml" +
                ",application/xv+xml" +
                ",application/excel")
        if os.name == 'posix':
            driver = webdriver.Firefox(ffprofile, executable_path=r'/usr/local/Cellar/geckodriver/0.19.1/bin/geckodriver')
        elif os.name == 'nt':
            driver = webdriver.Firefox(ffprofile)
        driver.implicitly_wait(30)
        
        driver.get(url_lk)
        time.sleep(1)
        driver.set_page_load_timeout(10)
        driver.find_element_by_id("Email").clear()
        driver.find_element_by_id("Email").send_keys(login)
        driver.find_element_by_id("Password").clear()
        driver.find_element_by_id("Password").send_keys(password)

        driver.find_element_by_xpath(u"//input[@value='Войти']").click()
        driver.find_element_by_link_text(u"Kабинет").click()
        driver.find_element_by_link_text(u"Мои документы").click()
        driver.find_element_by_xpath("(//button[@type='submit'])[2]").click()
        '''
        time.sleep(1)
        try:
            driver.get(url_file)
            time.sleep(10)
        except Exception as e:
            log.debug(e)
        #print(driver.page_source)
        #driver.find_element_by_css_selector("input.button-container-m.btn_ExportAll").click()
        #time.sleep(50)
        '''
        time.sleep(10)
        driver.quit()

    except Exception as e:
        log.debug('Exception: <' + str(e) + '>')

    dir_afte_download = set(os.listdir(download_path))
    new_files = list( dir_afte_download.difference(dir_befo_download))
    print(new_files)
    if len(new_files) == 0 :        
        log.error( 'Не удалось скачать файл прайса ')
        return False
    elif len(new_files)>1 :
        log.error( 'Скачалось несколько файлов. Надо разбираться ...')
        return False
    else:   
        new_file = new_files[0]                                                     # загружен ровно один файл. 
        new_ext  = os.path.splitext(new_file)[-1].lower()
        DnewFile = os.path.join( download_path,new_file)
        new_file_date = os.path.getmtime(DnewFile)
        log.info( 'Скачанный файл ' +new_file + ' имеет дату ' + time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(new_file_date) ) )
        
        print(new_ext)
        if new_ext in ('.xls','.xlsx','.xlsb','.xlsm','.csv'):
            if os.path.exists( filename_new) and os.path.exists( filename_old): 
                os.remove( filename_old)
                os.rename( filename_new, filename_old)
            if os.path.exists( filename_new) :
                os.rename( filename_new, filename_old)
            shutil.copy2( DnewFile, filename_new)
            return True



def config_read( cfgFName ):
    cfg = configparser.ConfigParser(inline_comment_prefixes=('#'))
    if  os.path.exists('private.cfg'):     
        cfg.read('private.cfg', encoding='utf-8')
    if  os.path.exists(cfgFName):     
        cfg.read( cfgFName, encoding='utf-8')
    else: 
        log.debug('Нет файла конфигурации '+cfgFName)
    return cfg



def is_file_fresh(fileName, qty_days):
    qty_seconds = qty_days *24*60*60 
    if os.path.exists( fileName):
        price_datetime = os.path.getmtime(fileName)
    else:
        log.error('Не найден файл  '+ fileName)
        return False

    if price_datetime+qty_seconds < time.time() :
        file_age = round((time.time()-price_datetime)/24/60/60)
        log.error('Файл "'+fileName+'" устарел!  Допустимый период '+ str(qty_days)+' дней, а ему ' + str(file_age) )
        return False
    else:
        return True



def make_loger():
    global log
    logging.config.fileConfig('logging.cfg')
    log = logging.getLogger('logFile')



def processing(cfgFName):
    log.info('----------------------- Processing '+cfgFName )
    cfg = config_read(cfgFName)
    filename_out = cfg.get('basic','filename_out')
    filename_in  = cfg.get('basic','filename_in')
    filename_new = cfg.get('download','filename_new')
    
    if cfg.has_section('download'):
        print('exec-download')
        result = download(cfg)
    if is_file_fresh( filename_new, int(cfg.get('basic','срок годности'))):
        #os.system( 'brullov_converter_xlsx.xlsm')
        #convert_csv2csv(cfg)
        convert2csv(cfg)
    folderName = os.path.basename(os.getcwd())
    if os.name == 'nt' :
        if os.path.exists(filename_out)  : shutil.copy2(filename_out , 'c://AV_PROM/prices/' + folderName +'/'+filename_out)
        if os.path.exists('python.log')  : shutil.copy2('python.log',  'c://AV_PROM/prices/' + folderName +'/python.log')
        if os.path.exists('python.log.1'): shutil.copy2('python.log.1','c://AV_PROM/prices/' + folderName +'/python.log.1')
    


def main( dealerName):
    """ Обработка прайсов выполняется согласно файлов конфигурации.
    Для этого в текущей папке должны быть файлы конфигурации, описывающие
    свойства файла и правила обработки. По одному конфигу на каждый 
    прайс или раздел прайса со своими правилами обработки
    """
    make_loger()
    log.info('          '+dealerName )
    for cfgFName in os.listdir("."):
        if cfgFName.startswith("cfg") and cfgFName.endswith(".cfg"):
            processing(cfgFName)


if __name__ == '__main__':
    myName = os.path.basename(os.path.splitext(sys.argv[0])[0])
    mydir    = os.path.dirname (sys.argv[0])
    print(mydir, myName)
    main( myName)
