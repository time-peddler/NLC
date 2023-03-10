import os, time,re,requests,json,math,random,sys
from fake_useragent import UserAgent
from guotu_sav_to_excel import NLC_parse
import logging

start_year=1999
end_year=2010

REQ_HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/101.0.4951.54 Safari/537.36',
}

def mkdir(path):
    path = path.strip()
    isExists = os.path.exists(path)
    if not isExists:
        os.makedirs(path)
        
def get_sav(search_word, save_file_path):

    init_url = 'http://read.nlc.cn/getSearchCode'

    res_json = requests.get(init_url,headers = REQ_HEADERS).json()
    msg = res_json['msg']
    
    search_url = f'http://opac.nlc.cn/F/{msg}?func=find-b&find_code=WRD&request={search_word}&local_base=NLC01&x=92&y=17&filter_code_1=WLN&filter_request_1=&filter_code_2=WYR&filter_request_2={start_year}&filter_code_3=WYR&filter_request_3={end_year}&filter_code_4=WFM&filter_request_4=&filter_code_5=WSL&filter_request_5='
    resp_html = requests.get(search_url,headers=REQ_HEADERS).content.decode('utf-8')
    
    # # 出现“款目”说明无词条
    # if '款目' in resp_html:
        # logging.info(f'None:{search_word}------>Not found')
        # return
        
    # 单独显示的词条
    if '上一条' in resp_html:
        list_url = re.findall('<a\s*href=\"(.+?func=short)',resp_html)[1]
        resp_html = requests.get(list_url,headers=REQ_HEADERS).content.decode('utf-8')
    
    # 总共检索条目
    total = re.findall('of\s+(\d+)\s*\(',resp_html)[0]
    
    if len(total) == 0:
        logging.info(f'None:{search_word}------>Not found')
        return
        
    save_url = re.findall('href=\"(.+?)\"\stitle=\"Save/M',resp_html)[0]
    # 如果条大于1000天则不下载
    if int(total)>=1000:
        logging.info(f'Max limited:{search_word}------>maaaaaaxx')
        return

    resp_html = requests.get(save_url,headers=REQ_HEADERS).content.decode('utf-8')

    action_url = re.findall('action=\"(.+?)\"',resp_html)[0]

    get_sav_url = action_url + '?func=short-mail&records=ALL&range=&format=037&own_format=7%23%23%23%23&own_format=200%23%23&own_format=210%23%23&own_format=30%23%23%23&own_format=6%23%23%23%23&own_format=SYS&own_format=205%23%23&SUBJECT=&NAME=&EMAIL=&text=&x=101&y=14'

    resp_html = requests.get(get_sav_url,headers=REQ_HEADERS).content.decode('utf-8')

    sav_url = re.findall('<a\s*href=\"(.+?sav)\">',resp_html)[0]

    sav_text = requests.get(sav_url,headers=REQ_HEADERS).content.decode('gb18030')
    # 保存为 sav 文件
    with open(save_file_path,'w') as f:
        f.write(sav_text)


if __name__ == '__main__':
    logging.basicConfig(
                        handlers=[
                        logging.FileHandler(
                        filename='log.log',
                        encoding='utf-8', mode='a+')],
                        format='%(asctime)s %(message)s',
                        level=getattr(logging, 'INFO')
                        )
    
    # 创建保存路径
    mkdir('./GUOTU_SUMMARY')
    
    with open('keywords.txt','r',encoding='utf-8') as f:
        keywords = [k.replace('\n','') for k in f.readlines()]
    
    for search_word in keywords:
        logging.info(f'downloading the keywords {search_word}......')
        save_file_path  = './GUOTU_SUMMARY/'+search_word+'_'+str(start_year)+'_'+str(end_year)+'.sav'
        if os.path.exists(save_file_path):
            continue
        try:
            get_sav(search_word, save_file_path)
            time.sleep(0.5)
        except Exception as err:
            logging.info(f'ERROR:{search_word}------>{err}')
            time.sleep(5)
