# UnboundLocalError: local variable 'driver' referenced before assignment
# 上記のエラーが出た場合はchrome driveのバージョンが低い

# このプログラムを動作させるためには以下のインストールが必要
# pyexcel
# selenium
# google chromeのドライバー
codon={
    "TTT":"F",
    "TTC":"F",
    "TTA":"L",
    "TTG":"L",
    "TCT":"S",
    "TCC":"S",
    "TCA":"S",
    "TCG":"S",
    "TAT":"Y",
    "TAC":"Y",
    "TAA":"*",
    "TAG":"*",
    "TGT":"C",
    "TGC":"C",
    "TGA":"*",
    "TGG":"W",    
    "CTT":"L",
    "CTC":"L",
    "CTA":"L",
    "CTG":"L",
    "CCT":"P",
    "CCC":"P",
    "CCA":"P",
    "CCG":"P",
    "CAT":"H",
    "CAC":"H",
    "CAA":"Q",
    "CAG":"Q",
    "CGT":"R",
    "CGC":"R",
    "CGA":"R",
    "CGG":"R",
    "ATT":"I",
    "ATC":"I",
    "ATA":"I",
    "ATG":"M",
    "ACT":"T",
    "ACC":"T",
    "ACA":"T",
    "ACG":"T",
    "AAT":"N",
    "AAC":"N",
    "AAA":"K",
    "AAG":"K",
    "AGT":"S",
    "AGC":"S",
    "AGA":"R",
    "AGG":"R",
    "GTT":"V",
    "GTC":"V",
    "GTA":"V",
    "GTG":"V",
    "GCT":"A",
    "GCC":"A",
    "GCA":"A",
    "GCG":"A",
    "GAT":"D",
    "GAC":"D",
    "GAA":"E",
    "GAG":"E",
    "GGT":"G",
    "GGC":"G",
    "GGA":"G",
    "GGG":"G"
    }

# Seleniumを使用
from operator import length_hint
from selenium import webdriver # installしたseleniumからwebdriverを呼び出せるようにする
# from selenium.webdriver.common.keys import Keys # webdriverからスクレイピングで使用するキーを使えるようにする。
import time # 今回は、プログラムをsleepするために使用
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
import pandas as pd
import traceback
import datetime
import sys

def scraping(string_web):
    try:
        options = webdriver.ChromeOptions()
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        options.use_chromium = True
        driver = webdriver.Chrome(service=Service("C://Users//hoshi//Downloads//涼君の研究//chromedriver_win32//chromedriver.exe"),options=options)
        driver.get("https://blast.ncbi.nlm.nih.gov/Blast.cgi?PROGRAM=blastp&PAGE_TYPE=BlastSearch&LINK_LOC=blasthome")
        driver.implicitly_wait(10) # 単位は秒，暗黙的な待機時間を設定，find_element等の処理時に、要素が見つかるまで指定した最大時間待機させる
        # ページ上のすべての要素が読み込まれるまで待機（15秒でタイムアウト判定）
        WebDriverWait(driver, 15).until(EC.presence_of_all_elements_located)

        # key="MALNLKSFSTNSKL"
        # key="MGVISILCNTCTKDLLLIQKNSHNMLTPSTLNTISTYCLKYANLPNFAKIKAINNSPFRRFYHGLFKQYVFFER"
        key=string_web
        driver.find_element(By.XPATH,"//*[@id='seq']").send_keys(key)
        driver.find_element(By.CSS_SELECTOR, "input.blastbutton").click()

        wait = WebDriverWait(driver, 180)
        # element=wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, selector)))
        # time.sleep(120)

        # tableElem=wait.until(EC.visibility_of_element_located((By.ID,"dscTable")))
        element=wait.until(EC.visibility_of_element_located((By.CLASS_NAME,"usa-alert-body")))

        tableElem = driver.find_element(By.ID,"dscTable")
        trs = tableElem.find_elements(By.TAG_NAME,"tr")

        #keysはいちいちfor文で回すと同じ結果しか得られないので外でまとめたほうが速い
        #   elems=trs[0].find_elements(By.TAG_NAME,"th")
        #   keys = []
        #   for elem in elems:
        #     keys.append(elem.text)

        #   values=[]#二次元配列にするため，prelistを作成
        #   for i in range(1,len(trs)):
        #       tds = trs[i].find_elements(By.TAG_NAME,"td")
        #       prelist=[]
        #       for j in range(0,len(tds)):
        #           prelist.append(tds[j].text)
        #       values.append(prelist)

        tds = trs[1].find_elements(By.TAG_NAME,"td")
        Protein=tds[1].text
        E_value=tds[8].text
        Accession=tds[11].text
        url=tds[11].find_elements(By.TAG_NAME,"a")[0].get_attribute("href")
        # url='=HYPERLINK('+url+','+ Accession+')' 
        url='=HYPERLINK("%s", "%s")' % (url, Accession) 
        # print(url)

        # values = [list(x) for x in zip(*values)] #keywのvaluesの型を合わせるために転置する

        # df = pd.DataFrame(values,columns=keys)

    except:  # クエリが見つからなかったときとtimeoutエラーの場合分け

        # if element==None:
        if not 'element' in locals():
            Protein="TIME OUT"
            E_value="TIME OUT"
            Accession="TIME OUT"
            url="TIME OUT"
            print('timeout')

        elif element.text=='':
            Protein="NOT FOUND"
            E_value="NOT FOUND"
            Accession="NOT FOUND"
            url="NOT FOUND"
        
        else:
            print(traceback.format_exc())

    finally:
        if driver is not None:
            driver.quit()
        return Protein,E_value,Accession,url

if len(sys.argv)==1:
    print('ファイルのパスをコマンドライン引数に渡してください')
    sys.exit()

ori=sys.argv[1]
base=ori.rsplit('\\',1)[0] #基準となるディレクトリ

file_name=ori
f = open(file_name, 'r')
# f = open(file_name, 'r',encoding='shift_jis')
ori_data = f.read()
f.close()
data=ori_data.split('\n',1)[1]
length=len(data)

def scraping_count(data_ori):
    count=0
    cursor=0
    data=data_ori
    while data.find('ATG',cursor)>0:
        cursor=data.find('ATG',cursor)
        end=cursor
        line=""
        # このwhileの中で探し出した開始コドンから終了コドンまでの文字列を抽出
        while True: # while trueとif breakによってdo while文のようになっている 
            th_str=data[end:end+3] # Mから文字列を抽出する
            if(not th_str in codon.keys()): 
                break
            chara=codon[th_str]
            # print(line)
            if chara=='*':
                break
            line+=chara
            end+=3
        if len(line)>=50:
            count+=1
        cursor+=1
    cursor=0
    data=data_ori[::-1]
    data=data.replace('A', '#').replace('T', 'A').replace('#', 'T')
    data=data.replace('G', '#').replace('C', 'G').replace('#', 'C')
    while data.find('ATG',cursor)>0:
        cursor=data.find('ATG',cursor)
        end=cursor
        line=""
        # このwhileの中で探し出した開始コドンから終了コドンまでの文字列を抽出
        while True: # while trueとif breakによってdo while文のようになっている 
            th_str=data[end:end+3] # Mから文字列を抽出する
            if(not th_str in codon.keys()): 
                break
            chara=codon[th_str]
            # print(line)
            if chara=='*':
                break
            line+=chara
            end+=3
        if len(line)>=50:
            count+=1
        cursor+=1
    return count

num=0
def annotation(data,count,num):
    # num=0
    lines=[]
    starts=[]
    ends=[]
    w_count=[]
    Proteins=[]
    E_values=[]
    Accessions=[]
    urls=[]
    cursor=0
    # while cursor:=data.find('ATG',cursor)>0:
    while data.find('ATG',cursor)>0:
        cursor=data.find('ATG',cursor)
        start=cursor
        end=cursor
        line=""
        # このwhileの中で探し出した開始コドンから終了コドンまでの文字列を抽出
        while True: # while trueとif breakによってdo while文のようになっている 
            th_str=data[end:end+3] # Mから文字列を抽出する
            if(not th_str in codon.keys()): 
                break
            chara=codon[th_str]
            # print(line)
            if chara=='*':
                break
            line+=chara
            end+=3
        if(len(line)>=10): #アミノ酸配列が10より大きかったら加える
            lines.append(line)
            starts.append(length-start+1)
            ends.append(length-end+1)
            w_count.append(len(line))
        cursor+=1

    for line_i in lines:
        # print(len(line_i))
        if len(line_i)<50:
            Protein="--"
            E_value="--"
            Accession="--"
            url="--"
        else:
            num+=1
            print(str(num)+'/'+str(count))
            print(datetime.datetime.now())
            Protein,E_value,Accession,url=scraping(line_i)
            
        Proteins.append(Protein)
        E_values.append(E_value)
        Accessions.append(Accession)
        urls.append(url)
    # '=HYPERLINK("https://en.wikipedia.org/wiki/2003", 2003)'

    df=pd.DataFrame({'start': starts,
                        'end': ends,
                        'string': lines,
                        'word count':w_count,
                        'Protein':Proteins,
                        'E_value':E_values,
                        # 'Accession':Accessions,
                        'URL':urls})
    return df,num

count=scraping_count(data)
# df_forward,num=annotation(data,count,num)

data=data.replace('A', '#').replace('T', 'A').replace('#', 'T')
data=data.replace('G', '#').replace('C', 'G').replace('#', 'C')
df_backward,num=annotation(data[::-1],count,num)

name=ori.rsplit('.txt',1)[0]

with pd.ExcelWriter(name+'_Anotation_backward_only.xlsx') as writer:
    # df_forward.to_excel(writer, index=False,sheet_name="forward")
    df_backward.to_excel(writer, index=False,sheet_name="backward")
    # writer.book["forword"].column_dimensions['A'].width = 15
    # writer.sheets['forward'].column_dimensions['A'].width = 8    
    # writer.sheets['forward'].column_dimensions['B'].width = 8
    # writer.sheets['forward'].column_dimensions['C'].width = 15
    # writer.sheets['forward'].column_dimensions['D'].width = 8
    # writer.sheets['forward'].column_dimensions['E'].width = 63.4
    # writer.sheets['forward'].column_dimensions['F'].width = 9.2
    # writer.sheets['forward'].column_dimensions['G'].width = 9.2
    writer.sheets['backward'].column_dimensions['A'].width = 8
    writer.sheets['backward'].column_dimensions['B'].width = 8
    writer.sheets['backward'].column_dimensions['C'].width = 15
    writer.sheets['backward'].column_dimensions['D'].width = 8
    writer.sheets['backward'].column_dimensions['E'].width = 63.4
    writer.sheets['backward'].column_dimensions['F'].width = 9.2
    writer.sheets['backward'].column_dimensions['F'].width = 9.2


#なんかto_csvだとsheet_nameの引数がないらしい


# print(df)

#進捗が分かるといいね
#この機能を搭載するには文字列をいったん順方向，逆方向から読んで，50文字以上の文字列の数を取得する必要がある

# 並列処理をすると，listに入る順番がばらばらになっちゃうから，dataframeでstartの昇順に直す必要がある
# chrome driverをいちいち立ち上げるのは冗長
# 複数タブを使った情報の取得とか？
# 最初にwebページをgetするときの待機時間とかって設定することができるの？

