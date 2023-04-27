#v3.1 2023/02/19 プロファイル仕様変更対応
#v3.2 2023/03/03 other profile profile_dtlのセレクター仕様変更対応
#v3.3 2023/03/17 セレクターの見直し
#v3.4 2023/04/06 メルカリショップサイトのセレクターの見直し
#v3.5 2023/04/11 chromeを検索ワードごとに接続し直すよう対応、ガベージコレクション対応
#v3.6 2023/04/28 一覧の取得方法、各種セレクタの修正
import datetime
import glob
import os
import os.path
from decimal import Decimal, InvalidOperation
import xlwings as xw
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import time
from bs4 import BeautifulSoup
import re
import traceback
import warnings
import shutil
import requests
from concurrent import futures
import threading
import gc

warnings.filterwarnings('ignore')

## グローバル変数 ##
#設定情報（辞書）
ds = {}

#GoogleChromDriver（リスト）
browser = []

## イニシャル処理 ##
def initial_set(mode):

    #設定情報（グローバル変数の宣言）
    global ds
    
    #モード（X:エクセル、B：バッチ）
    ds['mode'] = mode

    #メルカリURL
    ds['merUrl'] = 'https://jp.mercari.com'

    #メルカリショップURL
    ds['merShopUrl'] = 'https://mercari-shops.com'

    #カレントディレクトリ
    ds['curDir'] = os.getcwd()

    #インプット設定ファイル
    ds['inputFile'] =r'\mercari_scraping.xlsm'
    ds['inputFull'] = ds['curDir']  + ds['inputFile']
    
    #アウトプットディレクトリ
    now = datetime.datetime.now()
    outputDir = now.strftime("%Y%m%d%H%M%S")
    if not os.path.isdir(outputDir):
        os.mkdir(outputDir)
    ds['outputDir'] = ds['curDir'] + '\\' + outputDir

    #設定ファイル（エクセルブック）の呼び出し
    if mode == 'X':
        wb = xw.Book.caller()
    else:
        App = xw.App(visible=False)
        ds['App'] = App
        wb = xw.books.open(ds['inputFull'])

    ds['wb'] = wb

    #設定シート
    ws_setting = wb.sheets('設定')
    
    #リスト開始行
    ds['listStartRow'] = 4
    
    #非表示情報の数
    ds['othCnt'] = 3

    #URL列
    ds['urlCol'] = 2
    
    #最大画像数
    ds['picMaxCnt'] = 10

    #画像最小列数
    ds['picMinCnt'] = 4

    #検索キーワードリスト
    swList = []
    i = 0
    for sw in ws_setting.range('swList'):
        if i%3 == 0:
            swDict = {}
            swDict['no'] = sw.value
        elif i%3 == 1: 
            swDict['sw'] = sw.value
        elif i%3 == 2: 
            if sw.value is None:
                swDict['pg'] = 1
            else:
                swDict['pg'] = int(sw.value)

            if swDict['no'] is not None and swDict['sw'] is not None:
                swList.append(swDict)
        i += 1
    ds['swList'] = swList
    #print(ds['swList'][1]['sw'])
    #print(len(ds['swList']))
        
    #商品の状態
    i = 1
    itemCond = ''
    for pStatus in ws_setting.range('stList'):
        if pStatus.value:
            if itemCond == '':
                itemCond = '&item_condition_id=' + str(i)
            else:
                itemCond = itemCond + ',' + str(i)
        i += 1
    ds['itemCond'] = itemCond

    #除外リスト
    exclutionList = []
    for exclution in ws_setting.range('exclutionList'):
        if exclution.value is not None:
            exclutionList.append(exclution.value)
    ds['exclutionList'] = exclutionList

    #評価除外数を取得
    ds['exScore'] = ws_setting.range('exScore').value

    #コメント除外数を取得
    ds['exCount'] = ws_setting.range('exCount').value

    #価格による抽出
    ds['minPrice'] = ws_setting.range('minPrice').value
    ds['maxPrice'] = ws_setting.range('maxPrice').value
    
    #ページごとにブックを分けるかどうか
    ds['pgDiv'] = ws_setting.range('pgDiv').value
    if ds['pgDiv'] == '' or ds['pgDiv'] is None:
        ds['pgDiv'] = 'N'
    
    #GoogleChromeDriver同時起動数を取得
    ds['pararel'] = ws_setting.range('pararel').value
    if ds['pararel'] == '' or ds['pararel'] is None:
        ds['pararel'] = 1
    else:
        ds['pararel'] = int(ws_setting.range('pararel').value)

    #プロフィールによる除外を行うかどうか
    ds['isExProf'] = ws_setting.range('isExProf').value
    if ds['isExProf'] == '' or ds['isExProf'] is None:
        ds['isExProf'] = 'N'

    #for d in ds:
    #    print(d,':',ds[d])
    

## 除外リストの文字列を含むかどうかチェック ##
def checkExclution(param):
    existFlg = False
    if len(ds['exclutionList']) > 0:
        for exclution in ds['exclutionList']:
            if exclution in param:
                existFlg = True
                break
            
    return existFlg

## GoogleChromeを起動 ##
def startGC(pararel):

    options = Options()
    options.add_argument('--headless')
    options.add_argument('--log-level=2')

    for i in range(pararel):
        driver = webdriver.Chrome(executable_path=ChromeDriverManager().install(), options=options)
        driver.implicitly_wait(10)
        driver.set_window_size('1200', '1000')
        browser.append(driver)

## GoogleChromeをクローズ ##
def closeGC(pararel):
    if len(browser) > 0:
        for i in range(pararel):
            if browser[i] is not None:
                browser[i].close()
                browser[i].quit()
        browser.clear()

## ブックの作成 ##
def bookCreate(rs):
    #メルカリリストクリア
    macro=ds['wb'].macro('allClear')
    macro()

    #メルカリリストシート
    ws_list = ds['wb'].sheets('メルカリリスト')

    ws_list.range(1,2).value = rs['sWord']
    ws_list.range(1,3).value = rs['sPage']
    ws_list.range(2,3).value = rs['maxCol']
    ws_list.range(2,4).value = rs['iRow']

    macro=ds['wb'].macro('DataOutput')
    macro()

    ws_list.range(2,3).value = '取得件数:'

    iRow = 0
    lastCol = ds['picMinCnt']
    if lastCol < rs['maxCol']:
        lastCol = rs['maxCol']

    #[0] url
    #[1-10] 画像
    #[11-14] 出品者、評価、コメント数、価格
    #[15-18] 非表示情報（商品名、商品詳細、プロフィール、画像最大数）
    for itemInfo in rs['itemList']:
        iCol = 0
        i = 0
        for attr in itemInfo:
            if i == 0:
                sUrl = attr
                ws_list.range(ds['listStartRow'] + iRow , ds['urlCol'] + iCol).value = sUrl
                iCol += 1
            elif i > 0 and i <= 10:
                if i <= lastCol:
                    if attr != '' and attr is not None:
                        setPict(ws_list, ds['listStartRow'] + iRow , ds['urlCol'] + iCol, [sUrl,attr])
                    iCol += 1
            elif  i > 10 and i <= 14:
                ws_list.range(ds['listStartRow'] + iRow , ds['urlCol'] + iCol).value = attr
                iCol += 1
            else:
                pass
            i += 1
        iRow += 1

    #エクセルファイルを保存
    ds['wb'].save()

    #エクセルファイルコピー
    if ds['mode'] == 'B':
        if ds['pgDiv'] == 'Y':
            outputFull = ds['outputDir'] + '\\' + str(rs['sNo']).zfill(3) + '_' + rs['sWord'] + '-' + rs['sPage'].zfill(3) + '.xlsm'
        else:
            outputFull = ds['outputDir'] + '\\' + str(rs['sNo']).zfill(3) + '_' + rs['sWord'] + '.xlsm'
    
        shutil.copy(ds['wb'].name, outputFull)

## 画像を取得し、エクセルシートに設定 ##
def setPict(ws, rCell, cCell, pUrl):

    pUrl2 = re.sub('jpg.*','jpg',pUrl[1])
    idx = pUrl2.rfind('/')
    idx2 = pUrl[0].rfind('/')
    filename = pUrl[0][idx2+1:] + '_' + pUrl2[idx+1:]

    filenameFull = ds['outputDir']  + '\\' + filename

    #画像を取得
    response = requests.get(pUrl[1])

    if response.status_code == 200:
        image = response.content
        with open(filenameFull, "wb") as gFile:
            gFile.write(image)
                
        gCell = ws.range(rCell,cCell)

        pshp = ws.pictures.add(filenameFull)

        #画像のサイズ調整
        picPer = 0.9
        celw = gCell.width
        celh = gCell.height
        wPer = celw / pshp.width
        hPer = celh / pshp.height
        pshpHeightOrg = pshp.height

        if wPer < hPer:
            pshp.width = pshp.width * wPer * picPer
            if pshp.height == pshpHeightOrg:
                pshp.height = pshp.height * wPer * picPer
        else:
            pshp.width = pshp.width * hPer * picPer
            if pshp.height == pshpHeightOrg:
                pshp.height = pshp.height * hPer * picPer

        pshp.top = gCell.top + (celh - pshp.height) / 2
        pshp.left = gCell.left + (celw - pshp.width) / 2

        os.remove(filenameFull)

    else:
        print('画像の取得に失敗しました（pictUrl=',pUrl[1],'）')

## 商品情報を取得する ##
def itemGetFunc(itemKey,iAllRow,maxRow,bIdx):
    
    if iAllRow < maxRow:
        print('\r'+' 処理中 : ' + str(iAllRow) +' / ' + str(maxRow),end='')
    else:
        print('\r'+' 処理中 : ' + str(iAllRow) +' / ' + str(maxRow))
    
    #使用するwebBrowserのidxを決定
    for i in range(100):
        for bNo in range(len(bIdx)):
            if bIdx[bNo]:
                break
        if bIdx[bNo]:
            bIdx[bNo] = False
            break
        else:
            if i < 99:
                time.sleep(0.5)
                continue
            else:
                raise Exception('webBrowserのインデックスの空きがありません')

    item_name = itemKey[0]
    item_url = itemKey[1]

    #print('item_name=',item_name);
    #print('item_url=',item_url);

    iPar = bNo + 1
    browser[iPar].get(item_url)

    #商品詳細ページの情報を取得
    for i in range(30):
        try:
            #変数初期化
            shop_elements = []
            item_detail = ''
            picts = []
            pList = []
            pUrlList = []
            s_name =''
            s_score = ''
            s_count = ''
            item_price = ''
            profile_dtl = ''
            s_score_p = 0
            s_score_d = 0
            shop_flg = ''

            #メルカリショップかどうかの判定
            if ds['merShopUrl'] in item_url:
                shop_flg = 'Y'

            #print('iAllRow=', str(iAllRow),'i=',i)

            # HTMLを文字コードをUTF-8に変換してから取得します。
            html = browser[iPar].page_source.encode('utf-8')

            # BeautifulSoupで扱えるようにパースします
            soup = BeautifulSoup(html, "html.parser")

            #商品詳細
            if shop_flg == 'Y':
                #item_detail = soup.select('#__next > div.css-xu1r6a > div > div.css-1rr4qq7 > div:nth-child(2) > div.css-1x15fb3')
                #item_detail = soup.select('.chakra-text.css-ic9sg9')
                elements = soup.select('.css-0 p[class="chakra-text css-naeo47"]')
                item_detail =elements[0].contents[0]
            else:
                item_elem = soup.select('#item-info mer-show-more')
                item_detail = item_elem[0].contents[0].text

            #print('item_detail=',item_detail[0:30])

            #商品詳細除外
            if checkExclution(item_detail):
                #webBrowserの解放
                bIdx[bNo] = True
                return None

            #商品画像URLを取得
            if shop_flg == 'Y':
                picts =soup.select(".chakra-stack.css-tg402c img")
            else:
                picts =soup.select('div[class="slick-slider slick-vertical slick-initialized"] .slick-list img')
            
            for pict in picts:
                pict_url = pict.get('src')
                if pict_url not in pUrlList:
                    pUrlList.append(pict_url)

            for iCol in range(ds['picMaxCnt']):
                if iCol < len(pUrlList):
                    if pUrlList[iCol] not in pList:
                        pList.append(pUrlList[iCol])
                else:
                    pList.append('')

            #print('pList=',pList[0:30])

            #出品者、スコア、カウント
            if shop_flg == 'Y':
                #出品者
                elements = soup.select('ul[role="button"] p[class="chakra-text css-naeo47"]')
                s_name = elements[0].contents[0]

                #スコア
                elements = soup.select('a[class="chakra-link css-19p30tk"] > div[class="chakra-stack css-g9cw6v"] svg')
                for s_score in elements:
                    if 'css-1x7bnhf' in s_score.attrs['class']:
                        s_score_d = s_score_d + 1
                    if 'css-1ozvvh' in s_score.attrs['class']:
                        s_score_p = s_score_p + 1
                
                s_score = s_score_p
            
                #カウント       
                elements = soup.select('ul[role="button"] p[class="chakra-text css-95dobi"]')
                s_count = elements[0].contents[0]

            else:
                others = soup.select('#item-info mer-user-object')
                s_name = others[0].attrs['name']
                s_score = others[0].attrs['score']
                s_count = others[0].attrs['count']

            #print('s_name=',s_name)
            #print('s_score=',s_score)
            #print('s_count=',s_count)

            #出品者除外
            if checkExclution(s_name):
                #webBrowserの解放
                bIdx[bNo] = True
                return None

            #評価数除外
            if ds['exScore'] != '' and ds['exScore'] is not None and s_score != '' and s_score is not None:
                if Decimal(s_score) < Decimal(ds['exScore']):
                    #webBrowserの解放
                    bIdx[bNo] = True
                    return None
            
            #コメント数除外
            if ds['exCount'] != '' and ds['exCount'] is not None and s_count != '' and s_count is not None:
                if int(s_count) < int(ds['exCount']):
                    #webBrowserの解放
                    bIdx[bNo] = True
                    return None
        
            #価格
            if shop_flg == 'Y':
                elements = soup.select('.chakra-text')
                for element in elements:
                    if 'css-1ttq47g' in element.attrs['class']:
                        item_price = '\\' + element.contents[0]
                        break
                    if 'css-1vczxwq' in element.attrs['class']:
                        item_price = '\\' + element.contents[0]
                        break
            else:
                price_elem = soup.select('#item-info div[data-testid="price"] span:nth-child(1)')
                item_cur = price_elem[0].contents[0]
                price_elem = soup.select('#item-info div[data-testid="price"] span:nth-child(2)')
                item_price = item_cur + price_elem[0].contents[0]

            #print('item_price=',item_price)

            #プロフィール
            if ds['isExProf'] == 'Y' and len(ds['exclutionList']) > 0:

                #プロフィールURL
                #2023/03/03
                if shop_flg == 'Y':
                    #profile = soup.select('.css-2lzsxm a[class="chakra-link css-19p30tk"]')
                    profile = soup.select('ul[role="button"] .chakra-link')
                    w_url = profile[0].get('href')
                    profile_url = ds['merShopUrl'] + w_url
                else:
                    profile = soup.select('#item-info a[data-location="item_details:seller_info"]')
                    w_url = profile[0].get('href')
                    profile_url = ds['merUrl'] + w_url

                #print('profile_url=',profile_url)

                # プロフィールへアクセス
                browser[iPar].get(profile_url)

                for j in range(30):
                    try:
                        #print('iAllRow=', str(iAllRow),'j=',j)

                        #変数初期化
                        profile_dtl = ''

                        # HTMLを文字コードをUTF-8に変換してから取得します。
                        html = browser[iPar].page_source.encode('utf-8')

                        # BeautifulSoupで扱えるようにパースします
                        soup = BeautifulSoup(html, "html.parser")

                        #プロフィール詳細
                        if shop_flg == 'Y':
                            profile = soup.select('p[class="chakra-text css-h7jfmu"]')
                            profile_dtl = profile[0].contents[0]
                        else:
                            #profile = soup.select('#main > div.sc-d9c03bcb-0.eotRGA > section > div > div.sc-d3e72d5f-2.gVXSKx > mer-show-more > mer-text')
                            profile = soup.select('#main mer-show-more')
                            profile_dtl = profile[0].contents[0].contents[0]
                                                
                        #print('profile_dtl=',profile_dtl[0:30])

                        #プロフィール詳細除外
                        if checkExclution(profile_dtl):
                            #webBrowserの解放
                            bIdx[bNo] = True
                            return None
                    
                    except Exception as e:
                        time.sleep(0.5)
                        if j == 29:
                            #print('')
                            #print('プロフィールを取得できませんでした。','profile_url=',profile_url)
                            #webBrowserの解放
                            bIdx[bNo] = True
                            return None
                        elif j > 0 and j%15 == 0:
                            browser[iPar].get(profile_url)
                    else:
                        break

        except Exception as e:
            time.sleep(0.5)
            if i == 29:
                print('')
                print('商品情報を取得できませんでした。','item_url=',item_url)
                #webBrowserの解放
                bIdx[bNo] = True
                return None
            elif i > 0 and i%15 == 0:
                browser[iPar].get(item_url)
        else:
            break

    #商品リストに追加
    itemInfo = [item_url]
    itemInfo.extend(pList)
    itemInfo.append(s_name)
    itemInfo.append(s_score)
    itemInfo.append(s_count)
    itemInfo.append(item_price)
    #非表示情報
    itemInfo.append(item_name)
    itemInfo.append(item_detail)
    itemInfo.append(profile_dtl)
    itemInfo.append(len(pList))

    #webBrowserの解放
    bIdx[bNo] = True

    return itemInfo

## リストを取得する ##
def list_get(s_word,s_page):

    #メルカリサイトにアクセス
    p = int(s_page) - 1
    if int(s_page) == 1:
        url = ds['merUrl'] + '/search?keyword=' + s_word + '&shipping_payer_id=2&status=on_sale' + ds['itemCond']
    else:
        url = ds['merUrl'] + '/search?keyword=' + s_word + '&shipping_payer_id=2&status=on_sale' + ds['itemCond'] + '&page_token=v1:' + str(p)

    if ds['minPrice'] != '' and ds['minPrice'] is not None:
        url = url + '&price_min=' + str(int(ds['minPrice']))

    if ds['maxPrice'] != '' and ds['maxPrice'] is not None:
        url = url + '&price_max=' + str(int(ds['maxPrice']))

    browser[0].get(url)

    # 2秒待機
    time.sleep(2)

    for i in range(100):
        try:
            #商品リストを取得
            #elements = browser[0].find_elements_by_css_selector('#item-grid li')

            # HTMLを文字コードをUTF-8に変換してから取得します。
            html = browser[0].page_source.encode('utf-8')

            # BeautifulSoupで扱えるようにパースします
            soup = BeautifulSoup(html, "html.parser")

            itemLists = soup.select('#item-grid li')

            if len(itemLists) == 0:
                rs = {}
                rs['iRow'] = 0
                rs['maxCol'] = 0
                rs['itemList'] = None
                rs['sWord'] = s_word
                rs['sPage'] = s_page
                return rs

            #レコード数
            maxRow = len(itemLists)

            iRow = 0
            iAllRow = 0
            futureList = []
            itemList = []
            itemKey = []
            itemKeyList = []

            for item in itemLists:
                
                #商品名
                #item_tag = element.find_element_by_css_selector('mer-item-thumbnail')
                #item_name = str(item_tag.get_attribute('item-name'))
                element = item.select('span[data-testid="thumbnail-item-name"]')
                item_name = element[0].contents[0]

                #商品名除外
                if checkExclution(item_name):
                    iAllRow += 1
                    continue
                
                #リンク
                #a_tag = element.find_element_by_css_selector('a')
                #item_url = str(a_tag.get_attribute('href'))
                element = item.select('a[data-location="search_result:best_match:body:item_list:item_thumbnail"]')
                item_url = element[0].attrs['href']
                
                if ds['merShopUrl'] not in item_url:
                    item_url = ds['merUrl'] + item_url 
                
                itemKey = [item_name,item_url]
                itemKeyList.append(itemKey)

        except Exception as e:
            time.sleep(0.5)
            if i == 99:
                print('')
                print('一覧を取得できませんでした。','url=',url)
                continue
            elif i > 0 and i%30 == 0:
                print('')
                print('一覧を再取得します。','url=',url)
                browser[0].get(url)
        else:
            break
        
    #webBrowserのインデックスを初期化
    bIdx = []
    for i in range(ds['pararel']):
        bIdx.append(True)

    # 商品情報の取得並列実行（max_workers が最大の並列実行数）
    futureList = []
    with futures.ThreadPoolExecutor(max_workers=ds['pararel']) as executor:
        for itemKey in itemKeyList:
            iAllRow += 1
            future = executor.submit(itemGetFunc,itemKey,iAllRow,maxRow,bIdx)
            futureList.append(future)

            # テスト用
            # if iAllRow >= 20: #len(elements):
            #     break

    for x in futureList:
        if x.result() is not None:
            itemList.append(x.result())
            iRow += 1

    print(' 抽出数 : ' + str(iRow) + ' 処理件数 : ' + str(iAllRow) + ' 全数 : ' + str(maxRow))
    rs = {}
    rs['iRow'] = iRow
    if len(itemList) > 0:
        rs['maxCol'] = max([item[18] for item in itemList])
    else:
        rs['maxCol'] = 0
    rs['itemList'] = itemList
    rs['sWord'] = s_word
    rs['sPage'] = s_page

    return rs

## エクセルから呼び出す場合 ##
def exSearch():
    main('X')

## メイン処理 ##
def main(mode):
    try:
        #イニシャル処理
        initial_set(mode)

        if ds['mode'] == 'X':
            ds['wb'].sheets('メルカリリスト').range('B2').value = 'データ取得中...'

        #GoogleChromeを起動
        #startGC(1 + ds['pararel'])

        # 2秒待機
        #time.sleep(2)

        #elapsed_time = 0
        
        #検索ワードごとのループ
        for seach in ds['swList']:

            #GoogleChromeを起動
            startGC(1 + ds['pararel'])

            # 2秒待機
            time.sleep(2)

            elapsed_time = 0
            
            sItemList = []
            sRow = 0
            sMaxCol = 0
            
            #ページごとのループ
            for pg in range(int(seach['pg'])):

                startTime = time.time()
                s_word = seach['sw']
                s_page = str(pg + 1)
                
                #エクセルモードでは、対象頁のみを検索
                if ds['mode'] == 'X':
                    if int(s_page) != int(seach['pg']):
                       continue

                print('')
                print(str(int(seach['no'])).zfill(3),'_',s_word,'-',s_page.zfill(3))

                print(' メルカリデータ取得中...')

                #リストを取得する
                rs = list_get(s_word, s_page)
            
                #リスト取得時間
                elapsed_time = time.time() - startTime
                print(' 取得完了 / ',str(int(elapsed_time/60)).zfill(2) + ':' + str(int(elapsed_time%60)).zfill(2))

                #ページごとにブックに保存
                rs['sNo'] = int(seach['no'])
                if rs['itemList'] is None:
                    print(' 検索結果がありません')
                    break
                elif len(rs['itemList']) > 0:
                    if ds['pgDiv'] == 'Y':
                        bookCreate(rs)
                        elapsed_time = time.time() - startTime
                        print(' ファイル保存完了 / ',str(int(elapsed_time/60)).zfill(2) + ':' + str(int(elapsed_time%60)).zfill(2))
                else:
                    print(' 抽出対象データがありません')

                sRow = sRow + int(rs['iRow'])
                if sMaxCol < int(rs['maxCol']):
                    sMaxCol = int(rs['maxCol'])
                sItemList.extend(rs['itemList'])

            #検索ワードごとにブックに保存
            if len(sItemList) > 0:
                if ds['pgDiv'] != 'Y':
                    rs['iRow'] = sRow
                    rs['maxCol'] = sMaxCol
                    rs['itemList'] = sItemList
                    bookCreate(rs)
                    elapsed_time = time.time() - startTime
                    print(' ファイル保存完了 / ',str(int(elapsed_time/60)).zfill(2) + ':' + str(int(elapsed_time%60)).zfill(2))
            else:
                print(' 抽出対象データがありません')

            #GoogleChromeをクローズ
            closeGC(1 + ds['pararel'])
            # 2秒待機
            time.sleep(2)

            gc.collect() 
        
    except Exception as e:
        print('')
        print(str(e.args))
        print(traceback.format_exc())
        ds['wb'].sheets('メルカリリスト').range('B2').value = str(e.args) + traceback.format_exc()
        ds['wb'].save()

    except:
        print('')
        print('処理を終了しました')

    finally:
        #エクセルファイルを閉じる
        if ds['mode'] == 'B':
            ds['wb'].close()
            ds['App'].quit()
            print("処理完了")
        else:
            ds['wb'].sheets('メルカリリスト').range('A2').value = str(int(elapsed_time/60)).zfill(2) + ':' + str(int(elapsed_time%60)).zfill(2)
            ds['wb'].sheets('メルカリリスト').range('B2').value = 'メルカリデータ取得完了'
            ds['wb'].save()

        #GoogleChromeをクローズ
        closeGC(1 + ds['pararel'])


if __name__ == '__main__':
    main('B')