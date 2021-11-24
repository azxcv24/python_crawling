# pip install selenium
# pip install chromedriver-autoinstaller
# pip install bs4

from selenium import webdriver
import chromedriver_autoinstaller
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import xlsxwriter
import time
# 이미지 바이트 처리
from io import BytesIO
import requests
import math
from datetime import datetime, timedelta

# 다나와 사이트 검색

options = Options()
options.add_argument('headless');  # headless는 화면이나 페이지 이동을 표시하지 않고 동작하는 모드

# webdirver 설정(Chrome, Firefox 등)
chromedriver_autoinstaller.install()
driver = webdriver.Chrome(options=options)  # 브라우저 창 안보이기
#driver = webdriver.Chrome() # 브라우저 창 보이기
#상세페이지 드라이버
driver2 = webdriver.Chrome(options=options)  # 브라우저 창 안보이기
#driver2 = webdriver.Chrome() # 브라우저 창 보이기



# 크롬 브라우저 내부 대기 (암묵적 대기)
driver.implicitly_wait(5)
driver2.implicitly_wait(5)

# 브라우저 사이즈
driver.set_window_size(1920, 1280)
driver2.set_window_size(1920, 1280)


category = "RAM"
# 페이지 이동(열고 싶은 URL)
url1 ='http://prod.danawa.com/list/?cate=112752'
driver.get(url1)

# 페이지 내용(JSON형식으로 페이지 내용을 표시)
#print('Page Contents : {}'.format(driver.page_source))
# 제조사별 검색 (XPATH 경로 찾는 방법은 이미지 참조)
#mft_xpath = '//*[@id="dlMaker_simple"]/dd/div[2]/button[1]'
#WebDriverWait(driver1, 3).until(EC.presence_of_element_located((By.XPATH, mft_xpath))).click()
# 원하는 모델 카테고리 클릭 (XPATH 경로 찾는 방법은 이미지 참조)
#model_xpath = '//*[@id="selectMaker_simple_priceCompare_A"]/li[16]/label'
#WebDriverWait(driver1, 3).until(EC.presence_of_element_located((By.XPATH, model_xpath))).click()
# 2차 페이지 내용
# print('After Page Contents : {}'.format(driver.page_source))

# 검색 결과가 렌더링 될 때까지 잠시 대기
time.sleep(2)

# 현재 페이지
curPage = 1

# 크롤링할 전체 페이지수
totalPage = 10

#현재 날짜
timenow = datetime.today().strftime("%Y%m%d%H%M")

# Excel 처리 선언
workbook = xlsxwriter.Workbook('data/'+ category+'_crawling_result('+ timenow +').xlsx')
# 워크 시트
worksheet = workbook.add_worksheet()
# 엑셀 행 수
excel_row = 1
worksheet.set_column('A:A', 40)  # A 열의 너비를 40으로 설정
worksheet.set_row(0, 18)  # A열의 높이를 18로 설정
worksheet.set_column('B:B', 12)
worksheet.set_column('C:C', 12)
worksheet.set_column('D:D', 12)
worksheet.set_column('E:E', 40)
worksheet.set_column('F:F', 20)
worksheet.set_column('G:G', 12)
worksheet.set_column('H:H', 12)
worksheet.set_column('I:I', 20)
worksheet.write(0, 0, '제품 모델명')
worksheet.write(0, 1, '올림가격') ##
worksheet.write(0, 2, '크롤링한 최저가')
worksheet.write(0, 3, '할인율') ##
worksheet.write(0, 4, '스팩')
worksheet.write(0, 5, '프리뷰이미지')
worksheet.write(0, 6, '상세이미지')
worksheet.write(0, 7, '판매자')
worksheet.write(0, 8, '등록월')
worksheet.write(0, 9, '제품페이지')
worksheet.write(0, 10, '분류')


while curPage <= totalPage:
    # bs4 초기화
    soup = BeautifulSoup(driver.page_source, 'html.parser')

    # 상품 리스트 선택

    goods_list = soup.select('li.prod_item.prod_layer')
    #다나와 페이지중 마지막 쓸모없는 div제거
    goods_list.pop()

    # 페이지 번호 구분 출력
    print('----- Current Page : {}'.format(curPage), '------')
    for v in goods_list:
        # 상품명, 가격, 이미지 + (추가할것 : 등록일 , 각제품별 상세페이지로 들어가서 상세이미지, 제품 카테고리 써잇는거 모음, 가격에파는 판매자이름)
        #상품명
        name = v.select_one('p.prod_name > a').text.strip()
        #가격(정수로 변환,반올림으로 정가 추가)
        try:
            price1 = v.select_one('li.rank_one > p > a > strong').text.strip()
        except Exception as e:
            price1 = v.select_one('li > p > a > strong').text.strip()
        price1 = int(price1.replace(',',''))
        price0 = round(price1, -(int(math.log10(price1+60000)))) #TODO 올림으로 처리하고 싶은데 방법이 없어 보인다 지금은 반올림(반올림시 할인가격이 더 비싸는 현상 발생) -> 소수점이하로 보낸다음 올림해서 다시 정수로 가져오자!
        #판매자명(cpu와같이 몰이름이 아닌 사양일경우 그값을 가져오자!)
        try:
            mall = v.select_one('li.rank_one > div >p.memory_sect').text.strip()
        except Exception as e:
            mall = v.select_one('p.memory_sect').text.strip()
        #스팩
        spec_list = v.select_one('div.spec_list').text.strip()
        #등록일
        insert_date = v.select_one('dl.meta_item.mt_date').text.strip()
        #이미지링크
        img_link = v.select_one('div.thumb_image > a > img').get('data-original')
        if img_link == None:
            img_link = v.select_one('div.thumb_image > a > img').get('src')
        #이미지링크 처리
        imgLink = img_link.split("?")[0].split("//")[1]
        img_url = 'https://{}'.format(imgLink)

        # 이미지 요청 후 바이트 반환(다운할려고하는데 나중에)
        res = requests.get(img_url)  # 이미지 가져옴
        img_data = BytesIO(res.content)  # 이미지 파일 처리
        image_size = len(img_data.getvalue())  # 이미지 사이즈


        #상세이미지 가져오기!
        prod_url = v.select_one('p.prod_name > a').get('href')
        driver2.get(prod_url)


        soup2 = BeautifulSoup(driver2.page_source, 'html.parser')
        prod_page = soup2.select('div.detail_export > div.inner') #여기서 파싱할 페이지에서 해당 영역 지정

        img_url2 = ''

        for v2 in prod_page:
            time.sleep(5)
            img_url2 = v2.select_one('img').get('src') #하도 페이지마다 다르니까 img태그에서 가져온다!


        # 엑셀 저장(텍스트)
        worksheet.write(excel_row, 0, name)
        worksheet.write(excel_row, 1, price0)
        worksheet.write(excel_row, 2, price1)
        #worksheet.write(excel_row, 3, 만들어야함) #TODO 할인율계산하게
        worksheet.write(excel_row, 4, spec_list)
        # 엑셀 저장(이미지)
        if image_size > 0:  # 이미지가 있으면
            # worksheet.insert_image(excel_row, 2, img_url, {'image_data' : img_data})
            worksheet.write(excel_row, 5, img_url)  # image url 텍스트 저장
        worksheet.write(excel_row, 6, img_url2) #상세이미지링크
        worksheet.write(excel_row, 7, mall) #판매자
        worksheet.write(excel_row, 8, insert_date) #등록일
        worksheet.write(excel_row, 9, prod_url)
        worksheet.write(excel_row, 10, category)


        # 엑셀 행 증가
        excel_row += 1

        print(name, ', ', price0, ', ', price1, ', ',spec_list, ', ', img_url, ', ', mall , ', ', insert_date, ', ',prod_url, ', ',img_url2)
    print()



    # 페이지 수 증가
    curPage += 1

    if curPage > totalPage:
        print('Crawling succeed!')
        break

    # 페이지 이동 클릭
    cur_css = 'div.number_wrap > a:nth-child({})'.format(curPage)
    WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.CSS_SELECTOR, cur_css))).click()

    # BeautifulSoup 인스턴스 삭제
    del soup
    del soup2

    # 3초간 대기
    time.sleep(3)

# 브라우저 종료
driver.close()
driver2.close()

# 엑셀 파일 닫기
workbook.close() # 저장

#참조 코드 출처: https://link2me.tistory.com/2003