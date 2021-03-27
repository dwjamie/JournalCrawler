from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, NoSuchFrameException, TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
from itertools import product
import pandas as pd
# from time import sleep
# from random import uniform


def get_articles(start_date, end_date, largest_index=30):
    start_year, start_month = start_date.split('-')
    end_year, end_month = end_date.split('-')

    start_year = int(start_year)
    start_month = int(start_month)
    end_year = int(end_year)
    end_month = int(end_month)

    articles_ = []
    n_years = end_year - start_year + 1
    indexes = [str(index).zfill(3) for index in range(largest_index + 1)]
    if n_years == 1:
        months = [str(month).zfill(2) for month in range(start_month, end_month + 1)]
        for month, index in product(months, indexes):
            articles_.append(str(start_year) + month + index)
    elif n_years == 2:
        months_1 = [str(month).zfill(2) for month in range(start_month, 13)]
        months_2 = [str(month).zfill(2) for month in range(1, end_month + 1)]
        for month, index in product(months_1, indexes):
            articles_.append(str(start_year) + month + index)
        for month, index in product(months_2, indexes):
            articles_.append(str(end_year) + month + index)
    else:
        years = [str(year) for year in range(start_year + 1, end_year)]
        months_1 = [str(month).zfill(2) for month in range(start_month, 13)]
        months_2 = [str(month).zfill(2) for month in range(1, 13)]
        months_3 = [str(month).zfill(2) for month in range(1, end_month + 1)]
        for month, index in product(months_1, indexes):
            articles_.append(str(start_year) + month + index)
        for year, month, index in product(years, months_2, indexes):
            articles_.append(year + month + index)
        for month, index in product(months_3, indexes):
            articles_.append(str(end_year) + month + index)

    return articles_


def get_next_article_index(article_index_, articles_, state, patience=2):
    if state == 'success':
        next_index = article_index_ + 1
        if next_index < len(articles_):
            return next_index
        else:
            return None
    elif state == 'fail':
        article_ = articles_[article_index_]

        year = int(article_[:4])
        month = int(article_[4:6])
        index = int(article_[6:])

        if int(index) < patience:
            next_article = str(year) + str(month).zfill(2) + str(index + 1).zfill(3)
        elif int(index) == patience or month == 12:
            next_article = str(year + 1) + '01000'
        else:
            next_article = str(year) + str(month + 1).zfill(2) + '000'

        try:
            return articles_.index(next_article)
        except ValueError:
            return None
    else:
        raise ValueError


def get_article_info(article_):
    global driver, wait  # 将driver设为全局变量并在此引入，可以避免多次开关浏览器

    year = int(article_[:4])
    if 1994 <= year <= 1999:
        article_ = str(year - 1990) + article_[4:6] + '.' + article_[6:]

    url = 'https://kns8.cnki.net/KCMS/detail/detail.aspx?dbcode=CJFD&filename=TJYJ' + article_
    driver.get(url)

    # 获取标题，若标题为空则说明该页不存在
    try:
        title = driver.find_element_by_css_selector('h1').text
    except NoSuchElementException:
        return None

    # 若作者为空则跳过文章
    authors = driver.find_elements_by_css_selector('#authorpart a')
    if not authors:
        authors = driver.find_elements_by_css_selector('#authorpart span')
    if not authors:
        return pd.DataFrame()

    # 如有，删除作者的右上角数字
    try:
        authors = [author.text[:-len(author.find_element_by_css_selector('sup').text)] for author in authors]
    except NoSuchElementException:
        authors = [author.text for author in authors]
    authors = '; '.join(authors)

    # 获取地址
    departments = driver.find_elements_by_css_selector('a.author')
    if not departments:
        departments = driver.find_elements_by_css_selector('#authorpart+ h3 span')
    departments = [department.text for department in departments]
    if departments and departments[0][1] == '.':
        departments = [department[3:] for department in departments]
    departments = '; '.join(departments)

    # 获取期刊名
    try:
        journal_name = driver.find_element_by_css_selector('.top-tip a:nth-child(1)').text
    except NoSuchElementException:
        journal_name = None

    # 获取发表时间
    try:
        publish_time = driver.find_element_by_css_selector('.top-tip a+ a').text
    except NoSuchElementException:
        publish_time = None

    # 获取页数
    try:
        n_page = driver.find_element_by_css_selector('.total-inform span:nth-child(3)').text
        if '页数：' in n_page:
            n_page = int(n_page.replace('页数：', ''))
        else:
            n_page = None
    except NoSuchElementException:
        n_page = None

    # 获取页码
    try:
        pages = driver.find_element_by_css_selector('.total-inform span:nth-child(2)').text
        if '页码：' in pages:
            pages = pages.replace('页码：', '')
        else:
            if '页数：' in pages:
                n_page = int(pages.replace('页数：', ''))
            pages = None
    except NoSuchElementException:
        pages = None

    # 获取摘要
    try:
        abstract = driver.find_element_by_css_selector('#ChDivSummary').text
    except NoSuchElementException:
        abstract = None

    # 获取关键词
    keywords = driver.find_elements_by_css_selector('.keywords a')
    keywords = '; '.join([keyword.text.replace(';', '') for keyword in keywords])

    # 获取基金资助
    funds = driver.find_elements_by_css_selector('.funds a')
    funds = '; '.join([fund.text.replace('；', '') for fund in funds])

    # 获取DOI
    try:
        DOI = driver.find_element_by_css_selector('.top-space:nth-child(1) p').text
    except NoSuchElementException:
        DOI = None

    # 获取专辑
    try:
        album = driver.find_element_by_css_selector('.top-space:nth-child(2) p').text
    except NoSuchElementException:
        album = None

    # 获取专题
    try:
        theme = driver.find_element_by_css_selector('.top-space:nth-child(3) p').text
    except NoSuchElementException:
        theme = None

    # 获取分类号
    try:
        category = driver.find_element_by_css_selector('.top-space:nth-child(4) p').text
    except NoSuchElementException:
        category = None

    # 获取下载数
    try:
        n_download = driver.find_element_by_css_selector('#DownLoadParts span:nth-child(1)').text
        n_download = int(n_download.replace('下载：', ''))
    except NoSuchElementException:
        n_download = None

    # 获取被引数
    while True:
        try:
            n_cited = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '#rc3'))).text
            n_cited = int(n_cited[1:-1])
            break
        except TimeoutException:
            driver.refresh()
        except (NoSuchElementException, ValueError):
            n_cited = None
            break

    # 获取引用数
    while True:
        try:
            n_cite = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '#rc1'))).text
            n_cite = int(n_cite[1:-1])
            break
        except TimeoutException:
            driver.refresh()
        except (NoSuchElementException, ValueError):
            n_cite = None
            break

    # 获取参考文献
    references = []
    if n_cite and n_cite > 0:
        while True:
            try:
                driver.switch_to.frame('frame1')
                essay_boxes = wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, 'essayBox')))
                for i in range(len(essay_boxes)):
                    while True:
                        additional_references = essay_boxes[i].find_elements_by_css_selector('li')
                        additional_references = [
                            additional_reference.text[4:] for additional_reference in additional_references
                        ]
                        references.extend(additional_references)

                        try:
                            next_page = essay_boxes[i].find_element_by_link_text('下一页')
                            driver.execute_script("arguments[0].click();", next_page)
                            essay_boxes = wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, 'essayBox')))
                        except NoSuchElementException:
                            break
                break
            except TimeoutException:
                driver.refresh()
            except NoSuchFrameException:
                break
    references = ';; '.join(references)

    # 生成字典
    article_info_ = {
        '标题': title,
        '作者': authors,
        '地址': departments,
        '期刊名': journal_name,
        '发表时间': publish_time,
        '页码': pages,
        '页数': n_page,
        '摘要': abstract,
        '关键词': keywords,
        '基金资助': funds,
        'DOI': DOI,
        '专辑': album,
        '专题': theme,
        '分类号': category,
        '下载数': n_download,
        '被引数': n_cited,
        '引用数': n_cite,
        '参考文献': references,
        '网址': url
    }

    # 由字典生成数据框
    article_info_ = pd.DataFrame(article_info_, index=[article_])

    # 防止网页爬取速度过快
    # sleep(uniform(3, 5))

    return article_info_


if __name__ == '__main__':
    articles = get_articles(start_date='1984-01', end_date='1999-12')

    article_infos = pd.DataFrame()
    driver = webdriver.Chrome(ChromeDriverManager().install())
    wait = WebDriverWait(driver, 20)

    article_index = 0
    while True:
        article = articles[article_index]
        article_info = get_article_info(article)
        if article_info is not None:
            article_infos = article_infos.append(article_info)
            article_infos.to_csv('统计研究_测试.csv', index=False, encoding='utf_8_sig')  # 每爬一篇存一次，虽然开销更大，但可以实时观察爬取情况
            article_index = get_next_article_index(article_index, articles, 'success')
        else:
            article_index = get_next_article_index(article_index, articles, 'fail')
        if not article_index:
            break

    driver.quit()

    # article_infos.to_csv('统计研究_测试.csv', index=False, encoding='utf_8_sig')
