# -*- coding: utf-8 -*-
# @Time : 2021-02-26 8:31 p.m.
# @Author : Weihao Sun
# @File : Crawler.py
# @Software: PyCharm

import ssl
import re
from bs4 import BeautifulSoup
import urllib.request, urllib.response
import xlwt
import sqlite3


ssl._create_default_https_context = ssl._create_unverified_context

findlink = re.compile(r'<a href="(.*?)">')  # regular of movie info link
findImageSrc = re.compile(r'img.*src="(.*?)"', re.S)    # regular of cover image link
findTitle = re.compile(r'<span class="title">(.*)</span>')  # regular of movie title
findRating = re.compile(r'<span class="rating_num" property="v:average">(.*)</span>')   # regular of rating
findRatingNum = re.compile(r'<span>(\d*)人评价</span>')    # regular of number of rating people
findIntro = re.compile(r'<span class="inq">(.*)</span>')     # regular of movie introduction
findInfo = re.compile(r'<p class="">(.*?)</p>', re.S)   # regular of movie information


# Main function
def main():
    url = "https://movie.douban.com/top250?start="
    dataList = getData(url)
    savePathxls = "DoubanMovieTop250.xls"
    savePathDb = "DoubanMovieTop250.db"
    saveDataXls(dataList, savePathxls)
    saveDataDb(dataList, savePathDb)


def getData(url):
    dataList = []

    for i in range(0, 10):
        pageUrl = url + str(i * 25)     # url of each page
        htmlSrc = getSrc(pageUrl)       # get html source code
        decoder = BeautifulSoup(htmlSrc, "html.parser")

        # Analyze each data that has been found
        for item in decoder.find_all("div", class_ = "item"):
            # print(item)
            data = []   # save information of one movie
            item = str(item)

            # Find info link of one movie
            link = re.findall(findlink, item)[0]
            data.append(link)

            # Find the cover image source of one movie
            imgSrc = re.findall(findImageSrc, item)[0]
            data.append(imgSrc)

            # Find official Chinese name and one another name in other language (if there is) of one movie
            titles = re.findall(findTitle, item)
            if (len(titles) == 2):
                # Find Chinese title
                cTitle = titles[0]
                data.append(cTitle)
                # Find title in other language
                oTitle = titles[1].replace("/", "")
                data.append(oTitle)
            else : # No title in other language
                data.append(titles[0])
                data.append("")

            # Find rating score of one movie
            rating = re.findall(findRating, item)[0]
            data.append(rating)

            # Find number of rating people of one movie
            ratingNum = re.findall(findRatingNum, item)[0]
            data.append(ratingNum)

            # Find brief intro of one movie
            intro = re.findall(findIntro, item)
            if (len(intro) != 0):
                intro = intro[0].replace("。", "")
                data.append(intro)
            else :
                data.append("")

            # Find info of one movie
            info = re.findall(findInfo, item)[0]
            info = re.sub('<br(\s+)?/>(\s+)?', " ", info)
            info = re.sub('/', " ", info)
            data.append(info)

            dataList.append(data)
    return dataList

# Get the html source code of a url
def getSrc(url):
    # Get a web browser header
    header = {"User-Agent":"Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_3) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.182 Safari/537.36"}
    # Request object to access the url
    request = urllib.request.Request(url, headers = header)
    html = ""
    try:
        # Initialize a response to for the request, decode the html code
        response = urllib.request.urlopen(request)
        html = response.read().decode("utf-8")
        # print(html)
    except urllib.error.URLError as e:
        # Print error code
        if hasattr(e, "code"):
            print(e.code)
        # Print error reason
        if hasattr(e, "reason"):
            print(e.reason)
    return html


# Save the data as a .xls Excel file
def saveDataXls(datalist, path):
    workbook = xlwt.Workbook(encoding="utf-8", style_compression=0)
    sheet = workbook.add_sheet("Top 250", cell_overwrite_ok=True)

    # Write column names
    cols = ("电影详情链接 Detail_Link", "图片链接 Image_Link", "影片中文名 Official_Chinese_Title",
           "影片外文名 Other_Language_Title", "评分 Rating", "评价人数 Number_Ratings",
           "概况 Brief_Intro", "相关信息 Information")
    for i in range(0, 8):
        sheet.write(0, i, cols[i])

    # Write data in each cell
    for i in range(0, 250):
        data = datalist[i]
        for j in range(0, 8):
            sheet.write(i+1, j, data[j])
    workbook.save(path)
    print("Save excel successfully")


# Save data as database
def saveDataDb(datalist, path):
    # Initiate sql table
    sql = '''
        create table top250
        (id integer primary key autoincrement,
        info_link text,
        img_link text,
        cname varchar,
        oname varchar,
        rating numeric,
        ratingNum numeric,
        intro text,
        info text)'''
    connect = sqlite3.connect(path)
    cursor = connect.cursor()
    cursor.execute(sql)
    connect.commit()

    # Insert data into the table
    for data in datalist:
        for index in range(len(data)):
            if index == 4 or index == 5:
                continue
            data[index] = '"' + data[index] + '"'
        writeSql = '''
            insert into top250(info_link, img_link, cname, oname, rating, ratingNum, intro, info) 
            values(%s)'''%",".join(data)
        cursor.execute(writeSql)
        connect.commit()
    cursor.close()
    connect.close()


if __name__ == "__main__":
    main()
    print("Crawling Over")
