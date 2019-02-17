import requests
from bs4 import BeautifulSoup
import urllib
from retry import retry
import re
import pandas as pd
from email import message
import smtplib
import os

uri = 'https://toyokeizai.net/'
category = 'category/weeklyranking'

url = uri+category
latest_post_date = 0

@retry(urllib.error.HTTPError, tries=5, delay=2, backoff=2)
def soup_url(url):
    print("...get...html...")
    htmltext = requests.get(url).text
    soup = BeautifulSoup(htmltext, "lxml")
    return soup

#Gmailの認証データ
smtp_host = os.environ["smtp_host"]
smtp_port = os.environ["smtp_port"]
from_email = os.environ["from_email"] # 送信元のアドレス
to_email = os.environ["to_email"]  # 送りたい先のアドレス 追加時は,で追加
bcc_email = os.environ["bcc_email"]  #Bccのアドレス追加
username = os.environ["username"] # Gmailのアドレス
password = os.environ["password"] # Gmailのパスワード


#取得したページの情報から、最新の記事URLとリンクを抜き出す

@retry(urllib.error.HTTPError, tries=7, delay=1)
def get_latest_post(soup):
    latest_items = soup.find("div", id="latest-items")
    latest_post_url = latest_items.find_all("a", class_="link-box")[0].get("href")
    latest_post_date = latest_items.find_all("span", class_="date")[0].string
    return latest_post_url , latest_post_date


#取得したページの情報から、必要なデータを抜き出す

@retry(urllib.error.HTTPError, tries=7, delay=1)
def get_ranking(soup):
    df = pd.DataFrame(index=[], columns=["ranking", "title", "author", "publisher"])
    for el in soup.find_all("tr"):
        rank  = el.find("th", class_="data1")
        if rank:
            rank = rank.string
        else:
            rank = el.find("th", class_="data5")
            if rank:
                rank = rank.string
            else:
                rank = "not find"

        title  = el.find("td", class_="data2")
        if title:
            title = title.string
        else:
            title = el.find("td", class_="data6")
            if title:
                title = title.string
            else:
                title = "not find"

        author = el.find("td", class_="data3")
        if author:
            author = author.string
        else:
            author = el.find("td", class_="data7")
            if author:
                author = author.string
            else:
                author = "not find"

        publisher = el.find("td", class_="data4")
        if publisher:
            publisher = publisher.string
        else:
            publisher = el.find("td", class_="data8")
            if publisher:
                publisher = publisher.string
            else:
                publisher = "not find"

        print("{} {} {} {}".format(rank, title, author, publisher))
        series = pd.Series([rank, title, author, publisher], index = df.columns)

        if series["ranking"] != "not find":
            df = df.append(series, ignore_index = True)

    return df


def mail():
    # メールの内容を作成
    msg = message.EmailMessage()
    msg.set_content('東洋経済 Ranking') # メールの本文
    msg['Subject'] = '東洋経済 Ranking' + latest_post_date # 件名
    msg['From'] = from_email # メール送信元
    msg['To'] = to_email #メール送信先
    msg['Bcc'] = bcc_email #bcc送信先

    #添付ファイルを作成する。
    mine={'type':'text','subtype':'comma-separated-values'}
    attach_file={'name':'ToyoRankingBooks.csv','path':'./ToyoRankingBooks.csv'}
    file = open(attach_file['path'],'rb')
    file_read = file.read()
    msg.add_attachment(file_read, maintype=mine['type'],
    subtype=mine['subtype'],filename=attach_file['name'])
    file.close()

    # メールサーバーへアクセス
    server = smtplib.SMTP(smtp_host, smtp_port)
    server.ehlo()
    server.starttls()
    server.ehlo()
    server.login(username, password)
    server.send_message(msg)
    server.quit()

#一連の実行関数

def main():
    wr_html = soup_url(url)
    latest_post_url, latest_post_date = get_latest_post(wr_html)
    ranking_url = uri + latest_post_url +"?page=2"
    ranking_soup = soup_url(ranking_url)
    TK_ranking_df = get_ranking(ranking_soup)
    print(TK_ranking_df)

    with open("ToyoRankingBooks.csv",mode="w",encoding="cp932",errors="ignore")as f:
        TK_ranking_df.to_csv(f)
    mail()

    with open("ToyoRankingBooks.csv",mode="w",encoding="utf-8",errors="ignore")as f:
        TK_ranking_df.to_csv(f)
    mail()

if __name__ == '__main__':
    main()
