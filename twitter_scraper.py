# -*- coding: utf-8 -*-
from twython import Twython, TwythonError
from openpyxl import load_workbook, Workbook

CONSUMER_KEY    = "oFnKOZ1a4BJMOMjCkJbb7rv2i"
CONSUMER_SECRET = "8V6V7w26vy0kUl99vNmZg3Fod8RLl1nLuxslDhh0T0BwhxN6mD"
TOKEN_KEY       = "93475883-sypM3QYTxvr6UI5OkC3LGFxG4PrbdjUnZNaoj7hOp"
TOKEN_SECRET    = "t1t0lLVsgS7Skxz1M5yVikfTvDTX8oZILbVoMqT2ubSDH"

USER = 'user'
USERID = 'id_str'
MESSAGE = 'text'
USERURL = 'url'
META = 'search_metadata'
NEXTRESULT = 'next_results'
STATUSES = 'statuses'
EXCELFILE = 'scraped_content.xlsx'
SHEET = 'Scraped Tweets'

def getExcelSheet():
    #wb = load_workbook(EXCELFILE)
    wb = Workbook()
    ws = wb.active
    ws.cell(row=1, column=1, value='User ID')
    ws.cell(row=1, column=2, value='Username')
    ws.cell(row=1, column=3, value='Tweet Message')
    ws.cell(row=1, column=4, value='Keyphrase')
    ws.cell(row=1, column=5, value='Data retrieved')
    ws.cell(row=1, column=6, value='Website from User')
    ws.cell(row=1, column=7, value='E-Mail from Website')
    ws.cell(row=1, column=8, value='Whois Mail')
    ws.cell(row=1, column=9, value='Link to Contact Form')
    wb.save(EXCELFILE)
    return ws

def writeExcelSheet(ws, user_id, user_name, message, keyword, url):
     
    pass

def saveExcelSheet():
    pass

def getMaxID(response):
    maxId = response[META][NEXTRESULT].split('&')[0].split('?max_id=')[1]
    return maxId

def testLimit(tweetobj):
    response = tweetobj.search(q='baby', count = 100)
    maxID = getMaxID(response)
    user_id_list = []
    while NEXTRESULT in response[META]:
        for tweet in response[STATUSES]:
            if tweet[USER][USERID] not in user_id_list:
                user_id_list.append(tweet[USER][USERID])
            else:
                continue
        print len(user_id_list)
        maxID = getMaxID(response)
        response = tweetobj.search(q='baby', max_id=maxID, count = 100)
        

def parseTweetStatuses(response, keyword):
    for tweet in response[STATUSES]:
        if tweet[USER][USERID] not in user_id_list and tweet[USER][USERURL]:
            user_id_list.append(tweet[USER][USERID])
            writeExcelSheet(tweet[USER][USERID], tweet[USER], tweet[MESSAGE],
                    keyword, tweet[USER][USERURL])
        else:
            continue
    return user_id_list

def mainScraping(tweetobj, keyword, limit=None):
    user_id_list = []
    if limit:
        response = tweetobj.search(q=keyword, count=limit)
        parseTweetStatuses(response)
    else:
        response = tweetobj.search(q=keyword)
        while NEXTRESULT in response[META]:
            maxId = getMaxID(response) 
            response = tweetobj.search(q=keyword, max_id=maxId)
            user_id_list = parseTweetStatuses(response, ids)

def main():
    tweetobj = Twython(CONSUMER_KEY, CONSUMER_SECRET, TOKEN_KEY, TOKEN_SECRET)
    #mainScraping(tweetobj, keyword='baby', limit=5)
    #testLimit(tweetobj)
    test = getExcelSheet()

if __name__ == "__main__":
    main()
