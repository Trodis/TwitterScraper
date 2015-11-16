# -*- coding: utf-8 -*-
from twython import Twython, TwythonError
from openpyxl import load_workbook

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
EXCELFILE = 'twitter_scraper.xlsx'
SHEET = 'Sheet1'

def getExcelSheet():
    wb = load_workbook(EXCELFILE)
    ws = wb[SHEET]
    return ws

def writeExcelSheet():
    pass

def saveExcelSheet():
    pass

def getMaxID(response):
    maxId = response[META][NEXTRESULT].split('&')[0].split('?max_id=')[1]
    return maxId

def parseTweetStatuses(response, user_id_list):
    for tweet in response[STATUSES]:
        if tweet[USER][USERID] not in user_id_list and tweet[USER][USERURL]:
            user_id_list.append(tweet[USER][USERID])
            writeExcelSheet(tweet[USER], tweet[USER][USERID], tweet[MESSAGE],
                    tweet[USER][USERURL])
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
    mainScraping(tweetobj, keyword='baby', limit=5)

if __name__ == "__main__":
    main()
