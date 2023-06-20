# -*- coding: utf-8 -*-
import requests
import json
import pandas as pd

j=0

place_first_review_url = [
    "https://www.google.com/maps/preview/review/listentitiesreviews?authuser=0&hl=zh-TW&gl=tw&pb=!1m2!1y3776578326095458145!2y399780783196016697!2m1!2i" + str(j) + "!3e2!4m6!3b1!4b1!5b1!6b1!7b1!20b1!5m2!1sYQMlZOKzBYG7wAOij4fAAQ!7e81",
    "https://www.google.com/maps/preview/place?authuser=0&hl=zh-TW&gl=hk&pb=!1m14!1s0x3442ab63e0c770cb%3A0xed057ba13aed526b!3m9!1m3!1d4974.516711463821!2d121.59289!3d25.0541705!2m0!3m2!1i1103!2i1057!4f13.1!4m2!3d25.056133277412986!4d121.59866653382778!13m50!2m2!1i408!2i240!3m2!2i10!5b1!7m42!1m3!1e1!2b0!3e3!1m3!1e2!2b1!3e2!1m3!1e2!2b0!3e3!1m3!1e8!2b0!3e3!1m3!1e10!2b0!3e3!1m3!1e10!2b1!3e2!1m3!1e9!2b1!3e2!1m3!1e10!2b0!3e3!1m3!1e10!2b1!3e2!1m3!1e10!2b0!3e4!2b1!4b1!9b0!14m4!1s-zclZNbSM_iy2roPpqe48Ag!3b1!7e81!15i10555!15m48!1m8!4e2!18m5!3b0!6b0!14b1!17b1!20b1!20e2!2b1!4b1!5m6!2b1!3b1!5b1!6b1!7b1!10b1!10m1!8e3!11m1!3e1!17b1!20m2!1e3!1e6!24b1!25b1!26b1!29b1!30m1!2b1!36b1!43b1!52b1!55b1!56m2!1b1!3b1!65m5!3m4!1m3!1m2!1i224!2i298!107m2!1m1!1e1!22m1!1e81!29m0!30m3!3b1!6m1!2b1!32b1!37i640&q=%E4%B8%AD%E5%9C%8B%E9%9B%BB%E8%A6%96%E5%85%AC%E5%8F%B8&pf=t
]

for i in range(place_first_review_url):

    # url
    j=0
    
    while 1:
        url = place_first_review_url[i]
        j = j + 10
        
        # get the url
        review_text = requests.get(url).text
        # Replace special characters set for information security 取代掉為了資訊安全而設定的特殊字元
        pretext = ')]}\''
        review_text = review_text.replace(pretext,'')
        # replace string as json 把字串讀取成json
        soup = json.loads(review_text)
        
        # get the list that contain the reviews 取出包含留言的List 。
        conlist = soup[2]
        
        # create a empty dataframe建立一個空的dataframe
        df = pd.DataFrame(columns=["username", "time", "rate", "comment"])
        
        # Get the reviews one by one 逐筆抓出
        for i in conlist:
            username = i[0][1]
            time = i[1]
            rate = i[4]
            comment = i[3]
            df = df.append({"username": username, "time": time, "rate": rate, "comment": comment}, ignore_index=True)
        
        df = df.sort_values(by=['time'], ascending=True)
        
        # export to Excel file 輸出成excel檔案
        df.to_excel("N.Taichung_sort.xlsx", index=False)
