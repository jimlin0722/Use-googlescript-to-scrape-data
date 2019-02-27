# Use-googlescript-to-scrape-data
A practice to crawl stock price form [TWSE](http://www.twse.com.tw/zh/) and use google sheet as a mini database  

[Google Apps Script](https://developers.google.com/apps-script/) is a useful tool to extend G-Suite, base on javascript language  

#### Just a practice of web-scrape...
#### Some little challenge of this practice:
- Need header to access the json file in the [web](http://mis.twse.com.tw/stock/fibest.jsp?stock=2330) (example by TSMC)
- Need to scrape data frequently, since the TWSE update the price and volume every 5 sec  
#### Result and explanations
- Set cell 'A2' as our target stock, use function 'update' can get the latest price and volume
<br></br>
![update](https://github.com/jimlin0722/Use-googlescript-to-scrape-data/blob/master/images/update.gif)
<br></br>
- Use functions 'createTriggers' and 'deleteTrigger' to contral the scraper. Make function 'update' runs every 1 minute during trading hours(9:00 - 13:30)
<br></br>
![data](https://github.com/jimlin0722/Use-googlescript-to-scrape-data/blob/master/images/data.gif)
<br></br>
![graph](https://github.com/jimlin0722/Use-googlescript-to-scrape-data/blob/master/images/graph.png)
<br></br>
- Set function 'clean_table' run every day to delete the data out of trading hours

So far, I can only scrape data every 1 minute due to the quota of execution time([see document](https://developers.google.com/apps-script/guides/services/quotas))
