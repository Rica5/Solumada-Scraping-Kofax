# puppeteer-webscrap-app
Webscrap app with node.js and puppeteer deployed on heroku. 
<br>
## Site link
https://www.french-property.com/properties-for-sale?start_page=1

## puppeteer link
https://developers.google.com/web/tools/puppeteer

## Puppeteer builder pack for heroku
Your puppeteer app won't run if your isn't build with puppeteer-heroku-buildpack, <br>
Go to heroku > setting > builder-pack > add this link bellow <br>
https://github.com/jontewks/puppeteer-heroku-buildpack.git

## About release
### V1
Server runnin with heroku local <br>
Puppeteer open new browser and new page on every srap call <br>

### v2
Separated file for utile methode : login ()

### v3
Server running on heroku with puppeteer-heroku-buildpack

### v3-shared and v4
Handeling multiple scrap request

### v5
Handeling multiple request scrap <br>
App, puppeteer don't create new browser or page on every scrap <br>
Handle expired session

