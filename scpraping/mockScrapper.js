const puppeteer = require('puppeteer');
const {saveToCsv,csvToXls,freeBtachFile} = require('./file');
const {login} = require('./pageCheck')
const fs = require ('fs');

// DATA
var browser;    
var page;

const link1 = 'https://service.europe.arco.biz/ktmthinclient/Validation.aspx';  // standard link
const link2 = 'https://service.europe.arco.biz/ktmthinclient2/Validation.aspx'; // link used on error 

let link = link1                                                                // use standard link as default link
//finaliser
/* create browser */
async function createBrowser(){
    const browser = await puppeteer.launch({
        ignoreHTTPSErrors: true,
        headless: true,
        devtools: false,
        args: ['--no-sandbox', '--disable-setuid-sandbox']
    });
    console.log('[👍] browser .. ');
    return browser
}

/* create page from a browser */
async function createPage(browser){
    const page = await browser.newPage();
    console.log('[👍] new page created  ..');
    return page
}

/* Login at the first time */
async function firstLogin(){
    await page.goto(link)   // page navigate to the main page that is arco scraping when the website is opened for the first moment
        .then(async ()=>{
            console.log('[👍] Login page opened')
            await login(page)                                                               // fill credential for login
        })
        .catch((e)=>console.log("[i] Go to Login page error :: "+e))
}

/* RESTART browser */
const restartBrowser = async()=>{
    try{
        await browser.close()               // try to close last used browser
    }
    catch(e){                               // there is no browser to close
        console.log("[i] error on browser close :: browser already undefined")
    }
    finally{                                // in any case,     
        browser = await createBrowser()     // create a browser
        page = await createPage(browser)    // create a page
        await firstLogin(page)              // do the login
    }
}

/* DATA FETCH from arco-site */
async function fetchData(){

    // Seesion expiration check 
    // await sessionExpired(page)
    //     .then(async(expired)=>{
    //         if(expired===true){
    //             await restartBrowser()
    //             //await login(page)
    //             console.log("[👍] handle seesion expired")
    //         }
    //         else if(expired===false){
    //             console.log("[👍] session not expired")
    //         }
    //         else
    //             console.log('[x] Error handling promise resolve bool on session expired')
            
    //     })
    //     .catch((e)=>console.log("ERR catch sessionExpiored"))
    
    try{
        await page.goto(link)
        .then(()=>console.log("[👍] validation page opened"))
        .catch((e)=>console.log('Goto validation Fail page'))
        // Wait for table selector before scraping
        await page.waitForSelector('.x-grid3-row-table tr',{visible:true,timeout: 10000})
            .then(()=>console.log('Selector ok'))
            // error on timeout

    let rows = await page.evaluate(
            ()=> Array.from(window.document.querySelectorAll('.x-grid3-row-table tr'))
            .map((row,i)=>{
                let data = {
                    index : i+1,
                    batch : row.querySelector('div.x-grid3-col-name').innerText,
                    document : row.querySelector('div.x-grid3-col-6').innerText,
                    client : row.querySelector('div.x-grid3-col-batchType').innerText,
                    priority : row.querySelector('div.x-grid3-col-0').innerText,
                    date : row.querySelector('div.x-grid3-col-2').innerText,
                    status : row.querySelector('div.x-grid3-col-status').innerText
                }
                return data
            })
    )
            
    rows = rows.filter((e)=>e.status=="Ready")      // Record only ready data
    console.log("Total file scraped "+rows.length)  // Log the data length

    /* Saving file */
        await freeBtachFile()           // delete last batch file saved
        await saveToCsv(rows,'batch');  // await csv file before conversion
        csvToXls('batch');

        return (rows);
    }
    catch(e){                           // error unhandeled
        console.log("[i] Validation page not reached :: "+e)
        
        if(link==link1) link = link2    // change link
        else if(link==link2) link = link1
        else console.log('[i] Error handlink link change at catch error')
        await restartBrowser()
        return []                       // send empty array ass error mssg for client side    
    }
}

module.exports = {fetchData,restartBrowser}