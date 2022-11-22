const ObjectsToCsv = require('objects-to-csv');
const convertCsvToXlsx = require('@aternus/csv-to-xlsx');
const fs = require('fs')
const path = require('path');
//finaliser
/* Save data to csv */
async function saveToCsv(data,filename){
  const csv = new ObjectsToCsv(data);
  let destination = `./public/assets/${filename}.csv`
  await csv.toDisk(destination)            // save the csv to disk
    .then(console.log('[I] CSV saved '))   // log success
    .catch((err)=>console.log('csv write error :: '+err))
}

/* Convert the actual csv to xls*/
async function csvToXls(filename){
  let source = path.join(`./public/assets/${filename}.csv`)
  let destination = path.join(`./public/assets/${filename}.xls`)
  
  try {
    convertCsvToXlsx(source, destination);
    } catch (e) {
    console.error(e.toString());
  } 
}

/* try to delete last data */
async function freeBtachFile(){
  try{
    fs.unlinkSync(`./public/assets/batch.csv`);   // delete csv
    fs.unlinkSync(`./public/assets/batch.xls`);   // delete xls
    console.log('[fs] - Batch file cleaned')
  }catch(e){
    console.log('[fs] - unlinck failed '+e)
}
}

module.exports = {saveToCsv,csvToXls,freeBtachFile};