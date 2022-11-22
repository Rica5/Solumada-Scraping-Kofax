/*
    Asynch sleep method
*/

async function sleep(s){
    await new Promise(resolve => setTimeout(resolve, s));
}
//finaliser
module.exports = sleep;