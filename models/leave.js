const mongoose = require('mongoose');
const Leave = mongoose.Schema({
    m_code:String,
    num_agent:String,
    nom:String,
    date_start:String,
    date_end:String,
    duration:Number,
    type:String,
    status:String,
    rest:Number,
    validation:Boolean
})
module.exports = mongoose.model('cleave',Leave);