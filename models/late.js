const mongoose = require('mongoose');

const Late = mongoose.Schema({
    m_code:String,
    num_agent:String,
    nom:String,
    date:String,
    time:Number,
    reason:String,
    validation:Boolean
})
module.exports = mongoose.model('clate',Late);