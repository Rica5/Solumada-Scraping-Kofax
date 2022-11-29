const mongoose = require('mongoose');

const Absent = mongoose.Schema({
    m_code:String,
    num_agent:String,
    nom:String,
    date:String,
    reason:String,
    time_start:String,
    return:String,
    validation:Boolean,
    status:String,
})
module.exports = mongoose.model('cabsent',Absent);