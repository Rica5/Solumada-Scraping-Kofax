const mongoose = require('mongoose');

const Status = mongoose.Schema({
    m_code:String,
    num_agent:String,
    date:String,
    time_start: String,
    time_end:String,
    nom:String,
    worktime:Number,
    locaux:String,
})
module.exports = mongoose.model('cstatus',Status);