const mongoose = require('mongoose');
const Rem = mongoose.Schema({
   m_code:String,
   date:String, 
   remarques:String
})
module.exports = mongoose.model('remarque',Rem);