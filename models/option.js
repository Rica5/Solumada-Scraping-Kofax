const mongoose = require('mongoose');
const Opt = mongoose.Schema({
   add_leave:String,
   paie_generated:String,
   list_paie:String
})
module.exports = mongoose.model('option',Opt);