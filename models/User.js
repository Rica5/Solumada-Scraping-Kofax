const mongoose = require('mongoose');

const User = mongoose.Schema({
   username:String,
   last_name:String,
   first_name:String,
   password:String,
   m_code:String,
   num_agent:String,
   occupation:String,
   change:String,
   act_stat:String,
   act_loc:String,
   shift:String,
   late:String,
   count:Number,
   take_break:String,
   remaining_leave:Number,
   leave_taked:Number,
   leave_stat:String,
   save_at:String,
   sexe:String,
   situation:String,
   user_ht:Number,
   project:String,
   matr:String,
   usuel:String,
   situation:String,
   cin:String,
   adresse:String,
   sexe:String,
   cnaps_num:String,
   classification:String,
   contrat:String,
   date_fin:String
})
module.exports = mongoose.model('cuser',User);