const express = require("express");
const routeExp = express.Router();
const mongoose = require("mongoose");
const UserSchema = require("../models/User");
const StatusSchema = require("../models/status");
const LateSchema = require("../models/late");
const AbsentSchema = require("../models/absent");
const LeaveSchema = require("../models/leave");
const OptSchema = require("../models/option");
const RemSchema = require("../models/remarques");
const nodemailer = require("nodemailer");
const extra_fs = require('fs-extra');
const crypto = require('crypto');
const moment = require("moment");
const { PDFNet } = require('@pdftron/pdfnet-node');
const ExcelFile = require("sheetjs-style");
const fs = require('fs');
//Variables globales
var date_data = [];
var data = [];
var all_datas = [];
var num_file = 1;
var hours = 0;
var minutes = 0;
var notification = [];
var data_desired = {};
var monthly_leave = [];
var maternity = [];
var filtrage = {};
var exc_retard = ["RH", "MANAGER", "IT", "GERANT"];
var access = ["", "", ""];
var deduire = ["Mise a Pied", "Absent", "Congé sans solde"];
var ws_leave;
var ws_left;
var ws_individual;
var datestart_leave;
var dateend_leave;

//Mailing
var transporter = nodemailer.createTransport({
  service: "gmail",
  auth: {
    user: "ricardoramandimbisoa@gmail.com",
    pass: "dletiebgggunowhf",
  },
});
async function daily_restart(req) {
  mongoose
    .connect(
      "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
      {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
      }
    )
    .then(async () => {
      var now = moment().format("dddd");
      var opt_daily = await OptSchema.findOne({ _id: "636247a2c1f6301f15470344" });
      if (now != opt_daily.date_change) {
        await conge_define(req);
        await checkleave();
        await leave_permission();
        notification = [];
        await OptSchema.findOneAndUpdate({ _id: "636247a2c1f6301f15470344" }, { date_change: now });
      }
      else {
        console.log("Already done")
      }
    })
}
async function monthly_restart() {
  mongoose
    .connect(
      "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
      {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
      }
    )
    .then(async () => {
      var now = moment().format("MMMM");
      var opt_daily = await OptSchema.findOne({ _id: "636247a2c1f6301f15470344" });
      if (now != opt_daily.month_change) {
        await addin_leave();
        await OptSchema.findOneAndUpdate({ _id: "636247a2c1f6301f15470344" }, { month_change: now });
      }
      else {
        console.log("Leave already added");
      }
    })

}
async function addin_leave() {
  mongoose
    .connect(
      "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
      {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
      }
    )
    .then(async () => {
      var all_user = await UserSchema.find({});
      for (u = 0; u < all_user.length; u++) {
        if (all_user[u].shift != "SHIFT WEEKEND") {
          await UserSchema.findOneAndUpdate({ m_code: all_user[u].m_code }, { $inc: { leave_taked: 2.5 } });
        }
        else {
          await UserSchema.findOneAndUpdate({ m_code: all_user[u].m_code }, { $inc: { leave_taked: 0.75 } });
        }
      }
    })
}
//Page route
routeExp.route("/").get(async function (req, res) {
  session = req.session;
  if (session.occupation_u == "User") {
    res.redirect("/employee");
  } else if (session.occupation_a == "Admin") {
    res.redirect("/home");
  } else {
    await daily_restart(req);
    res.render("Login.html", { erreur: "" });
  }
});

//Post ip
routeExp.route("/getip").post(async function (req, res) {
  session = req.session;
  await set_ip(req.body.ip, req.session);
  res.send("Ok");
});
//Function set ip
async function set_ip(ip_get, session) {
  session.ip = ip_get;
}
//Login post
routeExp.route("/login").post(async function (req, res) {
  session = req.session;
  await login(req.body.username, req.body.pwd, req.session, res);
});
//Function login
async function login(username, pwd, session, res) {

  mongoose
    .connect(
      "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
      {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
      }
    )
    .then(async () => {
      try {
        let hash = crypto.createHash('md5').update(pwd.trim()).digest("hex");
        var logger = await UserSchema.findOne({
          username: username.trim(),
          password: hash,
        });
        if (logger) {
          //Tete
          if (logger.change != "n") {
            if (logger.occupation == "User") {
              session.occupation_u = "User";
              session.m_code = logger.m_code;
              session.shift = logger.shift;
              session.name = logger.first_name + " " + logger.last_name;
              session.num_agent = logger.num_agent;
              session.forget = "n";
              if (await StatusSchema.findOne({ m_code: session.m_code, time_end: "", date: { $ne: moment().format("YYYY-MM-DD") } }).sort({ _id: -1 })) {
                session.forget = JSON.stringify(await StatusSchema.findOne({ m_code: session.m_code, time_end: "", date: { $ne: moment().format("YYYY-MM-DD") } }));
                await UserSchema.findOneAndUpdate({ m_code: session.m_code }, { act_stat: "LEFTING", act_loc: "Not defined", late: "n", count: 0 });
              }
              if (logger.act_stat == "VACATION") {
                session.occupation_u = null;
                session.m_code = null;
                session.shift = null;
                res.render("Login.html", {
                  erreur: "Vous êtes en congé prenez votre temps",
                });
              }
              else {
                session.time = "y";
                var already_time = await StatusSchema.findOne({m_code:session.m_code,date:moment().format("YYYY-MM-DD")});
                var already_late = await LateSchema.findOne({m_code:session.m_code,date:moment().format("YYYY-MM-DD")});
                //Clean code retard
                if ((logger.late == "n" && already_time === null && already_late === null)) {
                  session.time = "y";
                  var start = "";
                  var today = moment().day();
                  switch (session.shift) {
                    case "SHIFT 1": start = "06:15"; break;
                    case "SHIFT 2": start = "12:15"; break;
                    case "SHIFT 3": start = "18:15"; break;
                    case "TL": start = "21:00"; break;
                    case "ENG": start = "09:00"; break;
                    case "IT": start = "21:00"; break;
                    case "RH": start = "21:00"; break;
                    case "MANAGER": start = "21:00"; break;
                    case "GERANT": start = "21:00"; break;
                    default: start = "08:00"; break;
                  }
                  switch (today) {
                    case 6: start = "08:00"; break;
                    case 7: start = "08:00"; break;
                  }
                  if (exc_retard.includes(session.shift)) {
                    start = "21:00";
                  }
                  var timestart = moment().add(3, 'hours').format("HH:mm");
                  var time = calcul_retard(start, timestart);
                  session.time = "y";
                  if (time > 10) {
                    session.time = time + " minutes";
                    var new_late = {
                      m_code: session.m_code,
                      num_agent: session.num_agent,
                      date: moment().format("YYYY-MM-DD"),
                      nom: session.name,
                      time: time,
                      reason: "",
                      validation: false
                    }
                    await LateSchema(new_late).save();
                    res.redirect("/employee");
                  }
                  else {
                    session.time = "y";
                    res.redirect("/employee");
                  }
                  await UserSchema.findOneAndUpdate({ m_code: session.m_code }, { late: "y" });
                }
                else {
                  res.redirect("/employee");
                }
              }
            } else if (logger.occupation == "Admin") {
              session.occupation_a = logger.occupation;
              filtrage = {};
              res.redirect("/home");
            }
            else {
              session.occupation_tl = "checker";
              res.redirect("/managementtl");
            }
          }
          else {
            session.mailconfirm = logger.username;
            res.render("change_password.html", {
              first: "y",
            });
          }
          //Pied
        } else {
          res.render("Login.html", {
            erreur: "Email ou mot de passe incorrect",
          });
        }
      } catch (error) {
        console.log(error)
        res.render("Login.html", {
          erreur: "Problème sur votre login, veuillez reessayez",
        });
      }

    });

}
//Page change password
routeExp.route("/changepassword").get(async function (req, res) {
  res.render("change_password.html", { first: "" });
});
//Check email
routeExp.route("/checkmail").post(async function (req, res) {
  session = req.session
  var email = req.body.email;
  mongoose
    .connect(
      "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
      {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
      }
    )
    .then(async () => {
      if (await UserSchema.findOne({ username: email })) {
        session.mailconfirm = email;
        session.code = randomCode();
        sendEmail(
          session.mailconfirm,
          "Code de verification",
          htmlVerification(session.code)
        );
        res.send("done");
      } else {
        res.send("error");
      }
    });
})
// Check code
routeExp.route("/checkcode").post(async function (req, res) {
  session = req.session;
  if (session.code == req.body.code) {
    res.send("match");
  } else {
    res.send("error");
  }
});
// update password
routeExp.route("/changepass").post(async function (req, res) {
  session = req.session
  var newpass = req.body.pass;
  let hash = crypto.createHash('md5').update(newpass).digest("hex");
  mongoose
    .connect(
      "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
      {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
      }
    )
    .then(async () => {
      if (await UserSchema.findOne({ username: session.mailconfirm, password: hash })) {
        res.send("error")
      }
      else {
        await UserSchema.findOneAndUpdate(
          { username: session.mailconfirm },
          { password: hash, change: "y" }
        );
        session.mailconfirm = null;
        session.code = null;
        res.send("Ok");
      }
    });
});

//Section user
//Forgot to say goodbye
routeExp.route("/forget").post(async function (req, res) {
  session = req.session;
  if (session.occupation_u == "User") {
    await update_last(req.body.timeforget, req.session, res);
  }
  else {
    res.send("retour");
  }
});
// function update last
async function update_last(time_given, session, res) {
  mongoose
    .connect(
      "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
      {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
      }
    )
    .then(async () => {
      var user = await UserSchema.findOne({ m_code: session.m_code });
      var last_time = await StatusSchema.findOne({ m_code: session.m_code, time_end: "", date: { $ne: moment().format("YYYY-MM-DD") } });
      if (parseInt(user.user_ht) != 0) {
        if (last_time) {
          if (hour_diff(last_time.time_start, time_given) <= (parseInt(user.user_ht) + 2)) {
            session.forget = "n";
            await StatusSchema.findOneAndUpdate({ m_code: session.m_code, time_end: "", date: { $ne: moment().format("YYYY-MM-DD") } }, { time_end: time_given }).sort({ _id: -1 });
            var left_data = await StatusSchema.find({ m_code: session.m_code, time_end: "" });
            for (i = 0; i < left_data.length; i++) {
              await StatusSchema.findOneAndUpdate({ _id: left_data[i]._id }, { time_end: calculate(left_data[i].time_start, parseInt(left_data[i].worktime)) });
            }
            res.send("Ok");
          }
          else {
            res.send("No");
          }
        }
        else {
          session.forget = "n";
          res.send("Ok");
        }
      }
      else {
        await StatusSchema.findOneAndUpdate({ m_code: session.m_code, time_end: "", date: { $ne: moment().format("YYYY-MM-DD") } }, { time_end: time_given }).sort({ _id: -1 });
        var left_data = await StatusSchema.find({ m_code: session.m_code, time_end: "" });
        for (i = 0; i < left_data.length; i++) {
          await StatusSchema.findOneAndUpdate({ _id: left_data[i]._id }, { time_end: calculate(left_data[i].time_start, 8) });
        }
        session.forget = "n";
        res.send("Ok");
      }
    });
  function calculate(begin, total) {
    return moment(begin, "HH:mm").add(total, "hours").format("HH:mm");
  }

}
routeExp.route("/reason").post(async function (req, res) {
  session = req.session;
  if (session.occupation_u == "User") {
    await reason_late(req.body.reason, session, res);
  }
  else {
    res.redirect("/")
  }
})
//reason late 
async function reason_late(reason, session, res) {
  mongoose
    .connect(
      "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
      {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
      }
    )
    .then(async () => {
      session.time = "y";
      await LateSchema.findOneAndUpdate({ m_code: session.m_code, reason: "", date: moment().format("YYYY-MM-DD") }, { reason: reason });
      res.send("Ok");
    })
}
// get hour
routeExp.route("/gethour").post(async function (req, res) {
  session = req.session;
  if (session.occupation_u == "User") {
    await get_hour(req.body.hour, req.session, res);
  }
  else {
    res.redirect("/")
  }
})
async function get_hour(h, session, res) {
  mongoose
    .connect(
      "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
      {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
      }
    )
    .then(async () => {
      await UserSchema.findOneAndUpdate({ m_code: session.m_code }, { user_ht: h });
      res.send("Ok");
    });

}
//Sending page for user
routeExp.route("/employee").get(async function (req, res) {
  session = req.session;
  if (session.occupation_u == "User") {
    mongoose
      .connect(
        "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
        {
          useUnifiedTopology: true,
          UseNewUrlParser: true,
        }
      )
      .then(async () => {
        var user = await UserSchema.findOne({ m_code: session.m_code });
        if (session.time != "y") {
          res.render("Working.html", { user: user, retard: session.time, forget: session.forget });
        }
        else {
          res.render("Working.html", { user: user, retard: "n", forget: session.forget });
        }

      })
  }
  else {
    res.redirect("/")
  }
})
//Getting user change
//Startwork
routeExp.route("/startwork").post(async function (req, res) {
  session = req.session;
  if (session.occupation_u == "User") {
    await startwork(req.body.timework, req.body.locaux, session, res);
  }
  else {
    res.send("error");
  }
});
//startwork
async function startwork(timework, locaux, session, res) {
  var date = moment().format("YYYY-MM-DD");
  var timestart = moment().add(3, 'hours').format("HH:mm");
  var new_time = {
    m_code: session.m_code,
    num_agent: session.num_agent,
    date: date,
    time_start: timestart,
    time_end: "",
    worktime: timework,
    nom: session.name,
    locaux: locaux
  }
  mongoose
    .connect(
      "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
      {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
      }
    )
    .then(async () => {

      if (await StatusSchema.findOne({ m_code: session.m_code, date: moment().format("YYYY-MM-DD") })) {
        var send_time = await StatusSchema.findOne({ m_code: session.m_code, date: moment().format("YYYY-MM-DD") });
        await AbsentSchema.findOneAndUpdate({ m_code: session.m_code, return: "Not come back", date: moment().format("YYYY-MM-DD") }, { return: timestart });
        await StatusSchema.findOneAndUpdate({ m_code: session.m_code, date: moment().format("YYYY-MM-DD") }, { time_end: "", locaux: loc_check(send_time.locaux, locaux), worktime: timework });
        res.send(send_time.time_start);
      }
      else {
        await StatusSchema(new_time).save();
        res.send(timestart);
      }


    });

}
//locaux check
function loc_check(l, n) {
  if (l == n || l.includes(n)) {
    return l;
  }
  else {
    return l + "/" + n;
  }
}
//Change hour
routeExp.route("/changing").post(async function (req, res) {
  session = req.session;
  if (session.occupation_u == "User") {
    await changing(req.body.ch_hour, req.session, res);
  }
  else {
    res.redirect("/")
  }
})
async function changing(ch, session, res) {
  mongoose
    .connect(
      "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
      {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
      }
    )
    .then(async () => {
      await UserSchema.findOneAndUpdate({ m_code: session.m_code }, { user_ht: ch });
      await StatusSchema.findOneAndUpdate({ m_code: session.m_code, time_end: "" }, { worktime: ch });
      res.send("Ok");
    })
}

//Change status
routeExp.route("/statuschange").post(async function (req, res) {
  session = req.session;
  if (session.occupation_u == "User") {
    await status_change(req.body.act_loc, req.body.act_stat, res);
  }
  else {
    res.send("error");
  }
});
async function status_change(lc, st, res) {
  mongoose
    .connect(
      "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
      {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
      }
    )
    .then(async () => {
      await UserSchema.findOneAndUpdate({ m_code: session.m_code }, { act_stat: st, act_loc: lc });
      res.send(st + "," + moment().add(3, 'hours').format("HH:mm"));
    });
}
//handleWork
routeExp.route("/handlework").post(async function (req, res) {
  session = req.session;
  if (session.occupation_u == "User") {
    await handlework(req.body.locaux, req.body.choice, session, res);
  }
  else {
    res.send("error");
  }
});
async function handlework(locaux, choice, session, res) {
  mongoose
    .connect(
      "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
      {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
      }
    )
    .then(async () => {
      var date = moment().format("YYYY-MM-DD");
      if (choice != "POSTE") {
        await StatusSchema.findOneAndUpdate({ m_code: session.m_code, date: date, time_end: "" }, { locaux: locaux + "/" + choice });
        await UserSchema.findOneAndUpdate({ m_code: session.m_code }, { act_loc: choice, late: "y" });
      }
      res.redirect("/exit_u");
    });
}
//Left work
routeExp.route("/leftwork").post(async function (req, res) {
  session = req.session;
  if (session.occupation_u == "User") {
    await leftwork(req.body.locaux, session, res);
  }
  else {
    res.send("error");
  }
});
async function leftwork(locaux, session, res) {
  var date = moment().format("YYYY-MM-DD");
  var timeend = moment().add(3, 'hours').format("HH:mm");
  mongoose
    .connect(
      "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
      {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
      }
    )
    .then(async () => {
      await StatusSchema.findOneAndUpdate({ m_code: session.m_code, date: date, time_end: "" }, { time_end: timeend });
      await UserSchema.findOneAndUpdate({ m_code: session.m_code }, { late: "n", count: 0 });
      res.redirect("/exit_u");
    });
}
//activity
routeExp.route("/activity").post(async function (req, res) {
  session = req.session;
  if (session.occupation_u == "User") {
    await activity(req.body.activity, session, req, res);
  }
  else {
    res.send("error");
  }

});
async function activity(activity, session, req, res) {
  if (activity != "ABSENT") {
    var counter = 0;
    switch (activity) {
      case "BREAK": counter = 1200000; session.place = "petit break"; break;
      case "DEJEUNER": counter = 2400000; session.place = "Déjeuner"; break;
      case "PAUSE": counter = 1800000; session.place = "Pause"; break;
    }
    mongoose
      .connect(
        "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
        {
          useUnifiedTopology: true,
          UseNewUrlParser: true,
        }
      )
      .then(async () => {
        await UserSchema.findOneAndUpdate({ m_code: session.m_code }, { $inc: { count: 1 }, take_break: "n" });
        var counts = await UserSchema.findOne({ m_code: session.m_code });
        if (counts.count > 6) {
          notification.push(session.name + " quitte son poste trop frequement ");
          const io = req.app.get('io');
          io.sockets.emit('notif', notification);
        }
        setTimeout(async function () {
          counts = await UserSchema.findOneAndUpdate({ m_code: session.m_code }, { $inc: { count: 1 }, take_break: "n" });
          if (counts.take_break == "n") {
            notification.push(session.name + " prend trop de temp au " + session.place);
            const io = req.app.get('io');
            io.sockets.emit('notif', notification);
          }
        }, counter);
      })
    res.send("Ok");
  }
  else {
    res.send("Ok");
  }

}
//Take break
routeExp.route("/takebreak").post(async function (req, res) {
  session = req.session;
  await take_break(session);
  res.send("Ok");
})
async function take_break(session) {
  mongoose
    .connect(
      "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
      {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
      }
    )
    .then(async () => {
      await UserSchema.findOneAndUpdate({ m_code: session.m_code }, { take_break: "y" });
    })
}
// Remarques 
routeExp.route("/rem").post(async function (req, res) {
  mongoose
    .connect(
      "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
      {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
      }
    )
    .then(async () => {
      var new_rem = {
        m_code: req.body.m_code,
        date: moment().format("YYYY-MM-DD"),
        remarques: req.body.rem
      }
      await RemSchema(new_rem).save();
      res.send("Ok");
    })
})
//Disconnect_user
routeExp.route("/exit_u").get(function (req, res) {
  session = req.session;
  session.occupation_u = null;
  session.mcode = null;
  session.num_agent = null;
  session.forget = null;
  session.time = null;
  res.redirect("/");
});
// Administrator parts
//Page home
routeExp.route("/home").get(async function (req, res) {
  session = req.session;
  if (session.occupation_a == "Admin") {

    mongoose
      .connect(
        "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
        {
          useUnifiedTopology: true,
          UseNewUrlParser: true,
        }
      )
      .then(async () => {
        var nbr_employe = await UserSchema.find({occupation:{$ne:"Admin"}});
        var nbr_actif = await UserSchema.find({ act_stat: { $ne: "VACATION" }, occupation:{$ne:"Admin"} });
        var nbr_leave = await UserSchema.find({ act_stat: "VACATION" });
        var nbr_retard = await LateSchema.find({ date: moment().format("YYYY-MM-DD"), validation: false });
        await monthly_restart();
        res.render("dashboard.html", { notif: notification, nbr_emp: nbr_employe.length, nbr_act: nbr_actif.length, nbr_leave: nbr_leave.length, nbr_retard: nbr_retard.length });
      });
  }
  else {
    res.redirect("/");
  }
})
//Page status
routeExp.route("/management").get(async function (req, res) {
  session = req.session;
  if (session.occupation_a == "Admin") {
    session.filtrage = null;
    mongoose
      .connect(
        "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
        {
          useUnifiedTopology: true,
          UseNewUrlParser: true,
        }
      )
      .then(async () => {
        var alluser = await UserSchema.find({});
        res.render("status.html", { users: alluser, notif: notification });
      })
  }
  else {
    res.redirect("/");
  }
})
//Page statustl
routeExp.route("/managementtl").get(async function (req, res) {
  session = req.session;
  if (session.occupation_tl == "checker") {
    mongoose
      .connect(
        "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
        {
          useUnifiedTopology: true,
          UseNewUrlParser: true,
        }
      )
      .then(async () => {
        var alluser = await UserSchema.find({});
        res.render("statustl.html", { users: alluser, notif: notification });
      })
  }
  else {
    res.redirect("/");
  }
})
//Listuser
routeExp.route("/userlist").get(async function (req, res) {
  session = req.session;
  if (session.occupation_a == "Admin") {
    session.filtrage = null;
    mongoose
      .connect(
        "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
        {
          useUnifiedTopology: true,
          UseNewUrlParser: true,
        }
      )
      .then(async () => {
        var alluser = await UserSchema.find({});
        res.render("userlist.html", { users: alluser, notif: notification });
      })
  }
  else {
    res.redirect("/");
  }
})
//Add employee
routeExp.route("/addemp").post(async function (req, res) {
  session = req.session;
  var email = req.body.email;
  var name = req.body.name;
  var last_name = req.body.last_name;
  var usuel = req.body.usuel;
  var mcode = req.body.mcode;
  var num_agent = req.body.num_agent;
  var matricule = req.body.matricule;
  var fonction = req.body.function_choosed;
  var occupation = req.body.occupation;
  var save_at = req.body.enter_date;
  var cin = req.body.cin;
  var sexe = req.body.gender;
  var situation = req.body.situation;
  var location = req.body.location;
  var num_cnaps = req.body.num_cnaps;
  var classification = req.body.classification;
  var contrat = req.body.contrat;
  var change = "n";
  var late = "n";
  mongoose
    .connect(
      "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
      {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
      }
    )
    .then(async () => {
      if (
        await UserSchema.findOne({
          $or:
            [
              { m_code: mcode },
              { username: email },
              { num_agent: num_agent },
              { matr: matricule }
            ]
        })
      ) {
        res.send("already");
      } else {
        var passdefault = randomPassword();
        let hash = crypto.createHash('md5').update(passdefault).digest("hex");
        var new_emp = {
          username: email,
          first_name: name,
          last_name: last_name,
          usuel: usuel,
          m_code: mcode,
          matr: matricule,
          occupation: occupation,
          shift: fonction,
          save_at: moment(save_at).format("YYYY-MM-DD"),
          cin: cin,
          sexe: sexe,
          situation: situation,
          adresse: location,
          cnaps_num: num_cnaps,
          classification: classification,
          contrat: contrat,
          password: hash,
          num_agent: num_agent,
          change: change,
          act_stat: "LEFTING",
          act_loc: "Not defined",
          late: late,
          count: 0,
          take_break: "n",
          remaining_leave: 0,
          leave_taked: 0,
          project: "",
          leave_stat: "n",
          user_ht: 0
        };
        await UserSchema(new_emp).save();
        sendEmail(
          email,
          "Authentification Solumada",
          htmlRender(email, passdefault)
        );
        res.send(email);
      }
    });
});
//getuser
routeExp.route("/getuser").post(async function (req, res) {
  var id = req.body.id;
  mongoose
    .connect(
      "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
      {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
      }
    )
    .then(async () => {
      var user = await UserSchema.findOne({ _id: id });
      res.send(user.username + "," + user.first_name + "," + user.last_name + "," + user.usuel
        + "," + user.m_code + "," + user.num_agent + "," + user.matr + "," + user.shift + "," + user.occupation
        + "," + user.save_at + "," + user.cin + "," + user.sexe + "," + user.situation
        + "," + user.adresse + "," + user.cnaps_num + "," + user.classification + "," + user.contrat);
    });
});
//Update User
routeExp.route("/updateuser").post(async function (req, res) {
  var id = req.body.id;
  var email = req.body.email;
  var name = req.body.name;
  var last_name = req.body.last_name;
  var usuel = req.body.usuel;
  var fonction = req.body.function_choosed;
  var occupation = req.body.occupation;
  var save_at = req.body.enter_date;
  var cin = req.body.cin;
  var sexe = req.body.gender;
  var situation = req.body.situation;
  var location = req.body.location;
  var num_cnaps = req.body.num_cnaps;
  var classification = req.body.classification;
  var contrat = req.body.contrat;
  mongoose
    .connect(
      "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
      {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
      }
    )
    .then(async () => {
      var user = await UserSchema.findOne({ _id: id });
      await StatusSchema.updateMany({ m_code: user.m_code }, { nom: name + " " + last_name });
      await UserSchema.findOneAndUpdate({ _id: id }, {
        username: email,
        first_name: name,
        last_name: last_name,
        usuel: usuel,
        occupation: occupation,
        shift: fonction,
        save_at: moment(save_at).format("YYYY-MM-DD"),
        cin: cin,
        sexe: sexe,
        situation: situation,
        adresse: location,
        cnaps_num: num_cnaps,
        classification: classification,
        contrat: contrat
      });
      res.send("Ok");
    })
})
//update project 
routeExp.route("/update_project").post(async function (req, res) {
  var choice = req.body.choice.split(",");
  var owner = req.body.owner;
  var combined = "";
  for (i = 0; i < choice.length; i++) {
    if (i == 0) {
      combined += choice[i];
    }
    else {
      combined += "/" + choice[i];
    }
  }
  mongoose
    .connect(
      "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
      {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
      }
    )
    .then(async () => {
      await UserSchema.findOneAndUpdate({ m_code: owner }, { project: combined });
      res.send("ok")
    })
})
//Drop user 
routeExp.route("/dropuser").post(async function (req, res) {
  var names = req.body.fname;
  names = names.split(" ");
  mongoose
    .connect(
      "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
      {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
      }
    )
    .then(async () => {
      await UserSchema.findOneAndDelete({ m_code: names });
      res.send("User deleted successfully");
    });
});
//Sheets
routeExp.route("/details").get(async function (req, res) {
  session = req.session;
  if (session.occupation_a == "Admin") {
    mongoose
      .connect(
        "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
        {
          useUnifiedTopology: true,
          UseNewUrlParser: true,
        }
      )
      .then(async () => {
        if (session.filtrage == "" && data_desired.datatowrite) {
          res.render("details.html", { timesheets: data_desired.datatowrite, notif: notification });
          session.filtrage = null;
        }
        else {
          var timesheets = await StatusSchema.find({ time_end: { $ne: "" } }).sort({
            "_id": -1,
          }).limit(200);
          data_desired.datatowrite = timesheets;
          data_desired.datalate = await LateSchema.find({ validation: true });
          data_desired.dataabsence = await AbsentSchema.find({});
          data_desired.dataleave = await LeaveSchema.find({});
          res.render("details.html", { timesheets: timesheets, notif: notification });
        }
      })
  }
  else {
    res.redirect("/");
  }
})
//Filter sheets
routeExp.route("/filter").post(async function (req, res) {
  session = req.session;
  if (session.occupation_a == "Admin") {
    session.filtrage = "";
    var searchit = req.body.searchit;
    var period = req.body.period;
    var datestart = "";
    var dateend = "";
    if (period == "t") {
      datestart = moment().format("YYYY-MM-DD");
    }
    else if (period == "tw") {
      datestart = moment().startOf("week").format("YYYY-MM-DD");
      dateend = moment().endOf("week").format("YYYY-MM-DD");
    }
    else if (period == "tm") {
      datestart = moment().startOf("month").format("YYYY-MM-DD");
      dateend = moment().endOf("month").format("YYYY-MM-DD");
    }
    else if (period == "spec") {
      datestart = req.body.datestart;
      dateend = req.body.dateend
    }
    else {
      datestart = "";
      dateend = "";
    }
    var datecount = [];
    var datatosend = [];
    mongoose
      .connect(
        "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
        {
          useUnifiedTopology: true,
          UseNewUrlParser: true,
        }
      )
      .then(async () => {

        var late_temp = [];
        var absent_temp = [];
        var temp_leave = [];
        late_temp.push([]);
        absent_temp.push([]);
        temp_leave.push([]);
        datestart == "" ? "" : datecount.push(1);
        dateend == "" ? "" : datecount.push(2);
        searchit == "" ? delete filtrage.search : filtrage.search = searchit;
        if (datecount.length == 2) {
          var day = moment
            .duration(
              moment(dateend, "YYYY-MM-DD").diff(moment(datestart, "YYYY-MM-DD"))
            )
            .asDays();
          var getdata;
          var getdatalate;
          var getdataabsence;
          var getdataleave;
          for (i = 0; i <= day; i++) {
            filtrage.date = datestart;
            date_data.push(filtrage.date);
            if (filtrage.search) {
              getdata = await StatusSchema.find({
                $or:
                  [{ m_code: { '$regex': searchit, '$options': 'i' } },
                  { nom: { '$regex': searchit, '$options': 'i' } },
                  { locaux: { '$regex': searchit, '$options': 'i' } },], date: filtrage.date, time_end: { $ne: "" }
              }).sort({
                "_id": -1,
              });
              getdatalate = await LateSchema.find({ date: filtrage.date}).sort({
                "_id": -1,
              });
              getdataabsence = await AbsentSchema.find({ date: filtrage.date }).sort({
                "_id": -1,
              });
              getdataleave = await LeaveSchema.find({ date_start: filtrage.date }).sort({
                "_id": -1,
              });
            }
            else {
              getdata = await StatusSchema.find({ date: filtrage.date, time_end: { $ne: "" } });
              getdatalate = await LateSchema.find({ date: filtrage.date});
              getdataabsence = await AbsentSchema.find({ date: filtrage.date });
              getdataleave = await LeaveSchema.find({ date_start: filtrage.date });
            }

            if (getdata.length != 0) {
              datatosend.push(getdata);
            }
            var addday = moment(datestart, "YYYY-MM-DD")
              .add(1, "days")
              .format("YYYY-MM-DD");
            datestart = addday;
            if (getdatalate.length != 0) {
              for (l = 0; l < getdatalate.length; l++) {
                late_temp[0].push(getdatalate[l]);
              }
            }
            if (getdataabsence.length != 0) {
              for (ab = 0; ab < getdataabsence.length; ab++) {
                absent_temp[0].push(getdataabsence[ab]);
              }
            }
            if (getdataleave.length != 0) {
              for (lv = 0; lv < getdataleave.length; lv++) {
                temp_leave[0].push(getdataleave[lv]);
              }
            }
          }

          for (i = 1; i < datatosend.length; i++) {
            for (d = 0; d < datatosend[i].length; d++) {
              datatosend[0].push(datatosend[i][d]);
            }
          }
          if (datatosend.length != 0) {
            data_desired.datatowrite = datatosend[0];
            data_desired.datalate = late_temp[0];
            data_desired.dataabsence = absent_temp[0];
            data_desired.dataleave = temp_leave[0];
            res.send(datatosend[0]);
          }
          else {
            data_desired.datatowrite = datatosend;
            data_desired.datalate = [];
            data_desired.dataabsence = [];
            data_desired.dataleave = [];
            res.send(datatosend);
          }

        } else if (datecount.length == 1) {
          if (datecount[0] == 1) {
            filtrage.date = datestart;
            if (filtrage.search) {
              datatosend = await StatusSchema.find({
                $or:
                  [{ m_code: { '$regex': searchit, '$options': 'i' } },
                  { nom: { '$regex': searchit, '$options': 'i' } },
                  { locaux: { '$regex': searchit, '$options': 'i' } },], date: filtrage.date, time_end: { $ne: "" }
              }).sort({
                "_id": -1,
              });
              data_desired.datalate = await LateSchema.find({ date: filtrage.date}).sort({
                "_id": -1,
              });
              data_desired.dataabsence = await AbsentSchema.find({ date: filtrage.date }).sort({
                "_id": -1,
              });
              data_desired.dataleave = await LeaveSchema.find({ date_start: filtrage.date }).sort({
                "_id": -1,
              });
            }
            else {
              datatosend = await StatusSchema.find({ date: filtrage.date, time_end: { $ne: "" } });
              data_desired.datalate = await LateSchema.find({ date: filtrage.date});
              data_desired.dataabsence = await AbsentSchema.find({ date: filtrage.date });
              data_desired.dataleave = await LeaveSchema.find({ date_start: filtrage.date });
            }
            data_desired.datatowrite = datatosend;
            session.searchit = searchit;
            res.send(datatosend);
          } else {
            filtrage.date = dateend;
            if (filtrage.search) {
              datatosend = await StatusSchema.find({
                $or:
                  [{ m_code: { '$regex': searchit, '$options': 'i' } },
                  { nom: { '$regex': searchit, '$options': 'i' } },
                  { locaux: { '$regex': searchit, '$options': 'i' } },], date: filtrage.date, time_end: { $ne: "" }
              }).sort({
                "_id": -1,
              });
              data_desired.datalate = await LateSchema.find({ date: filtrage.date}).sort({
                "_id": -1,
              });
              data_desired.dataabsence = await AbsentSchema.find({ date: filtrage.date }).sort({
                "_id": -1,
              });
              data_desired.dataleave = await LeaveSchema.find({ date_start: filtrage.date }).sort({
                "_id": -1,
              });
            }
            else {
              datatosend = await StatusSchema.find({ date: filtrage.date });
            }
            data_desired.datatowrite = datatosend;
            session.searchit = searchit;
            res.send(datatosend);
          }
        } else {
          delete filtrage.date;
          datatosend = await StatusSchema.find({
            $or:
              [{ m_code: { '$regex': searchit, '$options': 'i' } },
              { nom: { '$regex': searchit, '$options': 'i' } },
              { locaux: { '$regex': searchit, '$options': 'i' } },], time_end: { $ne: "" }
          }).sort({ "_id": -1, });
          data_desired.datatowrite = datatosend;
          data_desired.datalate = await LateSchema.find({ validation: true }).sort({
            "_id": -1,
          });
          data_desired.dataabsence = await AbsentSchema.find({}).sort({
            "_id": -1,
          });;
          data_desired.dataleave = await LeaveSchema.find({}).sort({
            "_id": -1,
          });;
          session.searchit = searchit;
          res.send(datatosend);
        }
      });
  }
  else {
    res.send("error");
  }
});

routeExp.route("/latelist").get(async function (req, res) {
  session = req.session;
  if (session.occupation_a == "Admin") {
    mongoose
      .connect(
        "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
        {
          useUnifiedTopology: true,
          UseNewUrlParser: true,
        }
      )
      .then(async () => {
        var latelist = await LateSchema.find({ validation: true, reason: { $ne: "" } });
        res.render("latelist.html", { latelist: latelist, notif: notification });
      });
  }
  else {
    res.redirect("/");
  }
});
//Validation page
routeExp.route("/validelate").get(async function (req, res) {
  session = req.session;
  if (session.occupation_a == "Admin") {
    session.filtrage = null;
    mongoose
      .connect(
        "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
        {
          useUnifiedTopology: true,
          UseNewUrlParser: true,
        }
      )
      .then(async () => {
        var latelist = await LateSchema.find({ validation: false });
        res.render("latevalidation.html", { latelist: latelist, notif: notification });
      });
  } else {
    res.redirect("/");
  }
});
//Validation
routeExp.route("/validate").post(async function (req, res) {
  session = req.session;
  if (session.occupation_a == "Admin") {
    var id = req.body.id;
    mongoose
      .connect(
        "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
        {
          useUnifiedTopology: true,
          UseNewUrlParser: true,
        }
      )
      .then(async () => {
        await LateSchema.findOneAndUpdate(
          { _id: id },
          { validation: true }
        );
        res.send("Ok");
      });
  }
  else {
    res.redirect("/");
  }
});
//Denied
routeExp.route("/denied").post(async function (req, res) {
  session = req.session;
  if (session.occupation_a == "Admin") {
    var id = req.body.id;
    mongoose
      .connect(
        "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
        {
          useUnifiedTopology: true,
          UseNewUrlParser: true,
        }
      )
      .then(async () => {
        await LateSchema.findOneAndDelete({ _id: id });
        res.send("Ok");
      });
  }
  else {
    res.send("retour");
  }
});
// routeExp.route("/newuser").get(async function (req, res) {
//   session = req.session;
//   if (session.occupation_a == "Admin"){
//   res.render("newuser.html");
//   }
//   else{
//     res.redirect("/")
//   }
// })
//Absent
routeExp.route("/absent").post(async function (req, res) {
  session = req.session;
  if (session.occupation_u == "User") {
    await absent(req.body.reason, req.body.stat, session, res);
  }
  else {
    res.send("error");
  }
});
//Absent
async function absent(reason, stat, session, res) {
  var timestart = moment().add(3, 'hours').format("HH:mm");
  var new_abs = {
    m_code: session.m_code,
    num_agent: session.num_agent,
    date: moment().format("YYYY-MM-DD"),
    nom: session.name,
    time_start: timestart,
    return: "Not come back",
    reason: reason,
    validation: false,
    status: "none"
  }
  mongoose
    .connect(
      "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
      {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
      }
    )
    .then(async () => {
      await AbsentSchema(new_abs).save();
      await StatusSchema.findOneAndUpdate({ m_code: session.m_code, date: moment().format("YYYY-MM-DD"), time_end: "" }, { time_end: timestart });
      await UserSchema.findOneAndUpdate({ m_code: session.m_code }, { act_loc: "Not defined", act_stat: "LEFTING" });
      res.send(stat);
    });
}
routeExp.route("/absencelist").get(async function (req, res) {
  session = req.session;
  if (session.occupation_a == "Admin") {
    mongoose
      .connect(
        "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
        {
          useUnifiedTopology: true,
          UseNewUrlParser: true,
        }
      )
      .then(async () => {
        var absent = await AbsentSchema.find({ validation: true });
        res.render("abscencelist.html", { absent: absent, notif: notification });
      })
  }
  else {
    res.redirect("/");
  }
})
//Validation page
routeExp.route("/valideabsence").get(async function (req, res) {
  session = req.session;
  if (session.occupation_a == "Admin") {
    session.filtrage = null;
    mongoose
      .connect(
        "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
        {
          useUnifiedTopology: true,
          UseNewUrlParser: true,
        }
      )
      .then(async () => {
        var absent = await AbsentSchema.find({ validation: false });
        res.render("absencevalidation.html", { absent: absent, notif: notification });
      });
  } else {
    res.redirect("/");
  }
});
//Validation
routeExp.route("/validate_a").post(async function (req, res) {
  session = req.session;
  if (session.occupation_a == "Admin") {
    var id = req.body.id;
    mongoose
      .connect(
        "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
        {
          useUnifiedTopology: true,
          UseNewUrlParser: true,
        }
      )
      .then(async () => {
        await AbsentSchema.findOneAndUpdate(
          { _id: id },
          { validation: true, status: "Accepter" }
        );
        res.send("Ok");
      });
  }
  else {
    res.redirect("/");
  }
});
//Denied
routeExp.route("/denied_a").post(async function (req, res) {
  session = req.session;
  if (session.occupation_a == "Admin") {
    var id = req.body.id;
    mongoose
      .connect(
        "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
        {
          useUnifiedTopology: true,
          UseNewUrlParser: true,
        }
      )
      .then(async () => {
        await AbsentSchema.findOneAndUpdate({ _id: id }, { status: "Non communiqué" });
        res.send("Ok");
      });
  }
  else {
    res.send("retour");
  }
});
//Page Leave
routeExp.route("/leave").get(async function (req, res) {
  session = req.session;
  if (session.occupation_a == "Admin") {
    mongoose
      .connect(
        "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
        {
          useUnifiedTopology: true,
          UseNewUrlParser: true,
        }
      )
      .then(async () => {
        var alluser = await UserSchema.find({ occupation: "User" });
        var vacation = await UserSchema.find({ occupation: "User", act_stat: "VACATION" });
        res.render("conge.html", { users: alluser, notif: notification, vacation: vacation });
      })
  }
  else {
    res.redirect("/");
  }
});
//List leave
routeExp.route("/leavelist").get(async function (req, res) {
  session = req.session;
  if (session.occupation_a == "Admin") {
    mongoose
      .connect(
        "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
        {
          useUnifiedTopology: true,
          UseNewUrlParser: true,
        }
      )
      .then(async () => {
        if (monthly_leave.length == 0) {
          var leavelist = await LeaveSchema.find({}).sort({
            "_id": -1,
          }).limit(100);
          res.render("congelist.html", { leavelist: leavelist, notif: notification });
        }
        else {
          res.render("congelist.html", { leavelist: monthly_leave, notif: notification });
        }

      });
  }
  else {
    res.redirect("/");
  }
});
//take leave
routeExp.route("/takeleave").post(async function (req, res) {
  var userid = req.body.id;
  var type = req.body.type;
  var leavestart = req.body.leavestart;
  var leaveend = req.body.leaveend;
  var val = req.body.court;
  var edit = req.body.edit;
  var deduction = " ( rien à deduire )";
  mongoose
    .connect(
      "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
      {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
      }
    )
    .then(async () => {
      var user = await UserSchema.findOne({ _id: userid });
      var taked;
      if (edit != "n") {
        if (deduire.includes(type)) {
          type = type + " ( a déduire sur salaire )"
        }
        else {
          type = type + deduction;
        }
        if (val == "n") {
          taked = date_diff(leavestart, leaveend) + 1;
        }
        else {
          taked = val;
          leaveend = leavestart;
        }
        var last_leave = await LeaveSchema.findOne({ _id: edit })
        if (last_leave.type == "Congé Payé ( rien à deduire )") {
          await UserSchema.findOneAndUpdate({ m_code: user.m_code }, { $inc: { remaining_leave: (last_leave.duration - taked), leave_taked: (last_leave.duration - taked) } });
          await LeaveSchema.findOneAndUpdate({ _id: edit }, { type: type, date_start: leavestart, date_end: leaveend, duration: taked, rest: last_leave.rest + (last_leave.duration - taked) });
        }
        else {
          await LeaveSchema.findOneAndUpdate({ _id: edit }, { type: type, date_start: leavestart, date_end: leaveend, duration: taked });
        }
        res.send("Ok");
      }
      else {
        if (await LeaveSchema.findOne({ m_code: user.m_code, status: "en attente" })) {
          res.send("already");
        }
        else if (await LeaveSchema.findOne({ m_code: user.m_code, date_start: leavestart }) || await LeaveSchema.findOne({ m_code: user.m_code, date_end: leaveend }) ||
          await LeaveSchema.findOne({ m_code: user.m_code, date_start: leaveend }) || await LeaveSchema.findOne({ m_code: user.m_code, date_end: leavestart })) {
          res.send("duplicata");
        }
        else {

          if (val == "n") {
            taked = date_diff(leavestart, leaveend) + 1;
          }
          else {
            if (val == 0.5) {
              leaveend = leavestart;
              taked = val;
            }
            else {
              leaveend = leavestart;
              taked = val;
            }

          }
          if (user.leave_stat == "y" && (type == "Congé Payé")) {
            if (deduire.includes(type)) {
              deduction = " ( a déduire sur salaire )";
            }
            var day_control = "Terminée";
            if (taked >= 1) {
              day_control = "en attente";
            }
            var new_leave = {
              m_code: user.m_code,
              num_agent: user.num_agent,
              nom: user.first_name + " " + user.last_name,
              date_start: leavestart,
              date_end: leaveend,
              duration: taked,
              type: type + deduction,
              status: day_control,
              rest: user.remaining_leave - taked,
              validation: false
            }
            await LeaveSchema(new_leave).save();
            await UserSchema.findOneAndUpdate({ m_code: user.m_code }, { $inc: { remaining_leave: -taked, leave_taked: -taked } });
            await conge_define(req);
            await checkleave();
            res.send("Ok");
          }
          else if (type == "Mise a Pied" || type == "Permission exceptionelle" || type == "Repos Maladie" || type == "Congé de maternité" || type == "Absent" || type == "Congé sans solde") {
            if (deduire.includes(type)) {
              deduction = " ( a déduire sur salaire )";
            }
            var day_control = "Terminée";
            if (taked >= 1) {
              day_control = "en attente";
            }
            var new_leave = {
              m_code: user.m_code,
              num_agent: user.num_agent,
              nom: user.first_name + " " + user.last_name,
              date_start: leavestart,
              date_end: leaveend,
              duration: taked,
              type: type + deduction,
              status: day_control,
              rest: user.remaining_leave,
              validation: false
            }
            await LeaveSchema(new_leave).save();
            await conge_define(req);
            await checkleave();
            res.send("Ok");
          }
          else {
            res.send("not authorized");
          }
        }
      }

    })
})
async function leave_permission() {
  mongoose
    .connect(
      "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
      {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
      }
    )
    .then(async () => {
      var user_allowed = await UserSchema.find({})
      for (a = 0; a < user_allowed.length; a++) {
        if (difference_year(user_allowed[a].save_at) && user_allowed[a].leave_stat == "n") {
          await UserSchema.findOneAndUpdate({ m_code: user_allowed[a].m_code }, { leave_stat: "y" });
        }
      }


    })
}
async function conge_define(req) {
  mongoose
    .connect(
      "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
      {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
      }
    )
    .then(async () => {
      console.log("début de congé verifier")
      try {
        var all_leave1 = await LeaveSchema.find({ status: "en attente" });
        for (i = 0; i < all_leave1.length; i++) {
          if (moment().format("YYYY-MM-DD") == all_leave1[i].date_start) {
            if (all_leave1[i].duration >= 1) {
              await UserSchema.findOneAndUpdate({ m_code: all_leave1[i].m_code }, { act_stat: "VACATION", act_loc: "Not defined" });
              await LeaveSchema.findOneAndUpdate({ _id: all_leave1[i]._id }, { status: "en cours" });
              const io = req.app.get('io');
              io.sockets.emit('status', "VACATION" + "," + all_leave1[i].m_code);
            }
            else {
              await LeaveSchema.findOneAndUpdate({ _id: all_leave1[i]._id }, { status: "Terminée" });
            }

          }
          else if (date_diff(moment().format("YYYY-MM-DD"), all_leave1[i].date_start) < 0) {
            if (date_diff(moment().format("YYYY-MM-DD"), all_leave1[i].date_start) * -1 < all_leave1[i].duration && all_leave1[i].duration > 1) {
              await UserSchema.findOneAndUpdate({ m_code: all_leave1[i].m_code }, { act_stat: "VACATION", act_loc: "Not defined" });
              await LeaveSchema.findOneAndUpdate({ _id: all_leave1[i]._id }, { status: "en cours" });
              const io = req.app.get('io');
              io.sockets.emit('status', "VACATION" + "," + all_leave1[i].m_code);
            }
            else {
              await LeaveSchema.findOneAndUpdate({ _id: all_leave1[i]._id }, { status: "Terminée" });
            }

          }
        }
      } catch (error) {
        await conge_define(req);
      }
    })
}
//checkleave
async function checkleave() {
  console.log("verification de congé verifier");
  mongoose
    .connect(
      "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
      {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
      }
    )
    .then(async () => {
      try {
        var all_leave2 = await LeaveSchema.find({ status: "en cours" });
        for (j = 0; j < all_leave2.length; j++) {
          if (date_diff(moment().format("YYYY-MM-DD"), all_leave2[j].date_end) < 0) {
            await UserSchema.findOneAndUpdate({ m_code: all_leave2[j].m_code }, { act_stat: "LEFTING" });
            await LeaveSchema.findOneAndUpdate({ _id: all_leave2[j]._id }, { status: "Terminée" });
            notification.push(all_leave2[j].nom + " devrait revenir du congé");
          }
        }
      } catch (error) {
        await checkleave();
      }

    })
}
routeExp.route("/fiche").get(function (req, res) {
  session = req.session;
  if (session.occupation_a == "Admin") {
    mongoose
      .connect(
        "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
        {
          useUnifiedTopology: true,
          UseNewUrlParser: true,
        }
      )
      .then(async () => {
        var opt = await OptSchema.findOne({ _id: "636247a2c1f6301f15470344" });
        res.render("paie.html", { opt: opt.paie_generated, list: opt.list_paie, notif: notification });
      })

  }
  else {
    res.redirect("/");
  }
})
// async function conge_change(){
//   mongoose
//   .connect(
//     "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
//     {
//       useUnifiedTopology: true,
//       UseNewUrlParser: true,
//     }
//   )
//   .then(async () => {
//           const file = ExcelFile.readFile("cr.xlsx")
//           let data = []
//           const temp = ExcelFile.utils.sheet_to_json(
//                 file.Sheets[file.SheetNames[0]])
//           temp.forEach((res) => {
//               data.push(res)
//           })
//           for (i=0;i<data.length;i++){
//             await UserSchema.findOneAndUpdate({m_code:data[i].MCODE},{remaining_leave:data[i].ANNEE,leave_taked:data[i].CONGES})
//             console.log(data[i].MCODE + " Done");
//           }
//         })
// }
//conge_change();
//conge_change();
//Paie code
routeExp.route("/paie").post(function (req, res) {
  // When a file has been uploaded
  var start = req.body.start;
  var end = req.body.end;
  var ouvrable = req.body.ouvrable;
  var ouvre = req.body.ouvre;
  var calendaire = req.body.calendaire;
  req.files["fileup"].mv(req.files["fileup"].name);
  setTimeout(() => {
    readfile(req.files["fileup"].name, res, start, end, ouvrable, ouvre, calendaire);
  }, 3000)
})
async function readfile(name_file, res, start, end, ouvrable, ouvres, calendaire) {
  mongoose
    .connect(
      "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
      {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
      }
    )
    .then(async () => {
      const file = ExcelFile.readFile(name_file)
      let data = []
      const temp = ExcelFile.utils.sheet_to_json(
        file.Sheets[file.SheetNames[0]])
      temp.forEach((res) => {
        data.push(res)
      })
      list_paie = [];
      fs.unlink(name_file, function (err) {
        if (err) {
          console.error(err);
        }
        console.log('Excel file deleted');
      });
      test_paie(data, 0, res, start, end, ouvrable, ouvres, calendaire);

    })
}
var list_paie = [];
var excel_paie = [];
function test_paie(data, number, res, start, end, ouvrable, ouvres, calendaire) {
  try {
    var m_code = data[number]["M-CODE"];
    var nom = data[number]["Nom & Prénoms"];
    var num_agent = data[number]["NUMBERING AGENT"];
    var salaire = parseFloat(data[number]["SALAIRE"]);
    var hs30 = parseFloat(data[number]["HEURE SUPPLE 130"]);
    var hs50 = parseFloat(data[number]["HEURE SUPPLE 150"]);
    var hs100 = parseFloat(data[number]["HEURE SUPPLE 200"]);
    var nuit_occ = parseFloat(data[number]["NUIT OCCASIONNELLE"]);
    var rendement = parseFloat(data[number]["RENDEMENT"]);
    var soir = parseFloat(data[number]["TOTAL MAJ NUIT"]);
    var ferie = parseFloat(data[number]["TOTAL MAJ FERIE"]);
    var weekend = parseFloat(data[number]["TOTAL MAJ WEEK END"]);
    var transport = parseFloat(data[number]["DEPLACEMENT"]);
    var repas = parseFloat(data[number]["NBRE REPAS"]);
    var conge_payes = parseFloat(data[number]["NBRE CONGE"]);
    var conge_avg = parseFloat(data[number]["CONGE MOYENNE"]);
    var salaire_imposable = 0;
    var abbatement_enfant = parseFloat(data[number]["ENFANT"]);
    var irsa = 0;
    var irsa_brut = 0;
    //Variable stockage
    var gain_base;
    var gain_hs30;
    var gain_hs50;
    var gain_cp;
    var gain_hs100;
    var gain_transport;
    var gain_repas;
    var gain_soir;
    var gain_occ;
    var gain_week;
    var gain_ferie;

    var ret_cnaps_ostie;
    var gain_enf;
    var ret_avance = parseFloat(data[number]["AVANCE"]);
    var gain_remb = parseFloat(data[number]["REMBOURSEMENT"]);
    var ret_remb = parseFloat(data[number]["REAJUSTEMENT"]);
    var ret_base = 0;

    gains_global(salaire, conge_payes, conge_avg);
    heure_supp130(salaire, hs30);
    heure_supp150(salaire, hs50);
    heure_supp200(salaire, hs100);
    transport_func(transport);
    repas_func(repas)
    maj_nuit(salaire, soir);
    nuit_occasionel(salaire, nuit_occ);
    maj_ferie(salaire, ferie);
    maj_weekend(salaire, weekend);
    abbatement_enfant_func(abbatement_enfant);
    var total_gain = gain_base + gain_cp + gain_hs30 + gain_hs50 + gain_transport + gain_repas + rendement + gain_soir + gain_occ + gain_ferie + gain_week + gain_remb;
    cnaps_ostie(total_gain);
    salaire_imposable = total_gain - ret_cnaps_ostie;
    calcul_irsa(salaire_imposable)
    var total_ret = ret_cnaps_ostie + ret_avance + parseFloat(irsa) + ret_remb;
    var final = arrondir(parseFloat((total_gain - total_ret).toFixed(0)));

    //Replace PDF
    const replaceText = async () => {
      try {
        const pdfdoc = await PDFNet.PDFDoc.createFromFilePath("Template.pdf");
        await pdfdoc.initSecurityHandler();
        const replacer = await PDFNet.ContentReplacer.create();
        const page = await pdfdoc.getPage(1);
        //En tête
        var user_get = await UserSchema.findOne({ m_code: m_code });
        //Via base de donnée
        if (user_get) {

        }
        else {
          user_get = {
            matr: "_____________", usuel: "_____________", project: "_____________", cin: "_____________", adresse: "_____________", cnaps_num: "_____________",
            save_at: "_____________", classification: "_____________", contrat: "_____________", leave_taked: "__", remaining_leave: "__", date_fin: "__________"
          }
        }
        await replacer.addString("matr", getting_null(user_get.matr));
        await replacer.addString("name_user", nom);
        await replacer.addString("first_name", getting_null(user_get.usuel));
        await replacer.addString("m_code", getting_null(m_code));
        await replacer.addString("occ_user", getting_null(user_get.project));
        await replacer.addString("cin_user", getting_null(user_get.cin));
        await replacer.addString("adr_user", getting_null(user_get.adresse));
        await replacer.addString("num_cnaps", getting_null(user_get.cnaps_num));
        await replacer.addString("enter_user", getting_null(user_get.save_at));
        await replacer.addString("class_user", getting_null(user_get.classification));
        await replacer.addString("contract_user", getting_null(user_get.contrat));
        await replacer.addString("start_d", moment(start).format("DD/MM/YYYY"));
        await replacer.addString("fin_d", moment(end).format("DD/MM/YYYY"));
        await replacer.addString("fin_c", user_get.date_fin);
        await replacer.addString("nbr_o1", ouvrable);
        await replacer.addString("nbr_o2", ouvres);
        await replacer.addString("nbr_ca", calendaire);
        await replacer.addString("irsa_brut", null_value(irsa_brut.toFixed(2)));
        await replacer.addString("conge_ac", getting_null(parseFloat(user_get.leave_taked)));
        await replacer.addString("an_act", moment().format("YYYY"));
        await replacer.addString("c_r", getting_null(user_get.remaining_leave));
        //Corps
        await replacer.addString("basic_sal", null_value(salaire.toFixed(0)));

        await replacer.addString("somme_gain", null_value(gain_base.toFixed(2)));
        await replacer.addString("ret_base", "");


        await replacer.addString("nbr_hs30", null_value(hs30.toFixed(1)));
        await replacer.addString("gain_hs30", null_value(gain_hs30.toFixed(2)));

        await replacer.addString("nbr_hs50", null_value(hs50.toFixed(1)));
        await replacer.addString("gain_hs50", null_value(gain_hs50.toFixed(2)));

        await replacer.addString("nbr_hs100", null_value(hs100.toFixed(1)));
        await replacer.addString("gain_hs100", null_value(gain_hs100.toFixed(2)));

        await replacer.addString("nbr_trans", null_value(transport.toFixed(1)));
        await replacer.addString("gain_trans", null_value(gain_transport.toFixed(2)));

        await replacer.addString("nbr_repas", null_value(repas.toFixed(1)));
        await replacer.addString("gain_repas", null_value(gain_repas.toFixed(2)));

        await replacer.addString("gain_rend", null_value(rendement.toFixed(2)));

        await replacer.addString("nbr_soir", null_value((soir + nuit_occ).toFixed(1)));
        await replacer.addString("gain_soir", null_value((gain_soir + gain_occ).toFixed(2)));

        await replacer.addString("nbr_ferie", null_value(ferie.toFixed(1)));
        await replacer.addString("gain_ferie", null_value(gain_ferie.toFixed(2)));

        await replacer.addString("nbr_week", null_value(weekend.toFixed(1)));
        await replacer.addString("gain_week", null_value(gain_week.toFixed(2)));

        await replacer.addString("nbr_cp", null_value(conge_payes.toFixed(1)));
        await replacer.addString("gain_cp", null_value(gain_cp.toFixed(2)));

        await replacer.addString("gain_préavis", "");
        await replacer.addString("ret_preavis", "");

        await replacer.addString("ret_cnaps", null_value((ret_cnaps_ostie / 2).toFixed(2)));
        await replacer.addString("ret_ostie", null_value((ret_cnaps_ostie / 2).toFixed(2)));

        await replacer.addString("sal_impos", null_value(salaire_imposable.toFixed(2)));

        await replacer.addString("nbr_enfant", null_value(abbatement_enfant.toFixed(0)));
        await replacer.addString("gain_enf", null_value(gain_enf.toFixed(2)));

        await replacer.addString("ret_irsa", null_value((irsa).toFixed(2)));
        await replacer.addString("ret_avance", null_value(ret_avance.toFixed(0)));

        await replacer.addString("gain_remb", null_value(gain_remb.toFixed(2)));
        await replacer.addString("ret_remb", null_value(ret_remb.toFixed(2)));

        await replacer.addString("tot_gain", null_value(total_gain.toFixed(2)));
        await replacer.addString("tot_ret", null_value(total_ret.toFixed(2)));
        await replacer.addString("sal_fin", null_value(final.toFixed(0)));

        await replacer.process(page);
        var output_path = './public/Paie/' + m_code + ".pdf";

        pdfdoc.save(output_path, PDFNet.SDFDoc.SaveOptions.e_linearized);
        // excel_paie.push([nom,m_code,num_agent,gain_soir.toFixed(2),gain_occ.toFixed(2),gain_week.toFixed(2),gain_ferie.toFixed(2),gain_hs30.toFixed(2),gain_hs50.toFixed(2),gain_hs100.toFixed(2),gain_repas.toFixed(2),gain_transport.toFixed(2),rendement.toFixed(2),])
      } catch (error) {
        res.send("Erreur sur " + data[number]["M-CODE"]);
        number = data.length * 2;
      }
    }
    PDFNet.runWithCleanup(replaceText, "demo:ricardoramandimbisoa@gmail.com:7afedebe02000000000e72b195b776c08a802c3245de93b77462bc8ad6").then(() => {
      list_paie.push([m_code, nom]);
      if (number + 1 < data.length) {
        test_paie(data, number + 1, res, start, end, ouvrable, ouvres, calendaire);
      }
      else {
        if (number == data.length * 2) {

        }
        else {
          PDFNet.shutdown();
          update_opt_paie("y", "add", list_paie);

          res.send(list_paie);
        }

      }
    });
  } catch (error) {
    res.send("Erreur veuillez réessayer ou contactez le développeur");
  }

  //Function
  function getting_null(val) {
    if (val) {
      return val + "";
    }
    else {
      return "aucun(e)"
    }
  }
  function gains_with(salaire_base, conge_payer) {
    return ((hour_calculator(conge_payer) * salaire_base) / 173.33)
  }
  function hour_calculator(nbr) {
    return 6 * nbr;
  }
  function gains_global(salaire_base, conge_payer, conge_avg) {
    if (conge_payer == 0 || conge_avg == 0) {
      gain_base = salaire_base;
      gain_cp = 0;
      return salaire_base
    }
    else {
      gain_cp = ((conge_avg / 30) * conge_payer);
      gain_base = salaire_base - gains_with(salaire_base, conge_payer);
      return salaire_base - gains_with(salaire_base, conge_payer);
    }
  }
  function maj_nuit(salaire_base, nuit) {
    gain_soir = (((salaire_base / 173.33) * 30) / 100) * nuit;
    return (((salaire_base / 173.33) * 30) / 100) * nuit
  }
  function nuit_occasionel(salaire_base, noc) {
    gain_occ = (((salaire_base / 173.33) * 50) / 100) * noc;
    return (((salaire_base / 173.33) * 50) / 100) * noc
  }
  function maj_weekend(salaire_base, week) {
    gain_week = (((salaire_base / 173.33) * 50) / 100) * week;
    return (((salaire_base / 173.33) * 50) / 100) * week
  }
  function maj_ferie(salaire_base, ferie) {
    gain_ferie = (((salaire_base / 173.33) * 100) / 100) * ferie
    return (((salaire_base / 173.33) * 100) / 100) * ferie
  }
  function heure_supp130(salaire_base, supp) {
    gain_hs30 = (((salaire_base / 173.33) * 130) / 100) * supp;
    return (((salaire_base / 173.33) * 130) / 100) * supp
  }
  function heure_supp150(salaire_base, supp) {
    gain_hs50 = (((salaire_base / 173.33) * 150) / 100) * supp
    return (((salaire_base / 173.33) * 150) / 100) * supp
  }
  function heure_supp200(salaire_base, supp) {
    gain_hs100 = (((salaire_base / 173.33) * 100) / 100) * supp
    return (((salaire_base / 173.33) * 100) / 100) * supp
  }
  function repas_func(nbr_repas) {
    gain_repas = nbr_repas * 3500;
    return nbr_repas * 3500
  }
  function transport_func(nbr_transport) {
    gain_transport = nbr_transport * 2400
    return nbr_transport * 2400
  }
  function cnaps_ostie(total_salaire) {
    ret_cnaps_ostie = (total_salaire / 100) * 2;
    return total_salaire / 100
  }
  function abbatement_enfant_func(ab_enf) {
    gain_enf = ab_enf * 2000;
    return ab_enf * 2000;;
  }
  function calcul_irsa(base_imp) {
    var t1 = 2500;
    var t2 = 10000;
    var t3 = 15000;

    if (base_imp <= 350000) {
      irsa = 3000 - gain_enf;
      irsa_brut = 3000;
      return 3000;
    }
    else if (base_imp >= 350001 && base_imp <= 400000) {
      irsa = (((base_imp - 350000) * 0.05)) - gain_enf;
      irsa_brut = ((base_imp - 350001) * 0.05);
      return ((base_imp - 350000) * 0.05)
    }
    else if (base_imp >= 400001 && base_imp <= 500000) {
      irsa = ((((base_imp - 400000) * 0.1) + t1)) - gain_enf;
      irsa_brut = (((base_imp - 400000) * 0.1) + t1);
      return (((base_imp - 400000) * 0.1) + t1)
    }
    else if (base_imp >= 500001 && base_imp <= 600000) {
      irsa = ((((base_imp - 500000) * 0.15) + t1 + t2)) - gain_enf;
      irsa_brut = (((base_imp - 500000) * 0.15) + t1 + t2);
      return (((base_imp - 500000) * 0.15) + t1 + t2)
    }
    else {
      irsa = ((((base_imp - 600000) * 0.2) + t1 + t2 + t3)) - gain_enf;
      irsa_brut = (((base_imp - 600000) * 0.2) + t1 + t2 + t3);
      return (((base_imp - 600000) * 0.2) + t1 + t2 + t3)
    }
  }
  function numberWithCommas(x) {
    return x.toString().replace(/\B(?=(\d{3})+(?!\d))/g, " ");
  }
  function null_value(given) {
    if (given == 0 || given == "0") {
      return " "
    }
    else {
      return numberWithCommas(parseFloat(given))
    }
  }
  function arrondir(nbr) {
    var string_number = nbr + "";
    var last_num = parseFloat(string_number[string_number.length - 2] + "" + string_number[string_number.length - 1]);
    if (last_num > 50) {
      return nbr + (100 - last_num)
    }
    else if (last_num < 50) {
      return nbr - last_num
    }
    else {
      return nbr
    }

  }

}
// Empty folder
routeExp.route("/empty").get(function (req, res) {
  extra_fs.emptyDirSync("./public/Paie");
  update_opt_paie("n", "delete", "");
  res.redirect("/fiche")
})
function update_opt_paie(opt_value, action, list_paie) {
  mongoose
    .connect(
      "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
      {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
      }
    )
    .then(async () => {
      if (action == "add") {
        await OptSchema.findOneAndUpdate({ _id: "636247a2c1f6301f15470344" }, { paie_generated: opt_value, list_paie: JSON.stringify(list_paie) });
      }
      else {
        await OptSchema.findOneAndUpdate({ _id: "636247a2c1f6301f15470344" }, { paie_generated: opt_value, list_paie: "" });
      }

    })
}
//Fin Paie code


//getuser
routeExp.route("/getuser_leave").post(async function (req, res) {
  var id = req.body.id;
  mongoose
    .connect(
      "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
      {
        useUnifiedTopology: true,
        UseNewUrlParser: true,
      }
    )
    .then(async () => {
      var user = await UserSchema.findOne({ _id: id });
      var conge = await LeaveSchema.findOne({ m_code: user.m_code, status: { $ne: "Terminée" } });
      res.send(user.first_name + ";" + user.last_name + ";" + user.shift + ";" + user._id + ";" + time_passed(user.save_at) + ";" + user.remaining_leave + ";" + user.leave_taked + ";" + JSON.stringify(conge));
    });
});

//Filter leave
routeExp.route("/monthly_leave").post(async function (req, res) {
  session = req.session;
  monthly_leave = [];
  datestart_leave = moment(req.body.datestart).format("YYYY-MM-DD");
  dateend_leave = moment(req.body.dateend).format("YYYY-MM-DD");
  if (session.occupation_a == "Admin") {
    mongoose
      .connect(
        "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
        {
          useUnifiedTopology: true,
          UseNewUrlParser: true,
        }
      )
      .then(async () => {
        maternity = await LeaveSchema.find({ type: "Congé de maternité ( rien à deduire )", status: "en cours" });
        var next_date = datestart_leave;
        var last_month = moment(next_date).add(-1, "months").format("YYYY-MM-DD");
        while (dateend_leave != moment(next_date).add(-1, "days").format("YYYY-MM-DD")) {
          var leave_spec = await LeaveSchema.find({ date_start: next_date, status: "Terminée" });
          var leave_spec2 = await LeaveSchema.find({ date_end: next_date, date_start: { '$regex': last_month.split("-")[0] + "-" + last_month.split("-")[1], '$options': 'i' } })
          monthly_leave.push(leave_spec);
          monthly_leave.push(leave_spec2);
          next_date = moment(next_date).add(1, "days").format("YYYY-MM-DD");
        }
        for (i = 1; i < monthly_leave.length; i++) {
          for (d = 0; d < monthly_leave[i].length; d++) {
            monthly_leave[0].push(monthly_leave[i][d]);
          }
        }
        monthly_leave = monthly_leave[0];
        res.send("Ok");
      })

  }
  else {
    res.send("error");
  }
})
//leave delete
routeExp.route("/delete_leave").post(async function (req, res) {
  session = req.session;
  var id = req.body.id;
  if (session.occupation_a == "Admin") {
    mongoose
      .connect(
        "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
        {
          useUnifiedTopology: true,
          UseNewUrlParser: true,
        }
      )
      .then(async () => {
        var leave_to_delete = await LeaveSchema.findOne({ _id: id });
        var all_leave_toupdate = await LeaveSchema.find({});
        if (leave_to_delete.type == "Congé Payé ( rien à deduire )") {
          await UserSchema.findOneAndUpdate({ m_code: leave_to_delete.m_code }, { $inc: { remaining_leave: leave_to_delete.duration, leave_taked: leave_to_delete.duration } });
          var begin = false;
          var val_to_add = 0;
          for (l = 0; l < all_leave_toupdate.length; l++) {
            if (all_leave_toupdate[l]._id + "" == leave_to_delete._id + "") {
              begin = true;
              val_to_add = leave_to_delete.duration;
            }
            if (begin) {
              await LeaveSchema.findOneAndUpdate({ _id: all_leave_toupdate[l]._id }, { $inc: { rest: val_to_add } });
            }
          }
        }
        if (leave_to_delete.status == "en cours") {
          await UserSchema.findOneAndUpdate({ m_code: leave_to_delete.m_code }, { act_stat: "LEFTING" });
        }
        await LeaveSchema.findOneAndDelete({ _id: id });
        res.send("Ok");
      })
  } else {
    res.send("error");
  }
})
routeExp.route("/leave_report").post(async function (req, res) {
  session = req.session;
  if (session.occupation_a == "Admin") {
    var newsheet_leave = ExcelFile.utils.book_new();
    var m_leave = [];
    var leave_report = [];
    var merging = [];
    newsheet_leave.Props = {
      Title: "Rapport de congé",
      Subject: "Rapport de congé",
      Author: "Solumada",
    };
    leave_report.push(["Les absences et Congés du " + moment(datestart_leave).format("DD/MM/YYYY") + " au " + moment(dateend_leave).format("DD/MM/YYYY"), "", "", "", "", "", ""]);
    var months = moment(datestart_leave).locale("Fr").format("MMMM YYYY");
    leave_report.push(["Numbering agent", "M-CODE", "Nombre de jours à payer et / ou de déduction sur salaire " + months, "", "", "", "Motifs - observations ou remarques"]);
    leave_report.push(["", "", "Congés payer", "Permission exceptionelle", "Consultation ou Repos\n maladie à payer", "Congés sans solde: déduction sur salaire", ""]);
    newsheet_leave.SheetNames.push("Conge " + months);
    for (i = 0; i < monthly_leave.length; i++) {
      if (m_leave.includes(monthly_leave[i].m_code)) {

      }
      else {
        m_leave.push(monthly_leave[i].m_code);
      }
    }
    m_leave = m_leave.sort();
    for (m = 0; m < m_leave.length; m++) {
      var count = 0;
      for (i = 0; i < monthly_leave.length; i++) {

        if (monthly_leave[i].m_code == m_leave[m]) {
          count++;
          if (monthly_leave[i].type.includes("Congé de maternité")) {
          }
          else {
            leave_report.push([monthly_leave[i].num_agent, monthly_leave[i].m_code, conge_payer(monthly_leave[i].type, monthly_leave[i].duration), permission_exceptionelle(monthly_leave[i].type, monthly_leave[i].duration), repos_maladie(monthly_leave[i].type, monthly_leave[i].duration), sans_solde(monthly_leave[i].type, monthly_leave[i].duration), monthly_leave[i].duration + " jour(s) de " + monthly_leave[i].type + date_rendered(monthly_leave[i].date_start, monthly_leave[i].date_end)]);
          }

        }

      }
      merging.push([m, count]);
    }
    leave_report.push(["", "", "", "", "", "", ""]);
    leave_report.push(["", "", "", "", "", "", ""]);
    for (mat = 0; mat < maternity.length; mat++) {
      leave_report.push([maternity[mat].num_agent, maternity[mat].m_code, "Congé de maternité depuis " + moment(maternity[mat].date_start).format("DD/MM/YYYY") + " jusqu'au " + moment(maternity[mat].date_end).format("DD/MM/YYYY")]);
    }
    leave_report.push(["", "", ""]);
    ws_leave = ExcelFile.utils.aoa_to_sheet(leave_report);
    ws_leave["!cols"] = [
      { wpx: 100 },
      { wpx: 60 },
      { wpx: 285 },
      { wpx: 150 },
      { wpx: 210 },
      { wpx: 210 },
      { wpx: 400 },
    ];
    var merge = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: 6 } },
      { s: { r: 1, c: 0 }, e: { r: 2, c: 0 } },
      { s: { r: 1, c: 1 }, e: { r: 2, c: 1 } },
      { s: { r: 1, c: 2 }, e: { r: 1, c: 5 } },
      { s: { r: 1, c: 6 }, e: { r: 2, c: 6 } }
    ];
    var last = 0;
    var field = 0;
    for (mr = 0; mr < merging.length; mr++) {
      if (merging[mr][1] > 1) {
        merge.push({ s: { r: merging[mr][0] + 3 + last, c: 0 }, e: { r: merging[mr][0] + 3 + last + merging[mr][1] - 1, c: 0 } });
        merge.push({ s: { r: merging[mr][0] + 3 + last, c: 1 }, e: { r: merging[mr][0] + 3 + last + merging[mr][1] - 1, c: 1 } });
        last = last + merging[mr][1] - 1;
        field++;
      }
    }
    ws_leave["!merges"] = merge;
    style3(last, maternity.length, field);
    newsheet_leave.Sheets["Conge " + months] = ws_leave;
    session.filename = "Rapport congé " + months + ".xlsx";
    ExcelFile.writeFile(newsheet_leave, session.filename);
    res.send("Ok");
  }
  else {
    res.send("error");
  }

})
//Leave restants
routeExp.route("/leave_left").post(async function (req, res) {
  session = req.session;
  if (session.occupation_a == "Admin") {
    var newsheet_left = ExcelFile.utils.book_new();
    var leave_left = [];
    var months = moment(datestart_leave).locale("Fr").format("MMMM YYYY");
    newsheet_left.Props = {
      Title: "Congé restants",
      Subject: "Congé restants",
      Author: "Solumada",
    };
    newsheet_left.SheetNames.push("Conge " + months);
    mongoose
      .connect(
        "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
        {
          useUnifiedTopology: true,
          UseNewUrlParser: true,
        }
      )
      .then(async () => {
        leave_left.push(["CONGES PAYES ARRETES DU MOIS DE " + months, "", "", "", "", "", ""]);
        leave_left.push(["Nom & Prénom", "Numbering Agent", "M-code", "Embauche", "Projet(s)", "Congés restants", "Congés ouvert"]);
        var data_leave_left = await UserSchema.find({ occupation: "User" }).sort({
          "first_name": 1,
        });
        for (dl = 0; dl < data_leave_left.length; dl++) {
          leave_left.push([data_leave_left[dl].first_name + " " + data_leave_left[dl].last_name, data_leave_left[dl].num_agent, data_leave_left[dl].m_code, moment(data_leave_left[dl].save_at).format("DD/MM/YYYY"), data_leave_left[dl].project, data_leave_left[dl].leave_taked, data_leave_left[dl].remaining_leave]);
        }
        leave_left.push(["", "", "", "", "", "", ""]);
        ws_left = ExcelFile.utils.aoa_to_sheet(leave_left);
        ws_left["!cols"] = [
          { wpx: 325 },
          { wpx: 125 },
          { wpx: 85 },
          { wpx: 125 },
          { wpx: 300 },
          { wpx: 125 },
          { wpx: 125 },
        ];
        const merge = [
          { s: { r: 0, c: 0 }, e: { r: 0, c: 6 } },
        ];
        ws_left["!merges"] = merge;
        style4(leave_left);
        newsheet_left.Sheets["Conge " + months] = ws_left;
        session.filename = "CONGE PAYES DU MOIS " + months + ".xlsx";
        ExcelFile.writeFile(newsheet_left, session.filename);
        res.send("Ok");
      })
  }
  else {
    res.send("error");
  }
})
function conge_payer(motif, number) {
  if (motif.includes('Congé Payé')) {
    return number;
  }
  else {
    return "";
  }
}
function permission_exceptionelle(motif, number) {
  if (motif.includes('Permission exceptionelle')) {
    return number;
  }
  else {
    return "";
  }
}
function repos_maladie(motif, number) {
  if (motif.includes('Repos Maladie')) {
    return number;
  }
  else {
    return "";
  }
}
function sans_solde(motif, number) {
  if (motif.includes('Absent') || motif.includes('Mise a Pied') || motif.includes('Congé sans solde')) {
    return number;
  }
  else {
    return "";
  }
}
function date_rendered(d1, d2) {
  if (d1 == d2) {
    return " le " + moment(d1).format("DD/MM/YYYY");
  }
  else {
    return " du " + moment(d1).format("DD/MM/YYYY") + " au " + moment(d2).format("DD/MM/YYYY");
  }
}
var individual_live = [];
//Leave stat 
routeExp.route("/absence_stat").post(async function (req, res) {
  session = req.session;
  var id = req.body.id;
  var an = req.body.an;

  var january; var february; var march; var april; var may; var june; var july; var august; var september; var october; var november; var december;
  var total_jouissance = 0; var total_conge_paye = 0; var total_sans_solde = 0; var total_ostie = 0; var total_permission_exceptionelle = 0; var total_absence = 0;
  var another = ["Mise a Pied ( a déduire sur salaire )", "Congé de maternité ( rien à deduire )", "Absent ( a déduire sur salaire )"];
  var render_month = ["Janvier", "Février", "Mars", "Avril", "May", "Juin", "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"];
  if (session.occupation_a == "Admin") {
    mongoose
      .connect(
        "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
        {
          useUnifiedTopology: true,
          UseNewUrlParser: true,
        }
      )
      .then(async () => {
        var user = await UserSchema.findOne({ _id: id });
        var newsheets_individual = ExcelFile.utils.book_new();
        newsheets_individual.Props = {
          Title: "Etat absence",
          Subject: "Etat absence",
          Author: "Solumada",
        };
        newsheets_individual.SheetNames.push(user.last_name);
        january = await LeaveSchema.find({ m_code: user.m_code, date_start: { '$regex': an + "-01", '$options': 'i' } });
        february = await LeaveSchema.find({ m_code: user.m_code, date_start: { '$regex': an + "-02", '$options': 'i' } });
        march = await LeaveSchema.find({ m_code: user.m_code, date_start: { '$regex': an + "-03", '$options': 'i' } });
        april = await LeaveSchema.find({ m_code: user.m_code, date_start: { '$regex': an + "-04", '$options': 'i' } });
        may = await LeaveSchema.find({ m_code: user.m_code, date_start: { '$regex': an + "-05", '$options': 'i' } });
        june = await LeaveSchema.find({ m_code: user.m_code, date_start: { '$regex': an + "-06", '$options': 'i' } });
        july = await LeaveSchema.find({ m_code: user.m_code, date_start: { '$regex': an + "-07", '$options': 'i' } });
        august = await LeaveSchema.find({ m_code: user.m_code, date_start: { '$regex': an + "-08", '$options': 'i' } });
        september = await LeaveSchema.find({ m_code: user.m_code, date_start: { '$regex': an + "-09", '$options': 'i' } });
        october = await LeaveSchema.find({ m_code: user.m_code, date_start: { '$regex': an + "-10", '$options': 'i' } });
        november = await LeaveSchema.find({ m_code: user.m_code, date_start: { '$regex': an + "-11", '$options': 'i' } });
        december = await LeaveSchema.find({ m_code: user.m_code, date_start: { '$regex': an + "-12", '$options': 'i' } });
        individual_live.push([user.first_name, user.last_name, "", "", "", "", "", "", "", "", "", ""]);
        individual_live.push(["Année", an, "", "M-code : ", user.m_code, "Numéro d'agent : ", user.num_agent, "", "", "", "", ""]);
        individual_live.push(["Date", "Période sollicitée", "", "Droit", "", "", "Congé Payé", "Congé sans solde", "OSTIE", "Pérmission Exceptionelle", "Absent et autres", "Observations"]);
        individual_live.push(["", "Début", "Fin", "Acquisition", "Jouissance", "Reste", "", "", "", "", "", ""]);
        var maternite = [];
        var months_to_write = [january, february, march, april, may, june, july, august, september, october, november, december];
        for (i = 0; i < 12; i++) {
          for (m = 0; m < months_to_write[i].length; m++) {
            if (months_to_write[i][m].type == "Congé de maternité ( rien à deduire )") {
              maternite = ["Congé Maternité", "Congé Maternité", "Congé Maternité"];
            }
            individual_live.push([convert_date1(months_to_write[i][m].date_start), convert_date2(months_to_write[i][m].date_start), convert_date2(months_to_write[i][m].date_end)
              , "", jouissance(months_to_write[i][m].type, months_to_write[i][m].duration), months_to_write[i][m].rest,
            another_conge_payer(months_to_write[i][m].type, months_to_write[i][m].duration), another_sans_solde(months_to_write[i][m].type, months_to_write[i][m].duration), ostie(months_to_write[i][m].type, months_to_write[i][m].duration),
            another_permission_exceptionelle(months_to_write[i][m].type, months_to_write[i][m].duration), absence(months_to_write[i][m].type, months_to_write[i][m].duration), write_maternity(maternite, months_to_write[i][m].type)]);
          }
          individual_live.push(["Fin " + render_month[i], "", "", "2.5", "", "", "", "", "", "", "", write_maternity(maternite, "")]);
          maternite.pop();
        }
        individual_live.push(["Fin " + "Année", "-", "-", "30", total_jouissance, user.remaining_leave + 30, total_conge_paye, total_sans_solde, total_ostie, total_permission_exceptionelle, total_absence, "Fin"]);
        ws_individual = ExcelFile.utils.aoa_to_sheet(individual_live);
        ws_individual["!cols"] = [
          { wpx: 90 },
          { wpx: 90 },
          { wpx: 90 },
          { wpx: 90 },
          { wpx: 90 },
          { wpx: 90 },
          { wpx: 90 },
          { wpx: 125 },
          { wpx: 90 },
          { wpx: 125 },
          { wpx: 125 },
          { wpx: 225 },
        ];
        const merge = [
          { s: { r: 2, c: 0 }, e: { r: 3, c: 0 } },
          { s: { r: 2, c: 1 }, e: { r: 2, c: 2 } },
          { s: { r: 2, c: 3 }, e: { r: 2, c: 5 } },
          { s: { r: 2, c: 6 }, e: { r: 3, c: 6 } },
          { s: { r: 2, c: 7 }, e: { r: 3, c: 7 } },
          { s: { r: 2, c: 8 }, e: { r: 3, c: 8 } },
          { s: { r: 2, c: 9 }, e: { r: 3, c: 9 } },
          { s: { r: 2, c: 10 }, e: { r: 3, c: 10 } }
        ];
        ws_individual["!merges"] = merge;
        style5();
        newsheets_individual.Sheets[user.last_name] = ws_individual;
        session.filename = user.last_name + ".xlsx";
        ExcelFile.writeFile(newsheets_individual, session.filename);
        individual_live = [];
        res.send("Ok");
      });
  }
  else {
    res.redirect("/");
  }
  function jouissance(type, value) {
    if (type == "Congé Payé ( rien à deduire )") {
      total_jouissance = total_jouissance + value;
      return value;
    }
    else {
      return 0;
    }

  }
  function write_maternity(matern, obs) {
    if (matern[0]) {
      return matern[0];
    }
    else {
      return obs;
    }
  }
  function another_conge_payer(type, value) {
    if (type == "Congé Payé ( rien à deduire )") {
      total_conge_paye = total_conge_paye + value;
      return value;
    }
    else {
      return "";
    }
  }
  function another_sans_solde(type, value) {
    if (type == "Congé sans solde ( a déduire sur salaire )") {
      total_sans_solde = total_sans_solde + value;
      return value;
    }
    else {
      return "";
    }
  }
  function ostie(type, value) {
    if (type == "Repos Maladie ( rien à deduire )") {
      total_ostie = total_ostie + value;
      return value;
    }
    else {
      return "";
    }
  }
  function another_permission_exceptionelle(type, value) {
    if (type == "Permission exceptionelle ( rien à deduire )") {
      total_permission_exceptionelle = total_permission_exceptionelle + value;
      return value;
    }
    else {
      return "";
    }
  }
  function absence(type, value) {
    if (another.includes(type)) {
      total_absence = total_absence + value;
      return value;
    }
    else {
      return "";
    }
  }
})
function convert_date1(d1) {
  return moment(d1).format("DD/MM/YYYY");
}
function convert_date2(d2) {
  return moment(d2).locale("Fr").format("DD MMMM");
}

routeExp.route("/session_end").get(async function (req, res) {
  res.render("block.html");
})
//Generate excel file
routeExp.route("/generate").post(async function (req, res) {
  session = req.session;
  if (session.occupation_a == "Admin") {
    var newsheet = ExcelFile.utils.book_new();
    newsheet.Props = {
      Title: "Timesheets",
      Subject: "Logged Time",
      Author: "Solumada",
    };
    newsheet.SheetNames.push("TOUS LES UTILISATEURS");
    mongoose
      .connect(
        "mongodb+srv://Rica:ryane_jarello5@cluster0.z3s3n.mongodb.net/Pointage?retryWrites=true&w=majority",
        {
          useUnifiedTopology: true,
          UseNewUrlParser: true,
        }
      )
      .then(async () => {

        var all_employes = [];
        for (i = 0; i < data_desired.datatowrite.length; i++) {
          if (all_employes.includes(data_desired.datatowrite[i].m_code)) {

          }
          else {
            all_employes.push(data_desired.datatowrite[i].m_code);
          }
          all_employes = all_employes.sort();
        }
        all_datas.push([
          "RAPPORT GLOBALE",
          "",
          "",
          "",
          "",
          "",
        ]);
        all_datas.push([
          "",
          "",
          "",
          "",
          "",
          "",
        ]);
        all_datas.push([
          "Nom & Prenom",
          "M-code",
          "Totale heure travail",
          "Totale Retard",
          "Totale absence",
          "Totale congé",
        ]);
        for (e = 0; e < all_employes.length; e++) {
          var name_user = await StatusSchema.findOne({ m_code: all_employes[e] });
          data.push([
            "FEUILLE DE => " + name_user.nom,
            "",
            "",
            "",
            "",
            "",
            ""
          ]);
          data.push([
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            ""
          ]);
          data.push([
            "M-code",
            "Numéro Agent",
            "Date",
            "Locaux",
            "Début",
            "Fin",
            "Heure calculer",
            "Heure choisi",
          ]);
          generate_excel(data_desired.datatowrite, data_desired.datalate, data_desired.dataabsence, data_desired.dataleave, all_employes[e]);
          if (newsheet.SheetNames.includes(all_employes[e])) {
          } else {
            newsheet.SheetNames.push(all_employes[e]);
          }
          newsheet.Sheets[all_employes[e]] = ws;
          hours = 0;
          minutes = 0;
          data = [];
        }
        global_Report(all_datas);
        newsheet.Sheets["TOUS LES UTILISATEURS"] = ws;
        all_datas = [];
        if (newsheet.SheetNames.length != 0) {
          if (all_employes.length <= 1) {
            session.filename = "N°" + num_file + " " + all_employes[0] + ".xlsx";
            num_file++;
          }
          else {
            session.filename = "N°" + num_file + " Feuille_de_temps.xlsx";
            num_file++;
          }
          ExcelFile.writeFile(newsheet, session.filename);
          delete filtrage.searchit;
          delete filtrage.date;
          delete filtrage.search;
          data_desired.datatowrite = await StatusSchema.find({});
        }
        res.send("Done");
      });
  } else {
    res.redirect("/");
  }
});
routeExp.route("/download").get(async function (req, res) {
  session = req.session;
  if (session.occupation_a == "Admin") {
    res.download(session.filename, function (err) {
      fs.unlink(session.filename, function (err) {
        if (err) {
          console.error(err);
        }
        console.log('File has been Deleted');
      });
    });
  }
});
//logout
routeExp.route("/exit_a").get(function (req, res) {
  session = req.session;
  session.occupation_a = null;
  res.redirect("/");
});
routeExp.route("/exit_tl").get(function (req, res) {
  session = req.session;
  session.occupation_tl = null;
  res.redirect("/");
});
function htmlVerification(code) {
  return (
    "<center><h1>VOTRE CODE D'AUTHENTIFICATION</h1>" +
    "<h3 style='width:250px;font-size:50px;padding:8px;background-color: rgba(87,184,70, 0.8); color:white'>" +
    code +
    "<h3></center>"
  );
}
function htmlRender(username, password) {
  var html =
    "<center><h1>Solumada Authentification</h1>" +
    '<table border="1" style="border-collapse:collapse;width:25%;border-color: lightgrey;">' +
    '<thead style="background-color: rgba(87,184,70, 0.8);color:white;font-weight:bold;height: 50px;">' +
    "<tr>" +
    '<td align="center">Nom utilisateur</td>' +
    '<td align="center"ot de passe</td>' +
    "</tr>" +
    "</thead>" +
    '<tbody style="height: 50px;">' +
    "<tr>" +
    '<td align="center">' +
    username +
    "</td>" +
    '<td align="center">' +
    password +
    "</td>" +
    "</tr>" +
    "</tbody>" +
    "</table>";
  return html;
}
function randomPassword() {
  var code = "";
  let v = "abcdefghijklmnopqrstuvwxyz0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ!é&#";
  for (let i = 0; i < 8; i++) {
    // 6 characters
    let char = v.charAt(Math.random() * v.length - 1);
    code += char;
  }
  return code;
}
function sendEmail(receiver, subject, text) {
  var mailOptions = {
    from: "Timesheets Optimum solution",
    to: receiver,
    subject: subject,
    html: text,
  };

  transporter.sendMail(mailOptions, function (error, info) {
    if (error) {
      console.log(error);
    } else {
      console.log("Email sent: " + info.response);
    }
  });
}
//Function Random code for verification
function randomCode() {
  var code = "";
  let v = "012345678";
  for (let i = 0; i < 6; i++) {
    // 6 characters
    let char = v.charAt(Math.random() * v.length - 1);
    code += char;
  }
  return code;
}
function calcul_timediff(startTime, endTime) {
  startTime = moment(startTime, "HH:mm:ss a");
  endTime = moment(endTime, "HH:mm:ss a");
  var duration = moment.duration(endTime.diff(startTime));
  //duration in hours
  hours += parseInt(duration.asHours());

  // duration in minutes
  minutes += parseInt(duration.asMinutes()) % 60;
  while (minutes > 60) {
    hours += 1;
    minutes = minutes - 60;
  }
  if (hours < 0 || minutes < 0) {
    hours = hours + 24;
    if (minutes != 0) {
      hours = hours - 1;
      minutes = minutes + 60;
    }
  }
  return parseInt(duration.asHours()) + "H " + (parseInt(duration.asMinutes()) % 60) + "MN";
}
function hour_diff(startday, endday) {
  startday = moment(startday, "HH:mm:ss a");
  endday = moment(endday, "HH:mm:ss a");
  var duration = moment.duration(endday.diff(startday));
  return parseInt(duration.asHours());
}
function convert_to_hour(mins) {
  var hc = 0;
  while (mins > 60) {
    hc += 1;
    mins = mins - 60;
  }
  return hc + "," + mins;
}
function difference_year(starting) {
  var startings = moment(moment(starting)).format("YYYY-MM-DD");
  var nows = moment(moment().format("YYYY-MM-DD"), "YYYY-MM-DD");
  var duration = moment.duration(nows.diff(startings));
  var years = duration.years();
  return years;
}
function time_passed(starting) {
  var startings = moment(moment(starting)).format("YYYY-MM-DD");
  var nows = moment(moment().format("YYYY-MM-DD"), "YYYY-MM-DD");
  var duration = moment.duration(nows.diff(startings));
  var years = duration.years();
  var months = duration.months();
  var days = parseInt(moment().format("DD")) - parseInt(moment(starting).format("DD"))
  var tp = years + " an(s) " + months + " mois " + days + " jour(s)";
  return tp;
}
function date_diff(starting, ending) {
  var startings = moment(moment(starting)).format("YYYY-MM-DD");
  var endings = moment(ending, "YYYY-MM-DD");
  var duration = moment.duration(endings.diff(startings));
  var dayl = duration.asDays();
  return parseInt(dayl.toFixed(0));
}
function calcul_retard(regular, arrived) {
  var time = 0;
  var lh = 0;
  var lm = 0;
  regular = moment(regular, "HH:mm:ss a");
  arrived = moment(arrived, "HH:mm:ss a");
  var duration = moment.duration(arrived.diff(regular));
  //duration in hours
  lh = parseInt(duration.asHours());
  // duration in minutes
  lm = parseInt(duration.asMinutes()) % 60;
  while (lm > 60) {
    lh += 1;
    lm = lm - 60;
  }
  lh = lh * 60;
  time = lh + lm;
  return time;
}
function style() {
  var cellule = ["A", "B", "C", "D", "E", "F", "G", "H"];
  for (c = 0; c < cellule.length; c++) {
    for (i = 1; i <= data.length; i++) {
      if (ws[cellule[c] + "" + i]) {
        if (i == 1 || i == 2) {
          ws[cellule[c] + "" + i].s = {
            font: {
              name: "Segoe UI Black",
              bold: true,
              color: { rgb: "398C39" },
            },
            alignment: {
              vertical: "center",
              horizontal: "center"
            },
          };
        }
        else if (i == 3) {
          ws[cellule[c] + "" + i].s = {
            fill: {
              patternType: "solid",
              fgColor: { rgb: "398C39" },
              bgColor: { rgb: "398C39" },
            },
            font: {
              name: "Segoe UI Black",
              bold: true,
              color: { rgb: "F5F5F5" }
            },
            border: {
              left: { style: "hair" },
              right: { style: "hair" },
              top: {
                style: "hair",
                bottom: { style: "hair" },
              },
            },
            alignment: {
              vertical: "center",
              horizontal: "center"
            },
          };
        }
        else {
          ws[cellule[c] + "" + i].s = {
            font: {
              name: "Verdana",
              color: { rgb: "777777" }
            },
            border: {
              left: { style: "hair" },
              right: { style: "hair" },
              top: {
                style: "hair",
                bottom: { style: "hair" },
              },
            },
            alignment: {
              vertical: "center",
              horizontal: "center"
            },
          };
        }
      }
    }
  }
}
function style3(last, maternity, field) {
  var cellule = ["A", "B", "C", "D", "E", "F", "G"];
  for (c = 0; c < cellule.length; c++) {
    for (i = 1; i <= monthly_leave.length + last + maternity + field; i++) {
      if (ws_leave[cellule[c] + "" + i]) {
        if (i == 1) {
          ws_leave[cellule[c] + "" + i].s = {
            font: {
              name: "Calibri",
              bold: true,
              sz: 18
            },
            alignment: {
              vertical: "center",
              horizontal: "center"
            },
          };
        }
        else if (i == 2 || i == 3) {
          ws_leave[cellule[c] + "" + i].s = {
            fill: {
              patternType: "solid",
              fgColor: { rgb: "FFFFFF" },
              bgColor: { rgb: "FFFFFF" },
            },
            font: {
              name: "Calibri",
              sz: 11,
              bold: true,
            },
            border: {
              left: { style: "thin" },
              right: { style: "thin" },
              top: {
                style: "thin",
                bottom: { style: "thin" },
              },
            },
            alignment: {
              vertical: "center",
              horizontal: "center"
            },
          };
        }
        else {
          ws_leave[cellule[c] + "" + i].s = {
            font: {
              name: "Calibri",
              sz: 11
            },
            border: {
              left: { style: "thin" },
              right: { style: "thin" },
              top: {
                style: "thin",
                bottom: { style: "thin" },
              },
            },
            alignment: {
              vertical: "center",
              horizontal: "center"
            },
          };
        }
      }
    }
  }
}
function style4(leave_left) {
  var cellule = ["A", "B", "C", "D", "E", "F", "G"];
  for (c = 0; c < cellule.length; c++) {
    for (i = 1; i <= leave_left.length; i++) {
      if (ws_left[cellule[c] + "" + i]) {
        if (i == 1) {
          ws_left[cellule[c] + "" + i].s = {
            font: {
              name: "Calibri",
              bold: true,
              sz: 18
            },
            alignment: {
              vertical: "center",
              horizontal: "center"
            },
          };
        }
        else if (i == 2) {
          ws_left[cellule[c] + "" + i].s = {
            fill: {
              patternType: "solid",
              fgColor: { rgb: "FFFFFF" },
              bgColor: { rgb: "FFFFFF" },
            },
            font: {
              name: "Calibri",
              sz: 14,
              bold: true,
            },
            border: {
              left: { style: "thin" },
              right: { style: "thin" },
              top: {
                style: "thin",
                bottom: { style: "thin" },
              },
            },
            alignment: {
              vertical: "center",
              horizontal: "center"
            },
          };
        }
        else {
          ws_left[cellule[c] + "" + i].s = {
            font: {
              name: "Calibri",
              sz: 14
            },
            border: {
              left: { style: "thin" },
              right: { style: "thin" },
              top: {
                style: "thin",
                bottom: { style: "thin" },
              },
            },
            alignment: {
              vertical: "center",
              horizontal: "center"
            },
          };
        }
      }
    }
  }
}
function style2() {
  var cellule = ["A", "B", "C", "D", "E", "F", "G"];
  for (c = 0; c < cellule.length; c++) {
    for (i = 1; i <= all_datas.length; i++) {
      if (ws[cellule[c] + "" + i]) {
        if (i == 1 || i == 2) {
          ws[cellule[c] + "" + i].s = {
            font: {
              name: "Segoe UI Black",
              bold: true,
              color: { rgb: "398C39" },
            },
            alignment: {
              vertical: "center",
              horizontal: "center"
            },
          };
        }
        else if (i == 3) {
          ws[cellule[c] + "" + i].s = {
            fill: {
              patternType: "solid",
              fgColor: { rgb: "398C39" },
              bgColor: { rgb: "398C39" },
            },
            font: {
              name: "Segoe UI Black",
              bold: true,
              color: { rgb: "F5F5F5" }
            },
            border: {
              left: { style: "hair" },
              right: { style: "hair" },
              top: {
                style: "hair",
                bottom: { style: "hair" },
              },
            },
            alignment: {
              vertical: "center",
              horizontal: "center"
            },
          };
        }
        else {
          ws[cellule[c] + "" + i].s = {
            font: {
              name: "Verdana",
              color: { rgb: "777777" }
            },
            border: {
              left: { style: "hair" },
              right: { style: "hair" },
              top: {
                style: "hair",
                bottom: { style: "hair" },
              },
            },
            alignment: {
              vertical: "center",
              horizontal: "center"
            },
          };
        }
      }
    }
  }
}
function style5() {
  var cellule = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L"];
  for (c = 0; c < cellule.length; c++) {
    for (i = 1; i <= individual_live.length; i++) {
      if (ws_individual[cellule[c] + "" + i]) {
        if (i == 1 || i == 2) {
          ws_individual[cellule[c] + "" + i].s = {
            font: {
              name: "Arial",
              sz: 10,
              bold: true,
            },
            alignment: {
              vertical: "center",
              horizontal: "center"
            },
          };
        }
        else if (cellule[c] == "H") {
          if (ws_individual[cellule[c] + "" + i].v == "" || ws_individual[cellule[c] + "" + i].v == "Congé sans solde") {
            ws_individual[cellule[c] + "" + i].s = {
              font: {
                name: "Arial",
                sz: 10,
                color: { rgb: "F7931A" }
              },
              border: {
                left: { style: "hair" },
                right: { style: "hair" },
                top: {
                  style: "hair",
                  bottom: { style: "hair" },
                },
              },
              alignment: {
                vertical: "center",
                horizontal: "center"
              },
            };
          }
          else if (i == individual_live.length) {
            ws_individual[cellule[c] + "" + i].s = {
              fill: {
                patternType: "solid",
                fgColor: { rgb: "BFBFBF" },
                bgColor: { rgb: "BFBFBF" },
              },
              font: {
                name: "Arial",
                sz: 10,
                color: { rgb: "F7931A" }
              },
              border: {
                left: { style: "hair" },
                right: { style: "hair" },
                top: {
                  style: "hair",
                  bottom: { style: "hair" },
                },
              },
              alignment: {
                vertical: "center",
                horizontal: "center"
              },
            };
          }
          else {
            ws_individual[cellule[c] + "" + i].s = {
              font: {
                name: "Arial",
                sz: 10,
                color: { rgb: "F7931A" },
                bold: true
              },
              border: {
                left: { style: "hair" },
                right: { style: "hair" },
                top: {
                  style: "hair",
                  bottom: { style: "hair" },
                },
              },
              alignment: {
                vertical: "center",
                horizontal: "center"
              },
            };
          }

        }
        else if (cellule[c] == "I") {
          if (ws_individual[cellule[c] + "" + i].v == "" || ws_individual[cellule[c] + "" + i].v == "OSTIE") {
            ws_individual[cellule[c] + "" + i].s = {
              font: {
                name: "Arial",
                sz: 10,
                color: { rgb: "FF2AE0" }
              },
              border: {
                left: { style: "hair" },
                right: { style: "hair" },
                top: {
                  style: "hair",
                  bottom: { style: "hair" },
                },
              },
              alignment: {
                vertical: "center",
                horizontal: "center"
              },
            };
          }
          else if (i == individual_live.length) {
            ws_individual[cellule[c] + "" + i].s = {
              fill: {
                patternType: "solid",
                fgColor: { rgb: "BFBFBF" },
                bgColor: { rgb: "BFBFBF" },
              },
              font: {
                name: "Arial",
                sz: 10,
                color: { rgb: "FF2AE0" }
              },
              border: {
                left: { style: "hair" },
                right: { style: "hair" },
                top: {
                  style: "hair",
                  bottom: { style: "hair" },
                },
              },
              alignment: {
                vertical: "center",
                horizontal: "center"
              },
            };
          }
          else {
            ws_individual[cellule[c] + "" + i].s = {
              font: {
                name: "Arial",
                sz: 10,
                color: { rgb: "FF2AE0" },
                bold: true
              },
              border: {
                left: { style: "hair" },
                right: { style: "hair" },
                top: {
                  style: "hair",
                  bottom: { style: "hair" },
                },
              },
              alignment: {
                vertical: "center",
                horizontal: "center"
              },
            };
          }

        }
        else if (cellule[c] == "J") {
          if (ws_individual[cellule[c] + "" + i].v == "" || ws_individual[cellule[c] + "" + i].v == "Pérmission Exceptionelle") {
            ws_individual[cellule[c] + "" + i].s = {
              font: {
                name: "Arial",
                sz: 10,
                color: { rgb: "5A9966" }
              },
              border: {
                left: { style: "hair" },
                right: { style: "hair" },
                top: {
                  style: "hair",
                  bottom: { style: "hair" },
                },
              },
              alignment: {
                vertical: "center",
                horizontal: "center"
              },
            };
          }
          else if (i == individual_live.length) {
            ws_individual[cellule[c] + "" + i].s = {
              fill: {
                patternType: "solid",
                fgColor: { rgb: "BFBFBF" },
                bgColor: { rgb: "BFBFBF" },
              },
              font: {
                name: "Arial",
                sz: 10,
                color: { rgb: "5A9966" }
              },
              border: {
                left: { style: "hair" },
                right: { style: "hair" },
                top: {
                  style: "hair",
                  bottom: { style: "hair" },
                },
              },
              alignment: {
                vertical: "center",
                horizontal: "center"
              },
            };
          }
          else {
            ws_individual[cellule[c] + "" + i].s = {
              font: {
                name: "Arial",
                sz: 10,
                color: { rgb: "5A9966" },
                bold: true
              },
              border: {
                left: { style: "hair" },
                right: { style: "hair" },
                top: {
                  style: "hair",
                  bottom: { style: "hair" },
                },
              },
              alignment: {
                vertical: "center",
                horizontal: "center"
              },
            };
          }
        }
        else if (cellule[c] == "K") {
          if (ws_individual[cellule[c] + "" + i].v == "" || ws_individual[cellule[c] + "" + i].v == "Absent et autres") {
            ws_individual[cellule[c] + "" + i].s = {
              font: {
                name: "Arial",
                sz: 10,
                color: { rgb: "0066BC" }
              },
              border: {
                left: { style: "hair" },
                right: { style: "hair" },
                top: {
                  style: "hair",
                  bottom: { style: "hair" },
                },
              },
              alignment: {
                vertical: "center",
                horizontal: "center"
              },
            };
          }
          else if (i == individual_live.length) {
            ws_individual[cellule[c] + "" + i].s = {
              fill: {
                patternType: "solid",
                fgColor: { rgb: "BFBFBF" },
                bgColor: { rgb: "BFBFBF" },
              },
              font: {
                name: "Arial",
                sz: 10,
                color: { rgb: "0066BC" }
              },
              border: {
                left: { style: "hair" },
                right: { style: "hair" },
                top: {
                  style: "hair",
                  bottom: { style: "hair" },
                },
              },
              alignment: {
                vertical: "center",
                horizontal: "center"
              },
            };
          }
          else {
            ws_individual[cellule[c] + "" + i].s = {
              font: {
                name: "Arial",
                sz: 10,
                color: { rgb: "0066BC" },
                bold: true
              },
              border: {
                left: { style: "hair" },
                right: { style: "hair" },
                top: {
                  style: "hair",
                  bottom: { style: "hair" },
                },
              },
              alignment: {
                vertical: "center",
                horizontal: "center"
              },
            };
          }
        }
        else if (ws_individual[cellule[c] + "" + i].v == "Fin Année") {
          ws_individual[cellule[c] + "" + i].s = {
            fill: {
              patternType: "solid",
              fgColor: { rgb: "BFBFBF" },
              bgColor: { rgb: "BFBFBF" },
            },
            font: {
              name: "Arial",
              sz: 10,
              bold: true
            },
            border: {
              left: { style: "hair" },
              right: { style: "hair" },
              top: {
                style: "hair",
                bottom: { style: "hair" },
              },
            },
            alignment: {
              vertical: "center",
              horizontal: "center"
            },
          };
        }
        else if (ws_individual[cellule[c] + "" + i].v == "Congé Maternité") {
          ws_individual[cellule[c] + "" + i].s = {
            fill: {
              patternType: "solid",
              fgColor: { rgb: "25AFF3" },
              bgColor: { rgb: "25AFF3" },
            },
            font: {
              name: "Arial",
              sz: 10,
              color: { rgb: "FFFFFF" },
              bold: true
            },
            border: {
              left: { style: "hair" },
              right: { style: "hair" },
              top: {
                style: "hair",
                bottom: { style: "hair" },
              },
            },
            alignment: {
              vertical: "center",
              horizontal: "center"
            },
          };
        }
        else if (ws_individual[cellule[c] + "" + i].v == "Mise a Pied ( a déduire sur salaire )") {
          ws_individual[cellule[c] + "" + i].s = {
            fill: {
              patternType: "solid",
              fgColor: { rgb: "E00012" },
              bgColor: { rgb: "E00012" },
            },
            font: {
              name: "Arial",
              sz: 10,
              bold: true
            },
            border: {
              left: { style: "hair" },
              right: { style: "hair" },
              top: {
                style: "hair",
                bottom: { style: "hair" },
              },
            },
            alignment: {
              vertical: "center",
              horizontal: "center"
            },
          };
        }
        else if (ws_individual[cellule[c] + "" + i].v == "Repos Maladie ( rien à deduire )") {
          ws_individual[cellule[c] + "" + i].s = {
            fill: {
              patternType: "solid",
              fgColor: { rgb: "FF2AE0" },
              bgColor: { rgb: "FF2AE0" },
            },
            font: {
              name: "Arial",
              sz: 10,
              bold: true
            },
            border: {
              left: { style: "hair" },
              right: { style: "hair" },
              top: {
                style: "hair",
                bottom: { style: "hair" },
              },
            },
            alignment: {
              vertical: "center",
              horizontal: "center"
            },
          };
        }
        else if (ws_individual[cellule[c] + "" + i].v == "Absent ( a déduire sur salaire )") {
          ws_individual[cellule[c] + "" + i].s = {
            fill: {
              patternType: "solid",
              fgColor: { rgb: "0066BC" },
              bgColor: { rgb: "0066BC" },
            },
            font: {
              name: "Arial",
              sz: 10,
              bold: true
            },
            border: {
              left: { style: "hair" },
              right: { style: "hair" },
              top: {
                style: "hair",
                bottom: { style: "hair" },
              },
            },
            alignment: {
              vertical: "center",
              horizontal: "center"
            },
          };
        }
        else if (ws_individual[cellule[c] + "" + i].v == "Permission exceptionelle ( rien à deduire )") {
          ws_individual[cellule[c] + "" + i].s = {
            fill: {
              patternType: "solid",
              fgColor: { rgb: "5A9966" },
              bgColor: { rgb: "5A9966" },
            },
            font: {
              name: "Arial",
              sz: 10,
              bold: true
            },
            border: {
              left: { style: "hair" },
              right: { style: "hair" },
              top: {
                style: "hair",
                bottom: { style: "hair" },
              },
            },
            alignment: {
              vertical: "center",
              horizontal: "center"
            },
          };
        }
        else if (ws_individual[cellule[c] + "" + i].v == "Congé sans solde ( a déduire sur salaire )") {
          ws_individual[cellule[c] + "" + i].s = {
            fill: {
              patternType: "solid",
              fgColor: { rgb: "F7931A" },
              bgColor: { rgb: "F7931A" },
            },
            font: {
              name: "Arial",
              sz: 10,
              bold: true
            },
            border: {
              left: { style: "hair" },
              right: { style: "hair" },
              top: {
                style: "hair",
                bottom: { style: "hair" },
              },
            },
            alignment: {
              vertical: "center",
              horizontal: "center"
            },
          };
        }
        else if (i == individual_live.length) {
          if (cellule[c] == "G") {
            ws_individual[cellule[c] + "" + i].s = {
              fill: {
                patternType: "solid",
                fgColor: { rgb: "BFBFBF" },
                bgColor: { rgb: "BFBFBF" },
              },
              font: {
                name: "Arial",
                sz: 10,
                color: { rgb: "000000" },
                bold: true
              },
              border: {
                left: { style: "hair" },
                right: { style: "hair" },
                top: {
                  style: "hair",
                  bottom: { style: "hair" },
                },
              },
              alignment: {
                vertical: "center",
                horizontal: "center"
              },
            };
          }
          else {
            ws_individual[cellule[c] + "" + i].s = {
              fill: {
                patternType: "solid",
                fgColor: { rgb: "BFBFBF" },
                bgColor: { rgb: "BFBFBF" },
              },
              font: {
                name: "Arial",
                sz: 10,
                color: { rgb: "000000" },
                bold: true
              },
              border: {
                left: { style: "hair" },
                right: { style: "hair" },
                top: {
                  style: "hair",
                  bottom: { style: "hair" },
                },
              },
              alignment: {
                vertical: "center",
                horizontal: "center"
              },
            };
          }
        }
        else {
          ws_individual[cellule[c] + "" + i].s = {
            font: {
              name: "Arial",
              sz: 10,
              color: { rgb: "000000" }
            },
            border: {
              left: { style: "hair" },
              right: { style: "hair" },
              top: {
                style: "hair",
                bottom: { style: "hair" },
              },
            },
            alignment: {
              vertical: "center",
              horizontal: "center"
            },
          };
        }
      }
    }

  }
}
//Fonction generate excel
function generate_excel(datatowrites, retard, absent, conge, code) {
  var counter = 0;
  var cum = 0;
  var cumg = 0;
  var cum_tot = "";
  var cum_abs = "";
  var cum_del = "";
  var nom = "";
  var m_codes = "";
  for (i = 0; i < datatowrites.length; i++) {
    if (datatowrites[i].time_end != "" && datatowrites[i].m_code == code) {
      counter++;
      var ligne = [
        datatowrites[i].m_code,
        datatowrites[i].num_agent,
        datatowrites[i].date,
        datatowrites[i].locaux,
        datatowrites[i].time_start,
        datatowrites[i].time_end,
        calcul_timediff(datatowrites[i].time_start, datatowrites[i].time_end),
        datatowrites[i].worktime,
      ];
      nom = datatowrites[i].nom;
      m_codes = datatowrites[i].m_code;
      data.push(ligne);

    }
  }
  totaltime = hours + "H " + minutes + "MN";
  data.push(["", "", "", "", "", "TOTALE", totaltime, ""]);
  cum_tot = totaltime;
  data.push(["", "", "", "", "", "", "", ""]);
  hours = 0; minutes = 0;
  if (retard.length != 0) {
    data.push(["", "", "", "", "", "", "", ""]);
    data.push(["", "", "", "Rapport retard", "", "", ""]);
    data.push(["M-code", "Numéro Agent", "Date", "Raison", "Temp", "", ""]);
    for (i = 0; i < retard.length; i++) {
      if (retard[i].m_code == code) {
        cum += retard[i].time;
        var lateligne = [
          retard[i].m_code,
          retard[i].num_agent,
          retard[i].date,
          retard[i].reason,
          retard[i].time + " minutes",
        ];
        data.push(lateligne);
      }
    }
    if (cum != 0) {
      cum = convert_to_hour(cum).split(",");
      cum_del = cum[0] + "H" + " " + cum[1] + " MN";
      data.push(["", "", "", "TOTAL", cum[0] + "H" + " " + cum[1] + " MN", ""]);
    }
    else {
      data.push(["", "", "", "TOTAL", "0H" + " " + "0 MN", ""]);
    }
  }
  if (absent.length != 0) {
    for (i = 0; i < absent.length; i++) {
      var latelignent = [];
      if (absent[i].return != "Not come back" && absent[i].m_code == code) {
        var lateligne = [
          absent[i].m_code,
          absent[i].num_agent,
          absent[i].date,
          absent[i].reason,
          absent[i].time_start,
          absent[i].return,
          absent[i].status,
        ];
        data.push(["", "", "", "", "", "", "", ""]);
        data.push(["", "", "", "ABSENCE AVEC RETOUR", "", "", "", ""]);
        data.push(["M-code", "Numéro Agent", "Date", "Raison", "Début", "Retourner", "Status", ""]);
        data.push(lateligne);
        calcul_timediff(absent[i].time_start, absent[i].return);
      }
      else {
        if (absent[i].m_code == code) {
          latelignent.push(
            absent[i].m_code,
            absent[i].num_agent,
            absent[i].date,
            absent[i].reason,
            absent[i].time_start,
            "n'a pas retourner",
            absent[i].status,
          );
        }
      }

    }
    totaltime = hours + "H " + minutes + "MN";
    cum_abs = totaltime;
    if (latelignent.length > 0) {
      data.push(["", "", "", "", "", "TOTAL", totaltime, ""]);
      data.push(["", "", "", "ABSENCE SANS RETOUR", "", "", "", ""]);
      data.push(["M-code", "Numéro Agent", "Date", "Raison", "Début", "Retour", "Status", ""]);
      data.push(latelignent);
    }
  }
  if (conge.length != 0) {
    for (i = 0; i < conge.length; i++) {
      if (conge[i].m_code == code) {
        var lateligne = [
          conge[i].m_code,
          conge[i].num_agent,
          conge[i].date_start,
          conge[i].date_end,
          conge[i].duration + " jour(s)",
          conge[i].type,
        ];
        data.push(["", "", "", "", "", "", "", ""]);
        data.push(["", "", "", "RAPPORT CONGE", "", "", "", ""]);
        data.push(["M-code", "Numéro agent", "Date Début", "Date Fin", "Nombre jour", "Type", ""]);
        data.push(lateligne);
        cumg += conge[i].duration;
      }
    }
    data.push(["", "", "", "", "", "TOTAL", cumg + " jour(s)"]);
  }
  each_data = [
    nom,
    m_codes,
    cum_tot,
    cum_del,
    cum_abs,
    cumg + " jour(s)"
  ]
  all_datas.push(each_data);
  ws = ExcelFile.utils.aoa_to_sheet(data);
  ws["!cols"] = [
    { wpx: 80 },
    { wpx: 130 },
    { wpx: 200 },
    { wpx: 250 },
    { wpx: 160 },
    { wpx: 100 },
    { wpx: 120 },
    { wpx: 120 }
  ];
  const merge = [
    { s: { r: 0, c: 0 }, e: { r: 1, c: 7 } },
    { s: { r: 3, c: 0 }, e: { r: counter + 2, c: 0 } },
    { s: { r: 3, c: 1 }, e: { r: counter + 2, c: 1 } },
    { s: { r: counter + 5, c: 0 }, e: { r: counter + 5, c: 7 } }
  ];
  ws["!merges"] = merge;
  style();
}
function global_Report(all_data) {
  ws = ExcelFile.utils.aoa_to_sheet(all_data);
  ws["!cols"] = [
    { wpx: 230 },
    { wpx: 80 },
    { wpx: 150 },
    { wpx: 150 },
    { wpx: 150 },
    { wpx: 150 },
    { wpx: 150 }
  ];
  const merge = [
    { s: { r: 0, c: 0 }, e: { r: 1, c: 5 } }
  ];
  ws["!merges"] = merge;
  style2();
}

module.exports = routeExp;