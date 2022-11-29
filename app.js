const express = require("express");
const app = express();
const cors = require("cors");
const bodyParser = require("body-parser");
const route = require("./public/route.js");
const PORT = process.env.PORT || 8080;
const methodOverride = require("method-override");
const expsession = require('cookie-session');
const moment = require("moment");
const fileUpload = require('express-fileupload')

app.use(methodOverride("X-HTTP-Method"));
app.use(methodOverride("X-HTTP-Method-Override"));
app.use(methodOverride("X-Method-Override"));
app.use(methodOverride("_method"));
app.use(expsession({
  name: 'session',
  keys: ["2C44-4D44-WppQ38S"],
  resave: true,
  saveUninitialized: true,
  overwrite:true,
  // Cookie Options
  maxAge: 15 * 60 * 60 * 1000 // 15 hours
}))

// Fichier static a utiliser
app.use(express.static("public"));
app.use(express.static("node_modules"));
app.use(express.static('public/assets'));
app.use(express.static('public/Paie'));
app.use(express.static('public/dist'));
app.use(express.static('public/src'));

// View de type html
app.engine("html", require("ejs").renderFile);
app.set("view engine", "html");
app.set("views", __dirname + "/public");

//app.use(express.static(__dirname + "/public"));

app.use(cors());
app.use(fileUpload());
app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());

const server = require("http").createServer(app);
const io = require("socket.io")(server);
io.on('connection', socket => {
  socket.on('actuel', msg =>{
      socket.broadcast.emit("status",msg);
  })
  socket.on('loc', lc =>{
    socket.broadcast.emit("locaux",lc);
})
});
app.set("io", io);
app.use('/', route);

server.listen(PORT, () => {
  const port = server.address().port;
  console.log(`Express is working on port ${port}`);
});

