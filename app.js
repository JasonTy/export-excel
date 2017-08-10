/**
 * Created by tangyitangyi on 2017/7/31.
 */
var express = require('express');//express  框架
var path = require('path');//跳转相关
var bodyParser = require('body-parser');
var cookieParser = require('cookie-parser');
var engine = require('express-dot-engine');
var session = require('express-session');
var uuid = require('uuid');
var dot = require('dot');
var requireDir = require('require-dir');
var fs = require('fs');


var app = express();

var routes = requireDir('./routes', {recurse: true});


app.engine('html', engine.__express);
app.set('views', path.join(__dirname, 'views'));
app.set('view engine', 'html');


app.use(bodyParser.json());
app.use(bodyParser.urlencoded({extended: false}));
app.use(cookieParser());
app.use(express.static(path.join(__dirname, 'public')));

app.use('/', routes.index);
app.use('/export', routes.index);

module.exports = app;