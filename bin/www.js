/**
 * Created by tangyitangyi on 2017/7/31.
 */
var http = require('http');
var app = require('../app.js')
http.createServer(app).listen(3016).on('listening', function () {
    console.log('Listening on port 3016');
});