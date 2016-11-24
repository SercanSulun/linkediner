const electron = require('electron');
const app = electron.app;
const BrowserWindow = electron.BrowserWindow;

app.on('ready', function() {
    var mainWindow = new BrowserWindow({width: 400,height: 400});
    mainWindow.loadURL("file://"+__dirname+"/Interface.html");
})
app.on('close', function() {
    mainWindow = null;
})
