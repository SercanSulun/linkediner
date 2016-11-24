alert('Sercan!');
var electron = require('electron');
var dialog = electron.remote.dialog;
var fs = require('fs');
var parser = new DOMParser();
var persons = [];
var fileReadEnd = 0;
var numberOfFiles = 0;

document.getElementById('select-folder').addEventListener('click',function(){
    dialog.showOpenDialog({properties: ['openDirectory']},function (fileNames) {
        if(fileNames === undefined){
            console.log("No file selected");
        }else{
          document.getElementById('actual-folder').value = fileNames[0];
        }
    });
},false);

document.getElementById('convert-button').addEventListener('click',function(){
    folderToRead = document.getElementById("actual-folder").value;
    fs.readdir(folderToRead, function(err,files){
        var filesToRead = files.filter(isHTMLFile);
        numberOfFiles = filesToRead.length;
        filesToRead.forEach(readFile,folderToRead);
    })
},false);

function readFile(filepath,index,arr) {
    fs.readFile(this+"\\"+filepath, 'utf-8', function (err, data) {
        if(err){
            alert("An error ocurred reading the file :" + err.message);
            return;
        }
        var codeStartIndex = data.indexOf('<code');
        var codeEndIndex = data.indexOf('</code>');
        var codeString = data.substring(codeStartIndex,codeEndIndex);
        codeStartIndex = codeString.indexOf('<!--');
        codeEndIndex = codeString.indexOf('-->');
        codeString = codeString.substring(codeStartIndex+4,codeEndIndex);
        var parsedJSON = JSON.parse(codeString);
        Array.prototype.push.apply(persons,parsedJSON.content.page.voltron_unified_search_json.search.results.filter(isPerson).map(function(e){
          return e.person;
        }));
        console.log(persons);
        fileReadEnd = fileReadEnd + 1;
        if(fileReadEnd == numberOfFiles) writeToCsv();
    });
}

function isPerson(value){
    return value['person']!==undefined
}

function isHTMLFile(value){
    return value.substring(value.lastIndexOf(".")) == ".html"
}

function writeToCsv(){
    var CsvString = "First Name;Last Name;Location;Industry;Headline;Past Headline\n"
    for(var i = 0; i < persons.length; i++){
        CsvString = CsvString
            +persons[i].firstName.replace(/;/g," - ").replace(/ı/g,"i").replace(/&amp/g,"&").replace(/ş/g,"s").replace(/Ş/g,"S").replace(/&#x2F/g,"/").replace(/ğ/g,"g").replace(/Ğ/g,"G")+";"
            +persons[i].lastName.replace(/;/g," - ").replace(/ı/g,"i").replace(/&amp/g,"&").replace(/ş/g,"s").replace(/Ş/g,"S").replace(/&#x2F/g,"/").replace(/ğ/g,"g").replace(/Ğ/g,"G")+";"
            +persons[i].fmt_location.replace(/;/g," - ").replace(/ı/g,"i").replace(/&amp/g,"&").replace(/ş/g,"s").replace(/Ş/g,"S").replace(/&#x2F/g,"/").replace(/ğ/g,"g").replace(/Ğ/g,"G")+";"
            +persons[i].fmt_industry.replace(/;/g," - ").replace(/ı/g,"i").replace(/&amp/g,"&").replace(/ş/g,"s").replace(/Ş/g,"S").replace(/&#x2F/g,"/").replace(/ğ/g,"g").replace(/Ğ/g,"G")+";"
            +persons[i].fmt_headline.replace(/;/g," - ").replace(/ı/g,"i").replace(/&amp/g,"&").replace(/ş/g,"s").replace(/Ş/g,"S").replace(/&#x2F/g,"/").replace(/ğ/g,"g").replace(/Ğ/g,"G")+";"
            +persons[i].snippets[0].fmt_heading.replace(/;/g," - ").replace(/ı/gi,"i").replace(/&amp/g,"&").replace(/ş/g,"s").replace(/Ş/g,"S").replace(/&#x2F/g,"/").replace(/ğ/g,"g").replace(/Ğ/g,"G")+"\n";
    }
    CsvString = CsvString.replace(/<B>/gi,'-').replace(/<\/B>/gi,'-');
    fs.writeFile(folderToRead+"\\Contacts_"+new Date().getTime()+".csv",CsvString,'ucs2',function(err){
        if (err) throw err;
        console.log('It\'s saved!');
        alert('It\'s saved!');
    })
}
