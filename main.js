var electron = require('electron');
var dialog = electron.remote.dialog;
var fs = require('fs');
var xl = require('excel4node');
var parser = new DOMParser();
var persons = [];
var fileReadEnd = 0;
var numberOfFiles = 0;
var headerArray = [{
    columnName: 'First Name',
    JSONFieldPath: 'firstName',
    maxCharLength: 0
}, {
    columnName: 'Last Name',
    JSONFieldPath: 'lastName',
    maxCharLength: 0
}, {
    columnName: 'Location',
    JSONFieldPath: 'fmt_location',
    maxCharLength: 0
}, {
    columnName: 'Industry',
    JSONFieldPath: 'fmt_industry',
    maxCharLength: 0
}, {
    columnName: 'Headline',
    JSONFieldPath: 'fmt_headline',
    maxCharLength: 0
}, {
    columnName: 'Past Headline',
    JSONFieldPath: 'snippets[0].fmt_heading',
    maxCharLength: 0
}];

document.getElementById('select-folder').addEventListener('click', function() {
    dialog.showOpenDialog({
        properties: ['openDirectory']
    }, function(fileNames) {
        if (fileNames === undefined) {
            console.log("No file selected");
        } else {
            document.getElementById('actual-folder').value = fileNames[0];
        }
    });
}, false);

document.getElementById('convert-button').addEventListener('click', function() {
    folderToRead = document.getElementById("actual-folder").value;
    fs.readdir(folderToRead, function(err, files) {
        var filesToRead = files.filter(isHTMLorHTMLFile);
        numberOfFiles = filesToRead.length;
        fileReadEnd = 0;
        filesToRead.forEach(readFile, folderToRead);
    })
}, false);

function readFile(filepath, index, arr) {
    var folderToRead = this;
    fs.readFile(this + "\\" + filepath, 'utf-8', function(err, data) {
        if (err) {
            alert("An error ocurred reading the file :" + err.message);
            return;
        }
        var codeStartIndex = data.indexOf('<code');
        var codeEndIndex = data.indexOf('</code>');
        var codeString = data.substring(codeStartIndex, codeEndIndex);
        codeStartIndex = codeString.indexOf('<!--');
        codeEndIndex = codeString.indexOf('-->');
        codeString = codeString.substring(codeStartIndex + 4, codeEndIndex);
        var parsedJSON = JSON.parse(codeString);
        Array.prototype.push.apply(persons, parsedJSON.content.page.voltron_unified_search_json.search.results.filter(isPerson).map(function(e) {
            return e.person;
        }));
        console.log(persons);
        fileReadEnd = fileReadEnd + 1;
        if (fileReadEnd == numberOfFiles) exportToExcel(folderToRead);
    });
}

function isPerson(value) {
    return value['person'] !== undefined
}

function isHTMLorHTMLFile(value) {
    var ext = value.substring(value.lastIndexOf("."));
    return ext == ".html"
}

function pathFinder(obj, i) {
    return obj[i]
}

function pathMapArrayIntoPath(obj) {
    return obj.replace(/\]/gi, '').replace(/\[/gi, '.')
}

function exportToExcel(pathToWrite) {

    var wb = new xl.Workbook();
    var ws = wb.addWorksheet('Persons');
    var currentHeader;

    var headerStyle = wb.createStyle({
        font: {
            bold: true,
            color: 'FFFFFF'
        },
        border: { // §18.8.4 border (Border)
            left: {
                style: 'medium', //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
                color: '236B8E' // HTML style hex value
            },
            right: {
                style: 'medium',
                color: '236B8E'
            },
            top: {
                style: 'medium',
                color: '236B8E'
            },
            bottom: {
                style: 'medium',
                color: '236B8E'
            }
        },
        fill: { // §18.8.20 fill (Fill)
            type: 'pattern', // Currently only 'pattern' is implimented. Non-implimented option is 'gradient'
            patternType: 'solid', //§18.18.55 ST_PatternType (Pattern Type)
            bgColor: '236B8E', // HTML style hex value. optional. defaults to black
            fgColor: '236B8E' // HTML style hex value. required.
        }
    });

    var rowstyle = wb.createStyle({
        border: { // §18.8.4 border (Border)
            top: {
                style: 'thin',
                color: '236B8E'
            },
            bottom: {
                style: 'thin',
                color: '236B8E'
            }
        }
    });

    var tableLeftSide = wb.createStyle({
        border: {
            left: {
                style: 'thin',
                color: '236B8E'
            }
        }
    })

    var tableRightSide = wb.createStyle({
        border: {
            right: {
                style: 'thin',
                color: '236B8E'
            }
        }
    })

    for (let i = 0; i < headerArray.length; i++) {
        currentHeader = headerArray[i];
        currentHeader.maxCharLength = currentHeader.columnName.length;
        ws.cell(1, i + 1).string(currentHeader.columnName).style(headerStyle);
        for (let j = 0; j < persons.length; j++) {
            var stringToAssign = pathMapArrayIntoPath(currentHeader.JSONFieldPath).split('.').reduce(pathFinder, persons[j]);
            stringToAssign = stringToAssign.replace(/<B>/gi, '').replace(/<\/B>/gi, '').replace(/&amp;/gi,'&');
            currentHeader.maxCharLength = Math.max(currentHeader.maxCharLength, stringToAssign.length);
            ws.cell(j + 2, i + 1).string(stringToAssign).style(rowstyle);
            if (i == 0) {
                ws.cell(j + 2, i + 1).style(tableLeftSide);
            } else if (i == headerArray.length - 1) {
                ws.cell(j + 2, i + 1).style(tableRightSide);
            }
        }
        ws.column(i + 1).setWidth(currentHeader.maxCharLength);
    }

    ws.row(1).filter({
        firstRow: 1,
        firstColumn: 1,
        lastRow: persons.length,
        lastColumn: headerArray.length
    });

    writeExcelFile(wb, pathToWrite);
}

function writeExcelFile(wb, pathToWrite) {
    wb.write(pathToWrite + '\\Contacts_'+new Date().getTime()+ '.xlsx', function(err, stats) {
        if (err) {
            if (err.code == "EBUSY") {
                alert('Cannot save file bacause it is in use. Please close the current file and try again.');
            } else {
                alert(err);
            }
        } else {
            alert('File is saved.')
        }
        console.log(stats);
        console.log(headerArray);
    });
}
