const {Dropbox} = require("dropbox");
var fs = require('fs');
var excel = require('excel4node');
const readline = require('readline-sync');
const config = require('./config');


const SCREENSHOTS_DIR = readline.question('Enter path:').toString().replace(/\\/g, '/');
console.log(SCREENSHOTS_DIR);

const DROPBOX_DIR = '/'+(SCREENSHOTS_DIR.split('/').pop());

var workbook = new excel.Workbook();
var worksheet = workbook.addWorksheet('Sheet 1');


var dbx = new Dropbox({accessToken: config.dropBoxKey});



let currentRow = 1;

let rows = {};



let level = -1;
let colors = [
    'd5a6bd',
    'ead1dc',
    'ffeef5',
]



run();


async function run() {
    await recursive('');
    // console.log(rows);
    Object.values(rows).forEach(row => row.exel());
    workbook.write('output.xlsx');
}

async function recursive(dirPath) {
    level += 1;

    var colorIndex = level < colors.length ? level : colors.length-1;
    const headerRowColor = colors[colorIndex ];
    // console.log(headerRowColor);

    const realFilePath = SCREENSHOTS_DIR+dirPath;
    let entries = fs.readdirSync(realFilePath);

    let hasFiles = false;
    let hasDirs = false;

    entries.forEach(e => {
        var q = e.indexOf('.');
        if(q > 0) {
            hasFiles = true;
        } else if(q === -1) {
            hasDirs = true;
        }
    });

    entries.sort((a,b) => fs.statSync(realFilePath+'/'+a).birthtime > fs.statSync(realFilePath+'/'+b).birthtime ? 1 : -1);

    if(hasFiles && hasDirs) {
        entries.sort((a, b) =>  b.indexOf('.'));
        let key = 'common'+dirPath;
        rows[key] = {
            row: currentRow++,
            exel: () => worksheet.cell(rows[key].row, 1, rows[key].row, 30, true).string('Общие').style(workbook.createStyle({
                fill: {
                    type: 'pattern', // the only one implemented so far.
                    patternType: 'solid', // most common.
                    fgColor: headerRowColor, // you can add two extra characters to serve as alpha, i.e. '2172d7aa'.
                    // bgColor: 'ffffff' // bgColor only applies on patternTypes other than solid.
                }
            })),
        }
    }

    for(let name of entries) {
        let splitName = name.split('.');
        if(splitName && splitName.length === 1) {
            let dirName = splitName[0];
            let key = dirPath+'/'+dirName;
            rows[key] = {
                row: currentRow++,
                exel: () => worksheet.cell(rows[key].row, 1, rows[key].row, 30, true).string(dirName).style(workbook.createStyle({
                    fill: {
                        type: 'pattern', // the only one implemented so far.
                        patternType: 'solid', // most common.
                        fgColor: headerRowColor, // you can add two extra characters to serve as alpha, i.e. '2172d7aa'.
                        // bgColor: 'ffffff' // bgColor only applies on patternTypes other than solid.
                    }
                })),
            }
            await recursive(key);
        } else {
            let bugName = splitName[0];

            let numberMatch = bugName.match(/(.*\D+)\d+$/);
            if(numberMatch && numberMatch.length > 1) {
                bugName = numberMatch[1];
            }
            let key = dirPath+'/'+bugName;
            var r = currentRow++;
            if(!rows[key]) {
                rows[key] = {
                    row: r,
                    files: []
                };
            }

            var path = DROPBOX_DIR+dirPath+ '/' + name;
            // console.log(path);
            var screenshot = fs.readFileSync(SCREENSHOTS_DIR+'/'+ dirPath+ '/' + name);
            var link = await dbx.sharingListSharedLinks({
                path: path
            }).then(res => res.result.links[0]).catch(e => {
                if(e.error.error['.tag'] !== 'path') throw e;
            }).catch(console.log);


            if(!link) {
                var dropboxFile = await dbx.filesUpload({
                    path: path,
                    contents: screenshot
                }).then((res) => res.result).catch(console.log);

                link = await dbx.sharingCreateSharedLinkWithSettings({
                    path: path
                }).then((res) => res.result).catch(console.log);
            }


            rows[key].files.push(link.url);
            rows[key].exel = () => {
                let links = rows[key].files;

                worksheet.cell(rows[key].row,1, rows[key].row+links.length-1,1, true).string(bugName).style(workbook.createStyle({

                }));
                for(let i = 0; i < links.length; i++) {
                    worksheet.cell(rows[key].row+i,2).formula('HYPERLINK("'+links[i]+'")').style(workbook.createStyle({

                    }));
                }

            };
        }
    }
    level -= 1;
}
