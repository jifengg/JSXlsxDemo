const fs = require('fs');
const xlsxDemo = require('./demo');

async function test2file() {

    var book = await xlsxDemo();
    var buffer = await book.SaveAsBuffer();
    var path = require('path').join(process.cwd(), 'out.xlsx');
    fs.writeFileSync(path, buffer);
    console.log('生成文件到本地', path);
    require('child_process').exec(path);
}

test2file();