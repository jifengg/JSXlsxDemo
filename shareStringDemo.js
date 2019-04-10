
let XlsxCore = require('js-xlsx-core');
let fs = require('fs');
require('xlsx-saver');
const {
    Book,
    Sheet,
    Cell,
    ShareString,
    CellStyle,
    CellAlignment,
    NumberFormat,
    Image,
    ImageOption,
    HorizontalAlignment,
    VerticalAlignment
} = XlsxCore;

async function ShareStringDemo() {
    var str = '这是一个文本用于测试使用了ShareString和未使用时对文件体积的影响。文字越长，用的次数越多，越能减少最终文件的体积。';
    var count = 1e4;
    var book = new Book();
    var sheet = book.CreateSheet('sheet1');
    for (var i = 0; i < count; i++) {
        sheet.AddText(str, i, 0);
    }
    var buffer = await book.SaveAsBuffer();
    var path = process.cwd() + '/NotUseShareString.xlsx';
    fs.writeFileSync(path, buffer);
    var size1 = fs.statSync(path).size;

    book = new Book();
    sheet = book.CreateSheet('sheet1');
    var shareString = book.CreateShareString(str);
    for (var i = 0; i < count; i++) {
        sheet.AddText(shareString, i, 0);
    }
    buffer = await book.SaveAsBuffer();
    path = process.cwd() + '/UseShareString.xlsx';
    fs.writeFileSync(path, buffer);
    var size2 = fs.statSync(path).size;

    console.log(`输出完毕，未使用ShareString文件大小：${size1} Byte；使用ShareString文件大小：${size2} Byte，节约了${((size1 - size2) / size2 * 100).toPrecision()}%。`);
}

ShareStringDemo();