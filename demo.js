(() => {
    if (typeof global == 'undefined') {
        window.require = function () { };
    }
    const request = require('request');
    let XlsxCore = require('js-xlsx-core');
    require('xlsx-saver')
    if (typeof global == 'undefined') {
        XlsxCore = window.XlsxCore;
    }
    let Book = XlsxCore.Book;
    let HorizontalAlignment = XlsxCore.HorizontalAlignment;
    let VerticalAlignment = XlsxCore.VerticalAlignment;

    async function xlsxDemo() {
        var book = new Book();
        //设置默认样式中的字号为11号
        book.DefaultCellStyle.FontSize = 11;

        var sheet = book.CreateSheet("第一页");

        //设置默认列宽和默认行高
        sheet.DefaultHeight = 25;
        sheet.DefaultWidth = 33;

        //设置A列宽
        sheet.SetColWidth(0, 35);
        //设置第二行高度
        sheet.SetRowHeight(1, 35);

        var row = 0;
        var col = 0;
        //添加默认文字
        sheet.AddText('一个普通文本', row++, col);
        sheet.AddText('第二行第一列', row++, col);

        sheet.AddText('粗体文字', row++, col, {
            Bold: true
        });
        sheet.AddText('斜体文字', row++, col, {
            Italic: true
        });
        sheet.AddText('下划线文字', row++, col, {
            Underline: true
        });
        sheet.AddText('粗体+斜体+下划线', row++, col, {
            Underline: true,
            Bold: true,
            Italic: true
        });
        sheet.AddText('楷体', row++, col, {
            FontName: '楷体'
        });
        sheet.AddText('字号18', row++, col, {
            FontSize: 18
        });
        sheet.AddText('蓝色字体', row++, col, {
            Color: '0000FF'
        });
        sheet.AddText('黄色背景', row++, col, {
            BGColor: 'FFFF00'
        });
        sheet.AddText('绿底红字粗体斜体下划线13号隶书', row++, col, {
            Color: 'FF0000',
            BGColor: '00FF00',
            FontName: '楷体',
            FontSize: 13,
            Bold: true,
            Italic: true,
            Underline: true
        });

        var link = book.CreateHyperlink('https://github.com');
        sheet.AddText('单元格样式超链接github.com', row++, col, {
            Color: 'FF0000',
            FontSize: 21
        }).Hyperlink = link;
        sheet.AddText('默认样式超链接appinn.com', row++, col).Hyperlink = book.CreateHyperlink('https://www.appinn.com');
        //var link = book.CreateHyperlink('https://github.com');
        sheet.AddText('超链接默认样式覆盖单元格样式v2ex.com', row++, col, {
            //超链接默认样式中包含Color和Underline=true,以下两个会被覆盖，FontSize保留
            Color: 'FF0000',
            Underline: false,
            FontSize: 16
        }).Hyperlink = book.CreateHyperlink('https://www.v2ex.com');

        (sheet.AddText('自定义超链接样式覆盖单元格样式v2ex.com', row++, col, {
            Color: 'FF0000',
            Underline: false,
            FontSize: 16
        }).Hyperlink = book.CreateHyperlink('https://www.v2ex.com')).Style = {
                FontName: '黑体',
                Italic: true,
                Color: '00FF00'
            };

        col = 1;
        row = 0;
        //设置B列宽度
        sheet.SetColWidth(1, 30);
        sheet.AddText('水平左对齐', row++, col, {
            Alignment: {
                Horizontal: HorizontalAlignment.Left
            }
        });
        sheet.AddText('水平居中对齐', row++, col, {
            Alignment: {
                Horizontal: HorizontalAlignment.Center
            }
        });
        sheet.AddText('水平右对齐', row++, col, {
            Alignment: {
                Horizontal: HorizontalAlignment.Right
            }
        });
        sheet.SetRowHeight(row, 40);
        sheet.AddText('垂直顶对齐', row++, col, {
            Alignment: {
                Vertical: VerticalAlignment.Top
            }
        });
        sheet.SetRowHeight(row, 40);
        sheet.AddText('垂直居中对齐', row++, col, {
            Alignment: {
                Vertical: VerticalAlignment.Center
            }
        });
        sheet.SetRowHeight(row, 40);
        sheet.AddText('垂直底对齐', row++, col, {
            Alignment: {
                Vertical: VerticalAlignment.Bottom
            }
        });
        sheet.SetRowHeight(row, 40);
        sheet.AddText('水平垂直居中对齐', row++, col, {
            Alignment: {
                Vertical: VerticalAlignment.Center,
                Horizontal: HorizontalAlignment.Center
            }
        });
        sheet.AddText('不支持换行\n显示\n的文本', row++, col);
        sheet.AddText('换行\n显示\n的文本', row++, col, {
            Alignment: {
                WrapText: true
            }
        });

        col = 2;
        row = 0;
        sheet.SetColWidth(col, 20);
        //常规数字
        sheet.AddText(123, row++, col);
        sheet.AddText(1234567890123456, row++, col);
        //百分比
        sheet.AddText(0.0123456, row++, col, {
            Format: {
                Code: '0.0000%'
            }
        });
        //保留2位小数
        sheet.AddText(12.3456, row++, col, {
            Format: {
                Code: '0.00'
            }
        });

        col++;
        row = 0;
        //使用CreateShareString创建在文档中多处使用的文本，可以减少文档的体积。如果是数字请不要用ShareString
        var shareString = book.CreateShareString('这是一个共享文本');
        for (var i = 0; i < 5; i++) {
            sheet.AddText(shareString, row++, col);
        }

        //创建一个通用的样式，多处使用时，可以减少文档的体积
        var shareStyle = book.CreateShareCellStyle({
            Color: '3F3FFF',
            Italic: true
        });
        var shareString2 = book.CreateShareString('共享样式的共享文本');
        sheet.SetColWidth(col, 20);
        for (var i = 0; i < 5; i++) {
            sheet.AddText(shareString2, row++, col, shareStyle);
        }

        //合并单元格
        col++;
        row = 0;
        sheet.SetColWidth(col, 15);
        sheet.SetColWidth(col + 1, 3);
        sheet.SetColWidth(col + 2, 3);

        sheet.MergeCell(row, col, row + 2, col);
        var s2 = book.CreateShareCellStyle({ Alignment: { Vertical: VerticalAlignment.Center, Horizontal: HorizontalAlignment.Center } });
        sheet.AddText('单元格上下合并', row++, col, s2);
        row += 2;

        sheet.MergeCell(row, col, row, col + 2);
        sheet.AddText('单元格左右合并', row++, col, s2);

        sheet.MergeCell(row, col, row + 2, col + 2);
        sheet.AddText('单元格上下左右合并', row++, col, s2);
        row += 2;

        //添加图片
        var img = null;
        if (typeof global != 'undefined') {
            sheet.SetRowHeight(row, 80);
            var imgData = await get('https://avatars0.githubusercontent.com/u/17020523?s=120&v=4');
            img = book.CreateImage(imgData, {
                Format: 'jpg',
                Type: 'buffer'
            });
            sheet.AddImage(img, row++, col, 66, 66);
            sheet.AddImage(img, row++, col, 33, 22);
        }

        //添加第二页
        var sheet2 = book.CreateSheet('第二页');
        row = 0;
        col = 0;
        sheet2.AddText('第二页第一个文本', row++, col);

        sheet2.SetRowHeight(row, 30);
        sheet2.AddText('加点样式', row++, col, {
            Bold: true,
            FontSize: 22
        });
        //把第一页的共享文本添加到第二页
        sheet2.AddText(shareString, row++, col);

        //把第一页的图片添加到第二页
        if (typeof global != 'undefined') {
            sheet2.AddImage(img, row++, col, 33, 22);
        }

        return book;
    }

    function get(url) {
        return new Promise((resolve, reject) => {
            var data = new Buffer(0);
            request.get(url).on('data', (d) => {
                //console.log(d);
                data = Buffer.concat([data, d]);
            }).on('end', () => {
                //console.log(data);
                resolve(data);
            }).on('error', (error) => {
                reject(error);
            });;
        })
    }


    if (typeof global == 'undefined') {
        window.xlsxDemo = xlsxDemo;
    } else {
        module.exports = xlsxDemo;
    }
})();