(() => {
    if (typeof global == 'undefined') {
        window.require = function () { };
    }
    let request = require('request');
    let XlsxCore = require('js-xlsx-core');
    require('xlsx-saver');
    if (typeof global == 'undefined') {
        XlsxCore = window.XlsxCore;
    }
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

    async function xlsxDemo() {
        try {
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
                Font: {
                    Bold: true
                }
            });
            sheet.AddText('斜体文字', row++, col, {
                Font: {
                    Italic: true
                }
            });
            sheet.AddText('下划线文字', row++, col, {
                Font: {
                    Underline: true
                }
            });
            sheet.AddText('粗体+斜体+下划线', row++, col, {
                Font: {
                    Underline: true,
                    Bold: true,
                    Italic: true
                }
            });
            sheet.AddText('楷体', row++, col, {
                Font: {
                    FontName: '楷体'
                }
            });
            sheet.AddText('字号18', row++, col, {
                Font: {
                    FontSize: 18
                }
            });
            sheet.AddText('蓝色字体', row++, col, {
                Font: {
                    Color: '0000FF'
                }
            });
            sheet.AddText('黄色背景', row++, col, {
                Fill: {
                    BGColor: 'FFFF00'
                }
            });
            sheet.AddText('绿底红字粗体斜体下划线13号隶书', row++, col, {
                Font: {
                    Color: 'FF0000',
                    FontName: '楷体',
                    FontSize: 13,
                    Bold: true,
                    Italic: true,
                    Underline: true
                },
                Fill: {
                    BGColor: '00FF00',
                }
            });

            sheet.AddText('默认样式超链接appinn.com', row++, col).Hyperlink = book.CreateHyperlink('https://www.appinn.com');
            var link = book.CreateHyperlink('https://github.com');
            sheet.AddText('红色21号字体超链接github.com', row++, col, {
                Font: {
                    Color: 'FF0000',
                    FontSize: 21
                }
            }).Hyperlink = link;
            sheet.AddText('绿色无下划线16号字体超链接v2ex.com', row++, col, {
                //超链接默认样式中包含Color和Underline=true,以下两个覆盖默认样式
                Font: {
                    Color: '00FF00',
                    Underline: false,
                    FontSize: 16
                }
            }).Hyperlink = book.CreateHyperlink('https://www.v2ex.com');

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
            sheet.AddText('换行\n\n显示\n的文本', row++, col, {
                Alignment: {
                    WrapText: true
                }
            });
            sheet.AddText('不支持换行\n显示\n的文本', row++, col);

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
            sheet.SetColWidth(col, 20);
            //使用CreateShareString创建在文档中多处使用的文本，可以减少文档的体积。如果是数字请不要用ShareString
            var shareString = book.CreateShareString('这是一个共享文本');
            for (var i = 0; i < 3; i++) {
                sheet.AddText(shareString, row++, col);
            }

            //创建一个共享字体，多处使用时，可以减少文档的体积
            var shareFont = book.CreateShareCellFont({
                FontName: '微软雅黑',
                Bold: true
            });
            shareString = book.CreateShareString('共享字体加粗微软雅黑');
            for (var i = 0; i < 3; i++) {
                sheet.AddText(shareString, row++, col, { Font: shareFont });
            }

            //创建一个通用的样式，多处使用时，可以减少文档的体积
            var shareStyle = book.CreateShareCellStyle({
                Font: {
                    Color: '3F3FFF',
                    Italic: true
                }
            });
            shareString = book.CreateShareString('共享样式的共享文本');
            for (var i = 0; i < 3; i++) {
                sheet.AddText(shareString, row++, col, shareStyle);
            }

            //最佳实践：创建通用样式，使用共享字体、共享填充，最大限度减少文档的体积。
            shareFont = book.CreateShareCellFont({
                Color: '2222FF'
            });
            var shareFill = book.CreateShareCellFill({ BGColor: 'E2E2E2' });
            shareStyle = book.CreateShareCellStyle({
                Font: shareFont,
                Fill: shareFill,
                Alignment: { WrapText: true }
            });
            shareString = book.CreateShareString('由共享字体+共享填充组成的共享样式');
            for (var i = 0; i < 3; i++) {
                sheet.SetRowHeight(row, 45);
                sheet.AddText(shareString, row++, col, shareStyle);
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
                var imgData = await get('https://avatars0.githubusercontent.com/u/17020523?s=120&v=4');
                img = book.CreateImage(imgData, {
                    Format: 'jpg',
                    Type: 'buffer'
                });
            } else {
                var imgData = '/9j/2wCEAAgGBgcGBQgHBwcJCQgKDBQNDAsLDBkSEw8UHRofHh0aHBwgJC4nICIsIxwcKDcpLDAxNDQ0Hyc5PTgyPC4zNDIBCQkJDAsMGA0NGDIhHCEyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMv/AABEIAB4AHgMBIgACEQEDEQH/xAGiAAABBQEBAQEBAQAAAAAAAAAAAQIDBAUGBwgJCgsQAAIBAwMCBAMFBQQEAAABfQECAwAEEQUSITFBBhNRYQcicRQygZGhCCNCscEVUtHwJDNicoIJChYXGBkaJSYnKCkqNDU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6g4SFhoeIiYqSk5SVlpeYmZqio6Slpqeoqaqys7S1tre4ubrCw8TFxsfIycrS09TV1tfY2drh4uPk5ebn6Onq8fLz9PX29/j5+gEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoLEQACAQIEBAMEBwUEBAABAncAAQIDEQQFITEGEkFRB2FxEyIygQgUQpGhscEJIzNS8BVictEKFiQ04SXxFxgZGiYnKCkqNTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqCg4SFhoeIiYqSk5SVlpeYmZqio6Slpqeoqaqys7S1tre4ubrCw8TFxsfIycrS09TV1tfY2dri4+Tl5ufo6ery8/T19vf4+fr/2gAMAwEAAhEDEQA/AMMXOcAcnpWvFo2rtaNd/YZhCvJZhtOPXB5x71oeHtW0Sz0KzvTNbWtwqMkoJG5nBxk4BcgjBwMDn8um0HUzr1jc7DO1oq+UjtHtDDBB2gkk/Un8OteXmPEuKw7k6dC0IuzlK+uttLaX+b7mVHBwklzS1fRHNXMSHSIZIbRhLENs8mwjkE5z+lZ8UpPA5rvX8O2zWa26tchDndyMk9yT36/jXD6pYS6HfvESxjP+rkPcen1rsybiGhmFSVFO0t0u6/z/AEIxWFlSipfecK1kqnBaRP8AeSu18CeJodASe0v5We3cgxumW2dcjHoc54759a0heeH5Dh9MljJ/55kY/nWpbaRpd3ARaxmPcMfvIUP8q9LM8NSxuFlQrLR/mtjDDVfZ1Yyvp19OprSfEHw6bRCqToWwBKY3wSOuOMdq898Q6pLreqvcQo4twAqKT198dif8Ktp4P1CFoVnvYXittzNHglQASRgEYOST1qlJ4eCHCXb491/+vXj5Hg6Ptnin8SSS2WjV9klunue3xAsLScaWEm5J3v8AJ27Lsf/Z';
                img = book.CreateImage(imgData, {
                    Format: 'png',
                    Type: 'base64'
                });
            }
            sheet.SetRowHeight(row, 80);
            sheet.AddImage(img, row++, col, 66, 66);
            sheet.AddImage(img, row++, col, 33, 22);

            //添加第二页
            var sheet2 = book.CreateSheet('第二页');
            row = 0;
            col = 0;
            sheet2.AddText('第二页第一个文本', row++, col);

            sheet2.SetRowHeight(row, 30);
            sheet2.AddText('加点样式', row++, col, {
                Font: {
                    Bold: true,
                    FontSize: 22
                }
            });
            //把第一页的共享文本添加到第二页
            sheet2.AddText(shareString, row++, col);

            //把第一页的图片添加到第二页
            sheet2.AddImage(img, row++, col, 33, 22);

            book.CreateSheet('第二页前面插入一页', 1);

            return book;

        } catch (error) {
            console.error('Demo出现错误', error);
            return null;
        }
    }

    function get(url) {
        return new Promise((resolve, reject) => {
            var data = [];
            request.get(url).on('data', (d) => {
                data.push(d);
            }).on('end', () => {
                buf = Buffer.concat(data);
                resolve(buf);
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