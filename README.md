# JSXlsxDemo
A demo of how to use [JSXlsxSaver](https://github.com/jifengg/JSXlsxSaver)

# Start

clone代码并初始化依赖包

```shell
git clone https://github.com/jifengg/JSXlsxDemo.git
cd JSXlsxDemo
npm i
```

# run in node

```shell
node run_in_node.js
```

初始化之后，执行以上命令行，将会生成一个`out.xlsx`文件，并自动打开。

测试通过的node版本
- v10.13.0

# run in browser

使用浏览器打开`index.html`，点击按钮`make and download test.xlsx`，将生成一个xlsx文件并使用浏览器下载该文件。

测试通过的浏览器
- Firefox 65.0.2 (64bit)
- Chrome 73.0.3683.75 (64bit)
- Microsoft Edge 41.16299.967.0

很遗憾IE11浏览器不支持箭头函数导致无法测试通过。

# support 支持的操作

- 可在单元格中存放字符串、数字和图片。（`时间类型暂时不支持，请使用字符串或数字存储。`）
- 设置单元格字体样式：字体名称、字号、文字颜色、是否粗体、是否斜体、是否下划线；
- 设置单元格纯颜色填充；
- 设置单元格超链接；
- 设置单元格水平垂直对齐方式；
- 设置单元格是否支持换行；
- 合并单元格；
- 设置默认字体样式、默认行高、默认列宽；
- 设置指定行的高度（行高所使用单位为磅，1厘米=28.6磅）；
- 设置指定列的宽度（列宽使用单位为1/10英寸，既1个单位为2.54毫米）；
- 设置单元格内图片大小（图片大小不能超过单元格，所以如果要设置比较大的图片，请设置合适的单元格大小）；
- 设置单元格中数字的显示格式，如`百分比`,`千分符`等等，具体的格式码可以参照[微软文档](https://support.office.com/zh-cn/article/%E6%9F%A5%E7%9C%8B%E6%9C%89%E5%85%B3%E8%87%AA%E5%AE%9A%E4%B9%89%E6%95%B0%E5%AD%97%E6%A0%BC%E5%BC%8F%E7%9A%84%E5%87%86%E5%88%99-c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5)。
- 通过`共享文本`、`共享字体`、`共享填充`、`共享样式`、`共享格式码`、`共享图片`来达到一处定义多处使用，可以大大减少最终的文件体积。

# important

代码使用的ES6的新特性（箭头函数，async/await，Promise，析构等），在不支持ES6的node或者浏览器上无法使用。
