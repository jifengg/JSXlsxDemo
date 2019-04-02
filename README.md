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

# important

代码使用的ES6的新特性（箭头函数，async/await，Promise，析构等），在不支持ES6的node或者浏览器上无法使用。
