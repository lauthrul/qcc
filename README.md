# qcc
企查查招标信息爬取，并导出到excel中

### 1. 首先获取 win_tid：
1. 调试模式找到文件 https://qcc-static.qichacha.com/qcc/pc-web/prod-23.02.70/common-88d3322f.80048e3e.js
2. 查找 `e.headers[i] = l` 处下断点
3.  打开数据页面，执行到断点后，在控制台打印：`(0,s.default)()` 的值即为win_tid

### 2. 获取 cookie

### 3. 替换`main.py`中的变量`win_tid`和`cookie`
爬取数据：
`.\venv\Scripts\python.exe .\main.py run 1 2`

清理数据：
`.\venv\Scripts\python.exe .\main.py clean`
