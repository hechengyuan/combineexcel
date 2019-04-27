## 代码使用方法

### 依赖环境

程序依赖`python3`和相关的3个包`xlrd, xlwt, xlutils`，需要检查是否安装python3，若未安装请先在[Python安装官方指引网站](https://docs.python.org/zh-cn/3/using/index.html)中查看如何安装，安装后执行：

```
pip3 install xlrd xlwt xlutils
```
安装完成即可开始运行。

### 功能模块

整个项目分为3个模块，`/cookie`存放运行出的数据，`module.py`写入的是项目执行依赖的组件，`combine.py`是主程序，每次使用，请直接在当前文件夹下运行`combine.py`。

```
cd /this_directory_path
python3 combine.py
```

### 需要哪些文件？放在哪儿？

需要日常记录的文件`订餐备注.xlsx`和在cbdiet后台下载的第二天的订餐表，并放在项目文件夹下，日常每日放好后，6点后即可运行程序。项目目录为：

```
README.md                cookie                   订餐备注.xlsx
__pycache__              module.py
combine.py               订餐表YEAR-month-day.xlsx
```

**请一定注意文件名不要更改，否则程序无法识别。**
