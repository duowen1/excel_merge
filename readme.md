一个非常好用的excel合并脚本，适用于各种excel文件的政审表、调查问卷、个人信息收集表等

# 场景
你现在收集上了很多很多调查问卷，令人遗憾的是，这些每份问卷都保存在一个excel文件中。现在你的上司想让你将这些问卷合并到一个excel中，你非常苦恼。但是有了python的辅助，这些都不是问题。按照此框架模板开发属于自己的程序，只需要几分钟就可以完成别人数小时的工作，从此告别加班。

## 原始问卷举例

## 运行结果举例


# 运行要求
参见文件
[requirements.txt](https://github.com/duowen1/excel_merge/blob/master/requirements.txt)

## 有关于xlrd
需要注意，xlrd在1.2.0版本后不再支持.xlsx文件，如果已经安装了高级版本并且想处理.xlsx文件需要首先卸载后再安装。方法如下：
```shell
pip3 uninstall xlrd
pip3 install xlrd==1.2.0
```

# 使用
## 1. 克隆代码
```
git clone git@github.com:duowen1/excel_merge.git
```

## 2. 修改配置文件

[titiles.txt](https://github.com/duowen1/excel_merge/blob/file/titles.txt)描述了问卷的项目以及格式，每一行代表一个调查内容。格式如下:
```
[项目名称] [行数] [列数] [可选参数]
```
- 项目名称：生成的excel表中该列的标题
- 行数：原始调查表中，数据所在的行数
- 列数：原始调查表中，数据所在的列数，请用**小写**字母表示
- 可选参数：可省略，如果该项目是一个整数则填写为i

例如：

## 3. 执行代码

```shell
cd project
pip3 install requirements
python do.py
```

# todo list
1. 增加对原始调查问卷列数大于Z时的支持；
2. 增加列数对大小写字母的支持；
3. 增加异常处理，当配置文件某行出现问题时自动跳过该项目；
4. 增加运行结果显示，展示成功汇总的数量、失败的数量；
5. 增加进度条，以直观显示运行进度；
6. 增加对原始问卷排序的功能；