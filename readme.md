一个非常好用的excel合并脚本

# 场景
你现在收集上了很多很多调查问卷，令人遗憾的是，这些每份问卷都保存在一个excel文件中。现在你的上司想让你将这些问卷合并到一个excel中，你非常苦恼。但是有了python的辅助，这些都不是问题。按照此框架模板开发属于自己的程序，只需要几分钟就可以完成别人数小时的工作，从此告别加班。

# Requirement
参见文件
[requirements.txt](https://github.com/duowen1/excel_merge/blob/master/requirements.txt)

## 有关于xlrd
需要注意，xlrd在1.2.0版本后不再支持.xlsx文件，如果已经安装了高级版本并且想处理.xlsx文件需要首先卸载后再安装。方法如下：
```shell
pip3 uninstall xlrd
pip3 install xlrd==1.2.0
```

# 使用
```shell
cd project
pip3 install requirements
python do.py
```
