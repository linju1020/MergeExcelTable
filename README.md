# 基于nodejs的excel表格合并工具

## 安装或升级
该工具是Nodejs命令行工具，安装方式如下（需全局安装）
```shell
npm install -g excelmergetool
```
## 查看版本
```shell
excelmergetool --version
```
## 查看帮助
```shell
excelmergetool --help
```
## 卸载
```shell
npm uninstall -g excelmergetool
```

## 怎样使用
输入命令，选择需要合并Excel的文件夹（文件夹里面不能包含其他的任何文件）
```shell
excelmergetool 合并默认第1个工作表
excelmergetool -i 1 合并第2个工作表
excelmergetool -name Sheet2 合并工作表名称为Sheet2的所有表
```