# read_excel_for_lang
一个读取excel文档内容，生成自定义的文件的小工具

说明：前段时间，工作上需要把添加多语言，添加西班牙语。公司的这套code，在Web上多语言写的有点费事。
语言文件是是不同的页面加载不同的.js.
其实就是一个文本文件，里面的内容类似如下：
```JavaScript
LangM.push({
	'bt_hlp1':'Table of Contents',
	'bt_hlp2':'Status',
	'bt_hlp3':'My Devices',
	'bt_hlp4':'Settings',
	'bt_hlp5':'Wireless',
	'bt_hlp6':'Network',
	'bt_hlp7':'LED',
	'':null});
```
关键是很多，有差不多一百多个这类的文本，每个语言文件里，少则四五十个，多则上千个。这要一个一个对应翻译成西班牙，也是醉了。后来发现pm给的翻译文件是excel的，就想想写一个自动生成这类文件的小工具。<br>
以前的公司的是在excel嵌套一个vb写的小东西，功能类似。<br>
--------------
<br/><br/>
最近翻出了code，稍微整理下发出了，就当备份下。
