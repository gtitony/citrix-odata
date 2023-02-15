# 简介
从 OData 接口 获取 Citrix Session 日志。
需指定参数：
 - $LastDays = 90 
 - $Username = “domain/username” 
 -  $Password = ”your password“
 -  $CitrixDDCURL = http://192.1.21.101


# 截图
执行PowerShell脚本后的日志。
![截屏 71.png](https://tonywalker-blog-wordpress.oss-cn-shanghai.aliyuncs.com/resources/cc6123fc75544dfb831c68fac12b914b.png?x-oss-process=style/watermark)



获取 Session 日志后的，Excel 文件内容：
![截屏 75.png](https://tonywalker-blog-wordpress.oss-cn-shanghai.aliyuncs.com/resources/8b3a152dc5c14ec19f60ac967838b86a.png?x-oss-process=style/watermark)


# 设定计划任务
1. 新建计划任务 
填写计划任务名称 ：Get-CitrixSessionDetails
![截屏 72.png](https://tonywalker-blog-wordpress.oss-cn-shanghai.aliyuncs.com/resources/88b5bf585c6041e1a177e0ef87ee2d3b.png?x-oss-process=style/watermark)

2. 设定触发器，可以每隔30天，无限期循环。
![截屏 73.png](https://tonywalker-blog-wordpress.oss-cn-shanghai.aliyuncs.com/resources/00876a2e380b48c18831166a0990334b.png?x-oss-process=style/watermark)

3. 启动 Powershell 脚本：
程序或脚本：C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe
添加参数：-ExecutionPolicy Bypass -File "C:\Citrix\Get-CitrixSessionDetails.ps1"

![截屏 74.png](https://tonywalker-blog-wordpress.oss-cn-shanghai.aliyuncs.com/resources/94d3d9cbef584e32a10b058285fa2503.png?x-oss-process=style/watermark)











