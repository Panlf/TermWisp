## 脚本处理

### Linux

- Shell 处理 MySQL 数据库 ( operate_db.sh )
- Shell 替换文件夹下所有文件的某一行数据 ( operate_file.sh )
- Shell 实时将 nginx 日志解析并保存到 MySQL ( nginx_log_to_mysql.sh )
- Shell 处理Nginx日志统计ip调用量（nginx_ip_count_offset.sh、nginx_ip_count_time.sh）

### Windows

- 获取当前文件夹占用空间 ( Get-ItemSize.ps1 )
- 验证身份证号是否符合规则 ( Validate-IDCard.ps1 )
	- 使用方式 `.\Validate-IDCard.ps1 -IDCard "xxxxxxx"`
- 更换Java的配置地址 ( Change-Java-Version.ps1 )
- 查找`pom.xml`文件中是否包含`ebean`字符串( FindStr.ps1 )

### Python
- 基于Funasr将video提取文本的python脚本 ( funasr_video_to_text.py )
- 扫描Windows下面是否还有敏感数据 ( scan-excel-files.py )
