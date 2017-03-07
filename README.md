## pylog - 解析分子日志文件

- Python3 xlwt包
- 命令行启动 -w 日志目录
- 解析全部.log文件 生成excel表格
- 可以直接传递日志文件作为命令行参数
- 也可在日志目录下建立config.log文件，将要解析的log名称写入(一行一个文件)

### 不足之处
- xlwt包只支持65535行，过大的日志无法解析
- 代码拓展性不够


### 问题已修正
- xlwt 更换为 xlsxwriter