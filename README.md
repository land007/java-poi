# java-poi

**自动刷格式程序说明**

  自动格式化Word文档格式程序使用Java为基础语言，使用开源组件poi对Office Word文档结构进行解析，对格式不正确的内容予以纠正，实现效果如下：

 ![img](https://raw.githubusercontent.com/land007/java-poi/master/image/微信图片_20200410124809.png)![img](https://raw.githubusercontent.com/land007/java-poi/master/image/微信图片_20200410124814.png)

**目前提供了三种方式方便大家使用**

（1）客户端程序（需安装Java）

![img](https://raw.githubusercontent.com/land007/java-poi/master/image/微信图片_20200410124819.png)

软件下载地址

https://github.com/land007/java-poi/releases/download/v1.2/WORD.V1.2.zip

测试用WORD下载地址

https://github.com/land007/java-poi/raw/master/test/%E6%8A%80%E6%9C%AF%E6%96%87%E4%BB%B6.docx

使用方法

a)    打开format.exe程序，显示WORD格式化文档窗体

b)    将需要刷格式文档拖动到WORD格式化文档窗体内，松开鼠标

c)    程序在当前文件夹生成另一个格式化好格式的文档

d)    点击格式化好格式的文档查看

（2）服务器程序（需安装Docker）

安装docker，运行镜像

```bash
docker run -it --restart=always -v c:/Users/jiayq/docker/data:/data -p 20022:22 -p 3000:3000 --log-opt max-size=1m --log-opt max-file=1 --name poi land007/java-poi:latest
```

打开测试docker服务器地址http://127.0.0.1:3000

![img](https://raw.githubusercontent.com/land007/java-poi/master/image/微信图片_20200410124825.png)

使用方法

a)    选择需要刷格式的文档，点击“提交”按钮，文件被上传到服务器刷新格式。

b)    当浏览器提示下载文件，选择好保存路径后点击“保存”按钮

c)    打开格式化好格式的文档查看

（3）在线服务（无需安装）

访问https://docx.qhkly.com:6060/，使用过程如（2）中所示。

**注意事项**

​    需要定义好文档目录层级，程序会标准化该层级的序号。

​    程序为特定结构设计，仅适用于编写标书格式内容，不能解决非标书各种格式，如需其他格式文档批量格式需重新定义和编写代码。

​    将需要格式化的word文档拖入此窗体，会在当前位置生成格式化后的文档，需注意保存原文档，防止不当操作丢失。

​           新的    原.docx -> 新.docx.docx

​           覆盖    原.docx.docx -> 新.docx.docx

**版本升级内容**

  2020/01 v1.1修复v1.0图片居中缩进问题；

  2020/01 v1.1修复v1.0正文（1）a）缩进问题；

  2020/01 v1.1修复v1.0中文件锁定问题；

  2020/04 v1.2修复v1.1表格内数字丢失问题；



 如有问题建议，请在下面地址反馈：

   https://github.com/land007/java-poi/issues

