# java-poi

## 自动刷格式程序说明

  自动格式化Word文档格式程序使用Java为基础语言，使用开源组件poi对Office Word文档结构进行解析，对格式不正确的内容予以纠正，实现效果如下：

 ![img](https://raw.githubusercontent.com/land007/java-poi/master/image/微信图片_20200410124809.png)![img](https://raw.githubusercontent.com/land007/java-poi/master/image/微信图片_20200410124814.png)

### 一、目前提供了三种方式方便大家使用

#### （一）客户端程序（需安装Java）

![img](https://raw.githubusercontent.com/land007/java-poi/master/image/微信图片_20200410124819.png)

*软件下载地址*

```
https://github.com/land007/java-poi/releases/download/v1.2/WORD.V1.2.zip
```

*测试用WORD下载地址*

```
https://github.com/land007/java-poi/raw/master/test/%E6%8A%80%E6%9C%AF%E6%96%87%E4%BB%B6.docx
```

*使用方法*

（1）打开format.exe程序，显示WORD格式化文档窗体

（2） 将需要刷格式文档拖动到WORD格式化文档窗体内，松开鼠标

（3）程序在当前文件夹生成另一个格式化好格式的文档

（4）点击格式化好格式的文档查看

#### （二）服务器程序（需安装Docker）

*安装docker，运行镜像*

```bash
docker run -it --restart=always -v c:/Users/jiayq/docker/data:/data -p 20022:22 -p 3000:3000 --log-opt max-size=1m --log-opt max-file=1 --name poi land007/java-poi:latest
```

*打开测试docker服务器地址*

```
http://127.0.0.1:3000
```

![img](https://raw.githubusercontent.com/land007/java-poi/master/image/微信图片_20200410124825.png)

*使用方法*

（1）选择需要刷格式的文档，点击“提交”按钮，文件被上传到服务器刷新格式。

（2）当浏览器提示下载文件，选择好保存路径后点击“保存”按钮

（3） 打开格式化好格式的文档查看

#### （三）在线服务（无需安装）

使用浏览器访问https://docx.qhkly.com:6060/ ，使用过程如（二）中所示。

### 二、注意事项

（一）   需要定义好文档目录层级，程序会标准化该层级的序号。


（二）   程序为特定结构设计，仅适用于编写标书格式内容，不能解决非标书各种格式，如需其他格式文档批量格式需重新定义和编写代码。


（三）  使用客户端的同学，将需要格式化的word文档拖入此窗体前，请注意保存原文档防止不当操作丢失。程序覆盖规则如下：

```
安全    原.docx -> 新.docx.docx
覆盖    原.docx.docx -> 新.docx.docx
```

### 三、版本升级内容

 ```
  2020/01 v1.1修复v1.0图片居中缩进问题；

  2020/01 v1.1修复v1.0正文（1）a）缩进问题；

  2020/01 v1.1修复v1.0中文件锁定问题；

  2020/04 v1.2修复v1.1表格内数字丢失问题；

  2020/07 v1.3重写软件代码，支持更广泛格式；
 ```

### 四、问题建议

如有问题建议，请在下面地址反馈：

   https://github.com/land007/java-poi/issues

