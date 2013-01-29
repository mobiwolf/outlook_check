outlook_check
=============

outlook 检查标题名和附件的宏脚本

Q：为什么需要这个功能？
A: 因为有时候我们忙，或者着急，或者三心二意了，脑子里面在想要怎么说，或者发送给哪些人，人脑偶尔会出错。所以借助工具来帮忙一下也是有必要的~~

使用方法：
1. 打开outlook
2. 按“Alt + F11” 键来打开VB Script,或者[工具]->[宏]->[Visual Basic 编辑器]
3. 点击左侧树状目录最下面的“ThisOutlookSession”，看到右边出现空白的编辑窗口
4. 把check.vb文件里面的代码拷贝到编辑窗口，保存，退出VB Script编辑。
5，为了保证立即生效，最好重启下邮件客户端。

注意：
1，目前确保支持2007版本！
2，识别是否有附件是判断在标题或者正文中是否存在“附件”两个字。如果你的需求是要发附件，但是标题或者正文没有附件两个字，那么识别不了。所以，最好有附件的时候，写出来“附件”两个字
3，

默认为强制禁止在标题或者正文中有提到附件，没有附件的邮件发送出去，为解决有人不需要强制禁止缺少附件的邮件发送出去。提供了强制开关，操作如下：
1，打开check.rb文件
2，找到bForceAttch这个变量，然后将bForceAttch = True False 修改为 bForceAttch = False 保存即可

即当bForceAttch = False 时，在提示确认对话框内，可以选择没有附件也发送

最后特别提示：工具只是辅助，慢慢养成良好的习惯最重要！


