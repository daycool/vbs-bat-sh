NameSpace = "http://schemas.microsoft.com/cdo/configuration/"'这个必须有，应该是VBS脚本链接微软网站获取某些支持应用的，删除的话vbs脚本会报错！
set Email = CreateObject("CDO.Message")'调用vbs邮件接口
Email.From = "qmw920@163.com" '发信人地址
Email.To = "253568176@qq.com" '收信人地址（qq邮箱也可）
Email.Subject = "test" '邮件主题
'x="d:\邮件.txt" '发信内容写在d:\邮件.txt中
'y="d\莫言小说.txt" '这是需发送的附件D盘的txt文档（也可以是其他附件，不要太大！）。
'Set fso=CreateObject("Scripting.FileSystemObject")’下面一般都是一些接口函数的调用，不做解释
'Set myfile=fso.OpenTextFile(x,1,Ture)
'c=myfile.readall
'myfile.Close
Email.Textbody = "你好啊我学习vbs改善邮件呢"
'Email.AddAttachment y
with Email.Configuration.Fields
	.Item(NameSpace&"sendusing") = 2
	.Item(NameSpace&"smtpserver") = "smtp.163.com" '这是163邮箱服务器地址，qq邮箱等请自行填写smtp地址
	.Item(NameSpace&"smtpserverport") = 25
	.Item(NameSpace&"smtpauthenticate") = 1
	.Item(NameSpace&"sendusername") = "qmw920" '发信人用户名
	.Item(NameSpace&"sendpassword") = "password" '发信人密码，也就是qmw920@163.com的邮箱密码！
	.Update
end with
Email.Send
Set Email=Nothing
