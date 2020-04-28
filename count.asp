<!DOCTYPE html>
<html>
<body>
<p>Visits:
<%

'赋值
Dim fsoObject                  
Dim tsObject                    
Dim filObject                   
Dim VisitorNumber         

'创建文件系统对象变量
Set fsoObject = Server.CreateObject("Scripting.FileSystemObject")
'初始化文件对象，设置路径和名称
Set filObject = fsoObject.GetFile(Server.MapPath("counter.txt"))
'打开txt
Set tsObject = filObject.OpenAsTextStream
'读取txt内容
VisitorNumber = CLng(tsObject.ReadAll)
'利用Session判断用户是否已访问过
If Session("user_is_recorded")<>"true" Then
    Session("user_is_recorded")="true"
'数据加一
VisitorNumber = VisitorNumber + 1
'创建新txt覆盖前一个
Set tsObject = fsoObject.CreateTextFile(Server.MapPath("counter.txt"))
'新数据写入新txt
tsObject.Write CStr(VisitorNumber)
'结束判断
End If
'重置服务器对象
Set fsoObject = Nothing
Set tsObject = Nothing
Set filObject = Nothing
'显示txt数据
Response.Write(VisitorNumber)

%>
</p>
</body>
</html>
