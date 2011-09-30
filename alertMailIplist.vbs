Class AlertMail
    Public Function PingCheck()
        Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
        ipList = GetIp()
        If ipList <> "" Then
            For Each ip in Split(ipList, ",")
                Set colItems = objWMIService.ExecQuery ("Select * from Win32_PingStatus " & "Where Address = '" & ip & "'")
                For Each objItem in colItems
                    If objItem.StatusCode <> 0 Then
                        SendMail(ip)
                    End if
                Next
                Set colItems = Nothing
            Next
        End If
        Set objWMIService = Nothing
    End Function
 
    Private Sub SendMail(ip)
        strServer = "smtp.mail.yahoo.co.jp:587"
        strTo = "送信先メールアドレス"
        strFrom = "アカウント名@yahoo.co.jp" & vbTab & "アカウント名:パスワード" & vbTab & "PLAIN"
        strFile = ""
        Set bobj = CreateObject("basp21")
        result = bobj.SendMail(strServer,strTo,strFrom,"Home Server Alert","Ping failed " & ip,strFile)
        Set bobj = Nothing
    End Sub
 
    Private Function GetIp()
        GetIp = ""
        strPath = Left(WScript.ScriptFullName,Len(WScript.ScriptFullName)-Len(WScript.ScriptName))
        inputFile = strPath & "alertMailIplist.ini"
        Set fsObj = WScript.CreateObject("Scripting.FileSystemObject")
        Set OpenTextFileObj = fsObj.OpenTextFile(inputFile,1)
        GetIp =  OpenTextFileObj.ReadLine
        Set fsObj = Nothing
        Set OpenTextFileObj = Nothing
    End Function
End Class
 
Set objAlertMail = New AlertMail
objAlertMail.PingCheck()
Set objAlertMail = Nothing