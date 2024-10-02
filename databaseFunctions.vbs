' databaseFunctions.vbs
Option Explicit

' Функция для подключения к базе данных
Function ConnectToDatabase()
    On Error Resume Next
    Set cnME = CreateObject("ADODB.Connection")
    cnME.Open "Provider=SQLOLEDB.1;UID=SA;PWD=48;APP=WS_GRAFIT;Data Source=Pacient;Initial Catalog=Kotelnaya"
    If Err.Number <> 0 Then
        WScript.Echo "Ошибка при подключении к базе данных: " & Err.Description
        ConnectToDatabase = False
    Else
        ConnectToDatabase = True
    End If
    On Error Goto 0
End Function

' Функция для получения списка тегов из базы данных
Function GetTagList()
    Dim rs, sqlQuery
    Set rs = cnME.Execute("SELECT Nam, Descr, Arhiv FROM Kotelnaya.dbo.Tags")
    If Not rs.EOF Then
        GetTagList = rs.GetRows()
    Else
        GetTagList = Null
    End If
    rs.Close
    Set rs = Nothing
End Function