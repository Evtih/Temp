' Функция для подключения к OPC серверу
Function ConnectToOPCServer()
    On Error Resume Next
    Set OPCServer = CreateObject("OPC.Automation.1")
    OPCServer.Connect "OPCServer.WinCC.1", "192.168.0.102"
    If Err.Number <> 0 Then
        WScript.Echo "Ошибка при подключении к OPC серверу: " & Err.Description
        ConnectToOPCServer = False
    Else
        Set OPCGroup = OPCServer.OPCGroups.Add("Group1")
        OPCGroup.UpdateRate = 1000
        OPCGroup.IsActive = True
        Set OPCItems = OPCGroup.OPCItems
        ConnectToOPCServer = True
    End If
    On Error Goto 0
End Function

' Функция для чтения значения тега
Function ReadOPCTag(tagName)
    Dim Item, Value, Quality, TimeStamp
    Set Item = OPCItems.AddItem(tagName, 1)
    Item.Read 1, Value, Quality, TimeStamp
    ReadOPCTag = Array(Value, Quality, TimeStamp)
End Function