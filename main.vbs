' main.vbs
Option Explicit

' ���������� ���������� ����������
Dim cnME, OPCServer, OPCGroup, OPCItems

' ������� ��� ��������� ������
Sub IncludeFile(fSpec)
    Dim fso, f
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.OpenTextFile(fSpec)
    ExecuteGlobal f.ReadAll()
    f.Close
End Sub

' �������� ��� ����������� �����
IncludeFile "databaseFunctions.vbs"
IncludeFile "opcFunctions.vbs"
IncludeFile "utilityFunctions.vbs"
IncludeFile "fileOperations.vbs"

' �������� ���������
Sub Main()
    If ConnectToDatabase() Then
        WScript.Echo "����������� � ���� ������ �������"
        
        Dim tagList
        tagList = GetTagList()
        
        If Not IsNull(tagList) Then
            WScript.Echo "�������� �����: " & UBound(tagList, 2) + 1
            
            If ConnectToOPCServer() Then
                WScript.Echo "����������� � OPC ������� �������"
                WScript.Echo
                
                ' ������� ��������� �������
                WScript.Echo "���                  | ��������     | �������� | ��������� �����    | ��������                                                  | �����"
                WScript.Echo String(140, "-")
                
                Dim i, tagName, tagValue
                For i = 0 To UBound(tagList, 2)
                    tagName = tagList(0, i)
                    tagValue = ReadOPCTag(tagName)
                    WScript.Echo PadRight(tagName, 20) & " | " & _
                               PadRight(CStr(tagValue(0)), 12) & " | " & _
                               PadRight(CStr(tagValue(1)), 8) & " | " & _
                               PadRight(CStr(tagValue(2)), 19) & " | " & _
                               PadRight(tagList(1, i), 60) & " | " & _
                               tagList(2, i)
                Next
                
                ' ���������� ������� �������� � ����
                WriteCurrentValuesToFile tagList
                
                ' ���������� �������� ������ � �����
                WriteArchiveDataToFiles tagList
                
                OPCServer.Disconnect
                WScript.Echo "���������� �� OPC ������� ���������"
            Else
                WScript.Echo "�� ������� ������������ � OPC �������"
            End If
        Else
            WScript.Echo "�� ������� �������� ������ ����� �� ���� ������"
        End If
        
        cnME.Close
        Set cnME = Nothing
        WScript.Echo "���������� � ����� ������ �������"
    Else
        WScript.Echo "�� ������� ������������ � ���� ������"
    End If
End Sub

' ������ �������� ���������
Main