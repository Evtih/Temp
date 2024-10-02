' fileOperations.vbs
Option Explicit

' ������� ��� �������������� ����
Function FormatDate(dateValue)
    FormatDate = Right("0" & Day(dateValue), 2) & "." & _
                 Right("0" & Month(dateValue), 2) & "." & _
                 Year(dateValue)
End Function

' ������� ��� �������������� �������
Function FormatTime(dateValue)
    FormatTime = Right("0" & Hour(dateValue), 2) & ":" & _
                 Right("0" & Minute(dateValue), 2) & ":" & _
                 Right("0" & Second(dateValue), 2)
End Function

' ������� ��� ������ ������� �������� ����� � ����
Sub WriteCurrentValuesToFile(tagList)
    Dim fso, file, i, line
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.CreateTextFile("Z:\Kotel.opc", True)
    
    line = "Req~" & FormatDate(Now) & " " & FormatTime(Now) & "~SyncRead~"
    
    For i = 1 To UBound(tagList, 2)
        line = line & "~" & tagList(0, i) & "~" & ReadOPCTag(tagList(0, i))(0) & "~" & ReadOPCTag(tagList(0, i))(1) & "~" & ReadOPCTag(tagList(0, i))(2)
    Next
    
    line = line & "~||"
    file.WriteLine line
    file.Close
    Set file = Nothing
    Set fso = Nothing
End Sub

' ������� ��� ������ �������� ������ � �����
Sub WriteArchiveDataToFiles(tagList)
    Dim fso, file, i, tagName, tagValue, line, fileName
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    For i = 1 To UBound(tagList, 2)
        tagName = tagList(0, i)
        If tagList(2, i) = "A" Then ' ���������, ����� �� ������������ ���� ���
            fileName = "Z:\" & tagName & ".txt"
            tagValue = ReadOPCTag(tagName)
            
            ' ��������� ������ � �������
            line = FormatDate(Now) & "~" & Hour(Now) & "~" & Int(Minute(Now) / 30) + 1 & "~" & FormatDate(Now) & "~" & tagValue(0)
            
            ' ���������� ������ � ����
            Set file = fso.OpenTextFile(fileName, 8, True) ' 8 - ForAppending
            file.WriteLine line
            file.Close
        End If
    Next
    
    Set fso = Nothing
End Sub