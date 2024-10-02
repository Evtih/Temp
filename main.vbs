' main.vbs
Option Explicit

' Объявление глобальных переменных
Dim cnME, OPCServer, OPCGroup, OPCItems

' Функция для включения файлов
Sub IncludeFile(fSpec)
    Dim fso, f
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.OpenTextFile(fSpec)
    ExecuteGlobal f.ReadAll()
    f.Close
End Sub

' Включаем все необходимые файлы
IncludeFile "databaseFunctions.vbs"
IncludeFile "opcFunctions.vbs"
IncludeFile "utilityFunctions.vbs"
IncludeFile "fileOperations.vbs"

' Основная процедура
Sub Main()
    If ConnectToDatabase() Then
        WScript.Echo "Подключение к базе данных успешно"
        
        Dim tagList
        tagList = GetTagList()
        
        If Not IsNull(tagList) Then
            WScript.Echo "Получено тегов: " & UBound(tagList, 2) + 1
            
            If ConnectToOPCServer() Then
                WScript.Echo "Подключение к OPC серверу успешно"
                WScript.Echo
                
                ' Выводим заголовок таблицы
                WScript.Echo "Тег                  | Значение     | Качество | Временная метка    | Описание                                                  | Архив"
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
                
                ' Записываем текущие значения в файл
                WriteCurrentValuesToFile tagList
                
                ' Записываем архивные данные в файлы
                WriteArchiveDataToFiles tagList
                
                OPCServer.Disconnect
                WScript.Echo "Отключение от OPC сервера выполнено"
            Else
                WScript.Echo "Не удалось подключиться к OPC серверу"
            End If
        Else
            WScript.Echo "Не удалось получить список тегов из базы данных"
        End If
        
        cnME.Close
        Set cnME = Nothing
        WScript.Echo "Соединение с базой данных закрыто"
    Else
        WScript.Echo "Не удалось подключиться к базе данных"
    End If
End Sub

' Запуск основной процедуры
Main