' Функция для форматирования строки до заданной длины
Function PadRight(str, length)
    PadRight = Left(str & Space(length), length)
End Function