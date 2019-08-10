'********************************************************************
' Имя: ListUsers.vbs                                                
' Язык: VBScript                                    
' Описание: Вывод на экран имен всех пользователей заданной группы
'********************************************************************
Option Explicit

'Объявляем переменные
Dim objGroup       ' Экземпляр объекта Group
Dim objUser        ' Экземпляр объекта User
Dim strResult      ' Строка для вывода на экран

'********************** Начало *************************************
' Связываемся с группой Пользователи компьютера Popov
Set objGroup = GetObject("WinNT://Popov/Пользователи,group")

strResult = "Все пользователи группы Пользователи на компьютере Popov:" & vbCrLf

' Перебираем элементы коллекции 
For Each objUser In objGroup.Members()
  ' Формируем строку с именами пользователей
  strResult = strResult & objUser.Name & vbCrLf
Next

' Вывод информации на экран
WScript.Echo strResult
'*************  Конец *********************************************