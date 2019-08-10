'*******************************************************************
' Имя: ExecWinApp.vbs
' Язык: VBScript
' Описание: Запуск и закрытие приложение (объект WshScriptExec)
'*******************************************************************
Option Explicit

Dim WshShell,theNotepad,Res,Text,Title   ' Объявляем переменные
' Создаем объект WshShell
Set WshShell = WScript.CreateObject("WScript.Shell")
WScript.Echo "Runnning notepad..."
' Запускаем приложение (создаем объект WshScriptExec)
Set theNotepad = WshShell.Exec("notepad")
WScript.Sleep 500   ' Приостанавливаем выполнение сценария
Text="Notepad is running (Status=" & theNotepad.Status & ")" & vbCrLf _
      & "Close notepad?"
Title=""
' Выводим диалоговое окно на экран
Res=WshShell.Popup(Text,0,Title,vbQuestion+vbYesNo)
' Определяем, какая кнопка нажата в диалоговом окне
If Res=vbYes Then
  theNotepad.Terminate ' Прерываем работу Блокнота
  ' Приостанавливаем выполнение сценария, для того чтобы Блокнот
  ' успел закрыться
  WScript.Sleep 100
  WScript.Echo "Notepad is closed (Status=" & theNotepad.Status & ")"
End If
'*************  Конец *********************************************/