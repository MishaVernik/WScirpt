'*******************************************************************
' Имя: InputPassw.vbs
' Язык: VBScript
' Описание: Ввод пароля без отображения на экране для соединения 
'           с WMI на удаленном компьютере
'*******************************************************************
Option Explicit
 
' Объявляем переменные
Dim strComputer       ' Имя компьютера
Dim strNamespace      ' Имя пространства имен
Dim strClass          ' Имя класса 
Dim strUser           ' Имя пользователя 
Dim strPassw          ' Пароль пользователя 
Dim objPassw          ' Объект ScriptPW
Dim objLocator        ' Объект SWbemLocator
Dim objService        ' Объект SWbemServices
Dim colInstances      ' Коллекция экземпляров класса WMI
Dim objInstance       ' Элемент коллекции
Dim strComputerRole   ' Роль компьютера в домене

'********************** Начало *************************************
' Присваиваем начальные значения переменным
strComputer = "POPOV"
strNamespace = "Root\CIMV2"
strClass = "Win32_ComputerSystem"
strUser = "POPOV\404_Popov"

'Создаем объект ScriptPW
Set objPassw = CreateObject("ScriptPW.Password")
' Выводим подсказку для ввода пароля
WScript.StdOut.Write "Введите пароль для " & strUser & ": "
'Вводим пароль
strPassw = objPassw.GetPassword()

'Создаем объект SWbemLocator
Set objLocator = CreateObject("WbemScripting.SWbemLocator")
'Соединяемся с пространством имен WMI от имени заданной учетной записи
Set objService = objLocator.ConnectServer(strComputer, strNamespace, strUser, strPassw)

' Создаем коллекцию экземпляров класса Win32_ComputerSystem
Set colInstances = objService.InstancesOf(strClass)

' Перебираем элементы коллекции 
For Each objInstance In colInstances
    ' Определяем описание роли
    Select Case objInstance.DomainRole 
        Case 0 
            strComputerRole = "Standalone Workstation"
        Case 1        
            strComputerRole = "Member Workstation"
        Case 2
            strComputerRole = "Standalone Server"
        Case 3
            strComputerRole = "Member Server"
        Case 4
            strComputerRole = "Backup Domain Controller"
        Case 5
            strComputerRole = "Primary Domain Controller"
    End Select
    
    ' Выводим результат на экран
    Wscript.Echo "Роль компьютера " & strComputer & ": " & strComputerRole
Next
'************************* Конец ***********************************