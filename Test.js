WScript.Echo("And this shit`s working");
//Создаем коллекцию файлов
Files=new Enumerator(Folder.Files);
//Цикл по всем файлам
for (; !Files.atEnd(); Files.moveNext()) 
  //Добавляем строку с именем файла
  s+=Files.item().Name+"\n";
//Выводим полученные строки на экран