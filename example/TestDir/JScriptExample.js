jsIncludeGear = new ActiveXObject("JScriptIncludeGear"); // Подключаем компоненту экспортирующую код механизма.
eval(jsIncludeGear.getSourceCode()); // Получение и инициализация кода механизма
 
jsImport.WriteCache = true; // Устанавливаем необходимость записи испортируемого кода в кеш-файл.
jsImport.require("7zip") // Подключение базовго модуля 7zip 
jsImport.require("sureFSO")  // Подключение базового модуля sureFSO
jsImport.require("\\EmbeddedDir\\embeddedScript.js") // Подключение произвольного скрипта по относительному пути, расположенного вглубь по дереву каталогов.
jsImport.require("..\\outsideScript.js") //Подключение произвольного скрипта по относительному пути, расположенного вверх по дереву каталогов.

eval(jsImport.initialization()); //Инициализация загруженного кода.

var Arch = new jlib7zaArchivator(); // Новый экземпляр базоваго класса экмпортируемого модулем 7zip
var FSO = new sureFSO(); // Новый экземпляр базоваго класса экмпортируемого модулем sureFSO

Arch.setDisplayMode('minimazed'); // Устанавливаем режим отображения вывода архиватора в консоли.

// Добавлям файлы к списку предстоящей архивации
Arch.toList("TestFile1.txt");
Arch.toList("TestFile2.txt");

// Запускаем архивацию
var newArch = Arch.Archivate("list", "TestArch_MM_DD");

//Обращаем внимание пользователя на проделанную полезную работу по тестовой упаковке при помощи модуля 7zip
WScript.Echo("[ENG]\nPay attention to the created archive " + newArch + "\nNow it will be deleted..."
+ "\n[RUS]\nОбратите внимание на созданный архив " + newArch + "\nСейчас он будет удален..." ); 

FSO.setOnErrorsMode("ignore"); // Устанавливаем режим игнорирования ошибок для работы с файловой системой через модуль sureFSO

//Преднамеренно генерируем исключение
FSO.DeleteFile(newArch + "[ENG]this is done to raise an exception / [RUS]Это сделано специально чтобы вызвать исключение");

//Удаляем созданный нами архив
FSO.DeleteFile(newArch);

// Демонстрируем доступность переменной объявленной в скрипте "\\EmbeddedDir\\embeddedScript.js"
WScript.Echo(embeddedScriptVar);

// Создаем экземпляр класса экспортируемого подключенным скриптом "..\\outsideScript.js"
var outsideScriptVar = new TestClass("outsideScriptVar","TestClass");

// Демонстируем досткпность сущьности объявленной в скрипте "..\\outsideScript.js"
WScript.Echo(outsideScriptVar.GetText())

// Демонстрируем собраные объектом FSO логи
WScript.Echo(FSO.logoutToString())





