jsIncludeGear = new ActiveXObject("JScriptIncludeGear"); // ���������� ���������� �������������� ��� ���������.
eval(jsIncludeGear.getSourceCode()); // ��������� � ������������� ���� ���������
 
jsImport.WriteCache = true; // ������������� ������������� ������ �������������� ���� � ���-����.
jsImport.require("7zip") // ����������� ������� ������ 7zip 
jsImport.require("sureFSO")  // ����������� �������� ������ sureFSO
jsImport.require("\\EmbeddedDir\\embeddedScript.js") // ����������� ������������� ������� �� �������������� ����, �������������� ������ �� ������ ���������.
jsImport.require("..\\outsideScript.js") //����������� ������������� ������� �� �������������� ����, �������������� ����� �� ������ ���������.

eval(jsImport.initialization()); //������������� ������������ ����.

var Arch = new jlib7zaArchivator(); // ����� ��������� �������� ������ ��������������� ������� 7zip
var FSO = new sureFSO(); // ����� ��������� �������� ������ ��������������� ������� sureFSO

Arch.setDisplayMode('minimazed'); // ������������� ����� ����������� ������ ���������� � �������.

// �������� ����� � ������ ����������� ���������
Arch.toList("TestFile1.txt");
Arch.toList("TestFile2.txt");

// ��������� ���������
var newArch = Arch.Archivate("list", "TestArch_MM_DD");

//�������� �������� ������������ �� ����������� �������� ������ �� �������� �������� ��� ������ ������ 7zip
WScript.Echo("[ENG]\nPay attention to the created archive " + newArch + "\nNow it will be deleted..."
+ "\n[RUS]\n�������� �������� �� ��������� ����� " + newArch + "\n������ �� ����� ������..." ); 

FSO.setOnErrorsMode("ignore"); // ������������� ����� ������������� ������ ��� ������ � �������� �������� ����� ������ sureFSO

//������������� ���������� ����������
FSO.DeleteFile(newArch + "[ENG]this is done to raise an exception / [RUS]��� ������� ���������� ����� ������� ����������");

//������� ��������� ���� �����
FSO.DeleteFile(newArch);

// ������������� ����������� ���������� ����������� � ������� "\\EmbeddedDir\\embeddedScript.js"
WScript.Echo(embeddedScriptVar);

// ������� ��������� ������ ��������������� ������������ �������� "..\\outsideScript.js"
var outsideScriptVar = new TestClass("outsideScriptVar","TestClass");

// ������������ ����������� ��������� ����������� � ������� "..\\outsideScript.js"
WScript.Echo(outsideScriptVar.GetText())

// ������������� �������� �������� FSO ����
WScript.Echo(FSO.logoutToString())





