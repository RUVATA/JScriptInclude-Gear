InstallationPath = WScript.ScriptFullName.replace(/[^\\]+$/, ""); 

objFSO = new ActiveXObject("Scripting.FileSystemObject");
WshShell = new ActiveXObject("WScript.Shell");

var sourceScriptComponent = objFSO.OpenTextFile(InstallationPath + "JScriptInclude.wsc").ReadAll(),
	sourceJScriptIncludeConfig = objFSO.OpenTextFile(InstallationPath + "jsModulesConfig.js").ReadAll(),
	sourceJScriptInclude = objFSO.OpenTextFile(InstallationPath + "JScriptInclude.js").ReadAll(),
	swap;
	
sourceScriptComponent = linemanPath(InstallationPath, sourceScriptComponent);
sourceJScriptIncludeConfig = linemanPath(InstallationPath, sourceJScriptIncludeConfig);
sourceJScriptInclude = linemanPath(InstallationPath, sourceJScriptInclude);

swap = objFSO.CreateTextFile(InstallationPath + "JScriptInclude.js", true);
swap.Write(sourceJScriptInclude);
swap.Close();
swap = null;
swap = objFSO.CreateTextFile(InstallationPath + "JScriptInclude.wsc", true);
swap.Write(sourceScriptComponent);
swap.Close();
swap = null;
swap = objFSO.CreateTextFile(InstallationPath + "jsModulesConfig.js", true)
swap.Write(sourceJScriptIncludeConfig);
swap.Close();
swap = null;

function linemanPath(replacement, source){
		var pattern =  /\$INSTALL_DIR\$/g;
		replacement = replacement.replace(/\\/g, "\\\\");
		return source.replace(pattern, replacement);
};

this.WshShell.Run("regsvr32 " + InstallationPath + "JScriptInclude.wsc",  0, false);


WScript.Echo(
  "[ENG]\nInstallation of JScriptInclude Gear completed successfully.\n\n[RUS]\n" 
+ "Установка JScriptInclude Gear произведена успешно.");

WScript.Echo(
  "[ENG]\nNow you will see our Documentation:\n\n[RUS]\n" 
+ "Сейчас Вам будет предложен просмотр Документации:\n\n"
+ InstallationPath + "Help.html");

this.WshShell.Run(InstallationPath + "Help.html",  0, false)
WScript.Sleep(6000)

WScript.Echo(
  "[ENG]\nAnd also run example of using:\n\n[RUS]\n" 
+ "А так же ознакомится с примером использования:\n\n"
+ InstallationPath + "example\\TestDir\\JScriptExample.jsl");

this.WshShell.Run('explorer.exe /select,"' + InstallationPath + 'example\\TestDir\\JScriptExample.js"',  1, false)
