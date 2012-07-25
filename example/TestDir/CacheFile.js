

function jlib7zaArchivator(){
		this.WshShell = new ActiveXObject("WScript.Shell");
		this.setDisplayModeChoice = 1;	
		this.List = [];	
		this.path7za = (function(){
		this.setDisplayModeChoice = 1;
		
			var byDefault = "7za.exe", path;
			
			for(module in  jsImport.modulesByDefault){
				if (module == "7zip"){
					path = jsImport.modulesByDefault[module].path;
					break;
				};
			};
						
			if(typeof path != "undefined"){path = path.substring(0,path.lastIndexOf("\\")+1) + "7za.exe"}else{path = "7za.exe"};
			
			return path;
			
		})();	
};
	
	(function(){	
		jlib7zaArchivator.prototype.setDisplayMode = function(mode){
			if(typeof mode != "undefined"){mode = mode.toLowerCase()}else{var mode = "visible"}; 		
			switch (mode) {
				case "visible":
					this.setDisplayModeChoice = 1;
					break;
				case "hidden":
					this.setDisplayModeChoice = 0;
					break;
				case "minimazed":
					this.setDisplayModeChoice = 2;
					break;
				case "maximazed":
					this.setDisplayModeChoice = 3;
					break;
			};
		};
		
		jlib7zaArchivator.prototype.toList = function(newFile){
			try{
				var toPush = USM.PathValidation(newFile)
				this.List.push(toPush);
				return -1				
			}catch(err){
				return 0
			};
		};
		
		jlib7zaArchivator.prototype.clearList = function(){
				this.List = null;
				this.List =new Array();
		};
		
		jlib7zaArchivator.prototype.SingleShell7za = function(TargetFile, MacroString){
		var TargetArchName = USM.MacrosBase.MacrosParser(MacroString) + ".zip ";
		var ShellComand = USM.PathValidation(this.path7za) + " a -tzip " + TargetArchName + TargetFile + " -mx=7";
			try{
				this.WshShell.Run(ShellComand,  this.setDisplayModeChoice, true);
			}catch(theError){
				WScript.Echo('Error from ActiveXObject("WScript.Shell")\nShellComand: '  + ShellComand + USM.getErrorToString(theError) + '\n' + TargetFile + " can not be successfully added to archive")
			};
			ShellComand = null;
			return TargetArchName;
		};
		
		jlib7zaArchivator.prototype.Archivate = function(TargetFile, MacroString){
		var	TargetArchName;
			if(TargetFile != 'list'){
				return this.SingleShell7za(USM.PathValidation(TargetFile), MacroString);				
			}else{				
				if(this.List.length > 0){TargetFile = this.List}else{WScript.Echo('error(jlib7zipArchivator):The list is empty\nUse the "jlib7zipArchivator.toList" method, or complete property "jlib7zipArchivator.List" manually')};
				for(var OneOfTargerts=0; OneOfTargerts < TargetFile.length; OneOfTargerts++){									
					TargetArchName = this.SingleShell7za(TargetFile[OneOfTargerts], MacroString);
				};
			};
			return TargetArchName;
			this.clearList();			
		};
		
	})();var USM = {};

USM.Date = new Date();
	
USM.MacrosBase = {
	DD: function(){var temp = (USM.Date.getDate()).toString();if(temp.length < 2){temp = "0" + temp};return temp;},
	MM: function(){var temp = (USM.Date.getMonth() + 1).toString();if(temp.length < 2){temp = "0" + temp};return temp;},
	YYYY: function(){return (USM.Date.getFullYear()).toString()},
	YY: function(){return ((USM.MacrosBase.YYYY()).toString()).substring(2)}
};

USM.MacrosBase.MacrosParser = function(MacroString){
		MacroString = MacroString.replace(/DD/g, USM.MacrosBase.DD());
		MacroString = MacroString.replace(/MM/g, USM.MacrosBase.MM());
		MacroString = MacroString.replace(/YYYY/g, USM.MacrosBase.YYYY());
		MacroString = MacroString.replace(/YY/g, USM.MacrosBase.YY());
		return MacroString;
};

USM.getErrorToString = function(err){
		var errString = "";
		   errString = errString +  "\n\tError:    type          |   " + err;
			errString = errString + "\n\t             number      |   " + err.number;
			errString = errString + "\n\t             description |   " + err.description;
		return errString;
};

USM.getParamsToString = function(params){
			var paramString = "";
				for(var param in params){
					paramString = paramString + "\n\tParam" + param + ": " + params[param]
				};
			return paramString;	
};

USM.PathValidation = function(TargetFile){	
	var ParentDir = WScript.ScriptFullName.replace(/[^\\]+$/, '');
	if(!(TargetFile.charAt(1) !== ":" || TargetFile.charAt(1) !== "\\") || TargetFile.charAt(TargetFile.length - 1) == "\\"){		
		var dot = TargetFile.charAt(1);
		if(dot == "."){
			var OutsidePath = ParentDir, counter= 0;                    		 
				while(dot == "."){
					counter++;
					dot = TargetFile.charAt(counter);
					OutsidePath = StepOutside(OutsidePath);
				};		 
			TargetFile = OutsidePath + TargetFile.substring(counter);
			dot = "undefined"; counter = "undefined"; OutsidePath = "undefined";
		} else {
			if(TargetFile.substring(0,1) == "."){TargetFile = TargetFile.substring(1)};
			if(TargetFile.substring(0,1) !== "\\"){TargetFile = "\\" + TargetFile};
			TargetFile = ParentDir + TargetFile.substring(1);
		};
	};
	
	function StepOutside(outside){
	   return outside.substring(0, outside.lastIndexOf("\\"));
	}; 
		
   return TargetFile.replace(/\\+/g, "\\\\")
};

var sureFSO = function(){
	this.ActiveX = new ActiveXObject("Scripting.FileSystemObject");
	this.logout = [];
};	

(function(){	
	sureFSO.prototype.logoutToString = function(){
	var logList = "***LogOut***";
		for(var obj in this.logout){
			logList = logList + "\n    " + (parseInt(obj) +1) + ")"+ "\n\tMethod: " + this.logout[obj].name + USM.getParamsToString(this.logout[obj].params) + USM.getErrorToString(this.logout[obj].err);
		};		
	return logList;
	};
		
	sureFSO.prototype.item = function (method, params, err) {
		this.name = method;
		this.params = params;
		this.err = err;
	}

	sureFSO.prototype.ErrorMsg = function(method, err, params){
		WScript.Echo('this.onErrorsMode is "stop" \nMethod: ' + method + USM.getParamsToString(params) + USM.getErrorToString(err));
	};

	sureFSO.prototype.setOnErrorsMode = function(mode){
		if (typeof this.setOnErrorsMode.Choice == "undefined"){this.setOnErrorsMode.Choice = 0};
			switch (mode) {
				case "stop":
					this.setOnErrorsMode.Choice = 0;
					break;
				case "ignore":
					this.setOnErrorsMode.Choice = 1;
					break;
			};
	};

	sureFSO.prototype.CopyFile = function(source, destanation, overwrite){
		var overwrite = overwrite || true;
		if(typeof source  != "undefined"){source = USM.PathValidation(source)}else{var source = "notSpecified"}; 
		if(typeof destanation  != "undefined"){destanation = USM.PathValidation(destanation)}else{var destanation = "notSpecified"}; 
		
		try{
			this.ActiveX.CopyFile(source, destanation, overwrite);
			return -1;
		}catch(err){
			if(this.setOnErrorsMode.Choice != 1){
				this.ErrorMsg("CopyFile", err, [source, destanation, overwrite])
				WScript.Quit();
			}else{
				this.logout.push(new this.item("CopyFile", [source, destanation, overwrite], err));
				return 0;
			};
		};
	};

	sureFSO.prototype.CopyFolder = function(source, destanation, overwrite){
		var overwrite = overwrite || true;
		if(typeof source  != "undefined"){source = USM.PathValidation(source)}else{var source = "notSpecified"}; 
		if(typeof destanation  != "undefined"){destanation = USM.PathValidation(destanation)}else{var destanation = "notSpecified"}; 
		
		try{
			this.ActiveX.CopyFolder(source, destanation, overwrite);
			return -1;
		}catch(err){
			if(this.setOnErrorsMode.Choice != 1){
				this.ErrorMsg("CopyFolder", err, [source, destanation, overwrite]);
				WScript.Quit();
			}else{
				this.logout.push(new this.item("CopyFolder", [source, destanation, overwrite], err));
				return 0;
			};
		};
	};

	sureFSO.prototype.DeleteFile = function(filespec, force){
		var force = force || true;
		if(typeof filespec  != "undefined"){filespec = USM.PathValidation(filespec)}else{var filespec = "notSpecified"}; 
		try{
			this.ActiveX.DeleteFile(filespec, force);
			return -1;
		}catch(err){
			if(this.setOnErrorsMode.Choice !== 1 ) {
				this.ErrorMsg("DeleteFile", err, [filespec, force]);
				WScript.Quit();
			}else{
				this.logout.push(new this.item("DeleteFile", [filespec, force], err));
				return 0;
			};
		};
	};

	sureFSO.prototype.DeleteFolder = function(filespec, force){
		var force = force || true;
		if(typeof filespec  != "undefined"){filespec = USM.PathValidation(filespec)}else{var filespec = "notSpecified"}; 
		try{
			this.ActiveX.DeleteFolder(filespec, force);
			return -1;
		}catch(err){
			if(this.setOnErrorsMode.Choice != 1){
				this.ErrorMsg("DeleteFolder", err, [filespec, force]);
				WScript.Quit();
			}else{
				this.logout.push(new this.item("DeleteFolder", [filespec, force], err));
				return 0;
			};
		};
	};
	
	sureFSO.prototype.CreateFolder = function(foldername, force){
		var force = force || true;
		if(typeof filespec  != "undefined"){filespec = USM.PathValidation(filespec)}else{var filespec = "notSpecified"}; 
		if(!this.ActiveX.FolderExists(foldername)){
			try{
				this.ActiveX.CreateFolder(foldername);
			}catch(err){			
				if(this.setOnErrorsMode.Choice != 1){
					this.ErrorMsg("CreateFolder", err, [foldername, force]);
					WScript.Quit();
				}else{
					this.logout.push(new this.item("CreateFolder", [foldername, force], err));
					return 0
				};
			};
		 return -1
		}else{
		 return 1
		};
	};

	sureFSO.prototype.FolderExists = function(foldername){
		var foldername = foldername || "notSpecified";
		var prereturn;
		
		try{
			prereturn = this.ActiveX.FolderExists(foldername);
			return prereturn;
		}catch(err){
			if(this.setOnErrorsMode.Choice != 1){
				this.ErrorMsg("FolderExists", err, [foldername]);
				WScript.Quit();
			}else{
				this.logout.push(new this.item("DeleteFolder", [foldername], err));
				return 0;
			};
		};
	}
	
	sureFSO.prototype.FileExists = function(filename){
		var filename  = filename || "notSpecified";
		var prereturn;
		
		try{
			prereturn = this.ActiveX.FileExists(filename);
			return prereturn;
		}catch(err){
			if(this.setOnErrorsMode.Choice != 1){
				this.ErrorMsg("FileExists", err, [filename]);
				WScript.Quit();
			}else{
				this.logout.push(new this.item("FileExists", [filename], err));
				return 1;
			};
		};
	};				
})();
var embeddedScriptVar = '[ENG]\nThis is "embeddedScriptVar" declared in embeddedScrpt.js\n[RUS]\nЁто значение переменной "embeddedScriptVar" об€вленной в embeddedScrpt.js';
 //ќбъ€вл€ем зависимость от скрипта по относительному пути.
	var TestClass = function(arg1, arg2){
		this.Pattern = "[ENG]\nHello, I'am " + arg1 + " Object \nA member of the " + arg2 + " class, declared in the emdeddedScript2.js" +
		"\n[RUS]\nѕривет, € объект " + arg1 + "\nя экземпл€р класса" + arg2 + " объ€вленного в скрипте emdeddedScript2.js"
		this.GetText = function(){
			return this.Pattern
		}
	};
	