jsImport.require("USM.js")

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
		
	})();