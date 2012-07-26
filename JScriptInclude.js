var jsImport = { 
 requiredSource: "",
 WriteCache: false,
 modulesByDefault: {},
 uploadedLibs: []
};

jsImport.defineBaseModule= function(moduleName, modulePath){	
	jsImport.modulesByDefault[moduleName] = {};
	jsImport.modulesByDefault[moduleName]["path"] = modulePath;
};
 
jsImport.getConfig = function(ConfFile){
	jsImport.fso = jsImport.fso || new ActiveXObject("Scripting.FileSystemObject");
	try{
		var SourseConf = jsImport.fso.OpenTextFile(ConfFile).ReadAll();
		eval(SourseConf);
	}catch(err){
		WScript.Echo("Error(jsImport.getConfig): Unable to load the configuration file\n"+err.description)
	}
}

jsImport.initialization = function(){
 var temp = jsImport.requiredSource;
 jsImport.requiredSource = null;
 if(jsImport.WriteCache == true){jsImport.Cache.Close()};
 return temp;
};


jsImport.require = function (TargetScript, ParentDir){
	jsImport.fso = jsImport.fso || new ActiveXObject("Scripting.FileSystemObject");
	if(jsImport.WriteCache == true){jsImport.Cache = jsImport.Cache || jsImport.fso.OpenTextFile("CacheFile.js", 2,true)}	
	var ParentDir = ParentDir || jsImport.spoofedParentDir || WScript.ScriptFullName.replace(/[^\\]+$/, '');  
	var TargetScriptName = TargetScript.substring(TargetScript.lastIndexOf("\\")+1,TargetScript.length - 3);
	TargetScript = PathValidation(TargetScript, ParentDir);		 
	
		jsImport.require.AlreadyLoaded = false;
		for (i in jsImport.uploadedLibs){
			if(jsImport.uploadedLibs[i] == TargetScriptName){jsImport.require.AlreadyLoaded = true};
		};		   
	   
   if(typeof jsImport.require.AlreadyLoaded == "undefined" || jsImport.require.AlreadyLoaded !== true){			  		   	   
	  try{	
			var SourceBuffer = jsImport.fso.OpenTextFile(TargetScript).ReadAll();				
			var requireArray = requireParser();				
			jsImport.requiredSource =  jsImport.requiredSource + SourceBuffer;	

			if(jsImport.WriteCache == true){jsImport.Cache.Write(SourceBuffer)};			
			
			if(requireArray.length != 0){					
				for(i in requireArray){
					jsImport.require(requireArray[i].replace(/\\+/g, "\\"), TargetScript.replace(/[^\\]+$/,""));
				};
			};
				
			jsImport.uploadedLibs.push(TargetScriptName); AlreadyLoaded = false;
			
		}catch(err){
			WScript.Echo("Error(jsImport.require): Target " + TargetScript + " can't be load\n" + err.description);
		};
	 
	function requireParser(){
		var pattern =  /jsImport.require\(["'](.*?)["']\)/g;
		var result;
		var requireArray = [];					
			while ((result = pattern.exec(SourceBuffer)) != null) {
				requireArray.push(result[1]);
				SourceBuffer = SourceBuffer.replace(result[0],"");
				pattern.lastIndex = 0;
			};
		return requireArray;
	};	
	
	function PathValidation(TargetScript, ParentDir){	
		if(!(TargetScript.charAt(1) == ":" || TargetScript.charAt(1) == "\\")){
			var dot = TargetScript.charAt(1);
			if(dot == "."){
				var OutsidePath = ParentDir, counter= 0;                    		 
				while(dot == "."){
					counter++;
					dot = TargetScript.charAt(counter);
					OutsidePath = StepOutside(OutsidePath);
				};		 
				TargetScript = (OutsidePath + TargetScript.substring(counter)).replace(/\\+/g, "\\") ;
				dot = "undefined"; counter = "undefined"; OutsidePath = "undefined";
			} else {
				if(TargetScript.substring(0,1) == "."){TargetScript = TargetScript.substring(1)};
				if(TargetScript.substring(0,2) !== "\\\\"){TargetScript = "\\\\" + TargetScript};
				TargetScript = (ParentDir + TargetScript.substring(1)).replace(/\\+/g, "\\");
			};
		};
		
		if( ! jsImport.fso.FileExists(TargetScript)){
			if(TargetScript.substring(TargetScript.length - 3, TargetScript.length) == ".js"){TargetScript = TargetScript.substring(0, TargetScript.length - 3)}
			var TagertModule = TargetScript.substring(TargetScript.lastIndexOf("\\")+1,TargetScript.length)					
			for(module in  jsImport.modulesByDefault){
				if (module == TagertModule){
					TargetScript = jsImport.modulesByDefault[module].path;
					break;
				};
			};
		};	
		return TargetScript;
		
		function StepOutside(outside){
			return outside.substring(0, outside.lastIndexOf("\\"));
		};
	};
};
};
jsImport.getConfig("B:\\JScriptInclude\\JS\\jsModulesConfig.js")