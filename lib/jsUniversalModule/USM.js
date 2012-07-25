var USM = {};

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