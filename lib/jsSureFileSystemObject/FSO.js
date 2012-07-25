jsImport.require("USM.js")

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
