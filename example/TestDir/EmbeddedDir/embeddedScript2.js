	var TestClass = function(arg1, arg2){
		this.Pattern = "[ENG]\nHello, I'am " + arg1 + " Object \nA member of the " + arg2 + " class, declared in the emdeddedScript2.js" +
		"\n[RUS]\n������, � ������ " + arg1 + "\n� ��������� ������" + arg2 + " ������������ � ������� emdeddedScript2.js"
		this.GetText = function(){
			return this.Pattern
		}
	};
	