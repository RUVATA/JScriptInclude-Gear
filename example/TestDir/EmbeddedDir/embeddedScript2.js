	var TestClass = function(arg1, arg2){
		this.Pattern = "[ENG]\nHello, I'am " + arg1 + " Object \nA member of the " + arg2 + " class, declared in the emdeddedScript2.js" +
		"\n[RUS]\nПривет, я объект " + arg1 + "\nЯ экземпляр класса" + arg2 + " объявленного в скрипте emdeddedScript2.js"
		this.GetText = function(){
			return this.Pattern
		}
	};
	