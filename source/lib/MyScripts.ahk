;-------------------------------------------
; その他自作関数ライブラリ
; by akatubame
;-------------------------------------------
;
;-------------------------------------------

; 文字列を入力して変数に取得
_Inputbox(Title="", Prompt="", Default="", Timeout="", Repeat="", HIDE="", Width="", Height="", X="", Y="", Font=""){
	
	; ダイアログの自動調整（ウィンドウのXY座標、縦横サイズ）
	W      := ( Strlen(Default) > StrLen(Prompt) ) ? 150 + 10 * Strlen(Default) : 150 + 10 * StrLen(Prompt)
	H      := 140 + 20 * ( (StrLen(Prompt) / 30) - 1 )
	Width  := Width  ? Width  : (W > 400) ? 400 : (W < 200) ? 200 : W
	Height := Height ? Height : (StringGetPos(Prompt, "`n") != -1) ? H + 10 : H
	X      := X      ? X      : (A_ScreenWidth  - Width)  / 2
	Y      := Y      ? Y      : (A_ScreenHeight - Height) / 2
	
	InputBox, OutputVar, %Title%, %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%,, %Timeout%, %Default%
	
	; 入力をキャンセルした場合
	If (ErrorLevel)
		Exit
	
	; 空文字列を入力した場合、指定スイッチがあればリピート
	While (OutputVar == "" && Repeat!="") {
		InputBox, OutputVar, %Title%, %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%,, %Timeout%, %Default%
		If (ErrorLevel)
			Exit
	}
	return OutputVar
}
; 指定ファイルが存在すればリネームしてフラグを返す
_FileRename(fname, newName="", flag=0, PassExt=0){
	Attrib := FileExist(fname)
	If (Attrib)
	{
		dir     := _FileGetDir(fname)
		ext     := _FileGetExt(fname)
		default := (PassExt == 1) ? _FileGetNoExt(fname) : _FileGetName(fname)
		newName := (newName != "") ? newName : _Inputbox("ファイル名の変更", "変更後のファイル名を入力", default)
		
		If (Attrib == "D")
			FileMoveDir, %fname%, %dir%\%newName%, %flag%
		Else If (PassExt)
			FileMove, %fname%, %dir%\%newName%.%ext%, %flag%
		Else
			FileMove, %fname%, %dir%\%newName%, %flag%
		return ErrorLevel
	}
	Else
		return 0
}
; 指定ファイルが存在すれば指定先へコピーしてフラグを返す
_FileCopy(fname, fname2, flag=0){
	Attrib := FileExist(fname)
	If (Attrib)
	{
		If (Attrib == "D")
			FileCopyDir, %fname%, %fname2%, %flag%
		Else
			FileCopy, %fname%, %fname2%, %flag%
		return ErrorLevel
	}
	Else
		return 0
}
; 指定ワードを内容に持つファイルを新規作成 ( data=書き込む内容, fname=生成するファイル, enc=文字コード )
_FileNewAppend(data, fname, enc=""){
	FileDelete, %fname%
	FileAppend, %data%, %fname%, %enc%
}
; ダイアログからファイルを選択
_FileSelectFile(Path="", Prompt="", Filter=""){
	FileSelectFile, file, 3, %Path%, %Prompt%, %Filter%
	If (!file)
		Exit
	return file
}
; パスから拡張子を取得(.zipでなくzip)
_FileGetExt(path){
	SplitPath(path, name, dir, ext, noext, drive)
	return _RemoveSpace(ext)
}
; パスからファイル名(拡張子を除く)を取得
_FileGetNoExt(path){
	SplitPath(path, name, dir, ext, noext, drive)
	return noext
}
; パスからファイル名(拡張子付き)を取得
_FileGetName(path){
	SplitPath(path, name, dir, ext, noext, drive)
	return name
}
; パスから親ディレクトリを取得
_FileGetDir(path){
	SplitPath(path, name, dir, ext, noext, drive)
	return dir
}
; 指定フォーマットの文字列をオブジェクトに格納
_ObjFromStr(String, Rows="`n", Equal="=", Indent="`t"){
	obj := [], kn := []
	IndentLen := StrLen(Indent)
	Loop, parse, String, %Rows%
	{
		if A_LoopField is space
			continue
		Field := RTrim(A_LoopField, " `t`r")
		
		CurLevel := 1, k := "", v := ""
		While (SubStr(Field,1,IndentLen) == Indent) {
			StringTrimLeft, Field, Field, %IndentLen%
			CurLevel++
		}
		
		EqualPos := InStr(Field, Equal)
		if (EqualPos == 0)
			k := Field
		else
			k := SubStr(Field, 1, EqualPos-1), v := SubStr(Field, EqualPos+1)
		
		k := Trim(k, " `t`r"), v := Trim(v, " `t`r")
		kn[CurLevel] := k
		if !(EqualPos == 0)
		{
			if (CurLevel == 1)
			obj[kn.1] := v
			else if (CurLevel == 2)
			obj[kn.1][k] := v
			else if (CurLevel == 3)
			obj[kn.1][kn.2][k] := v
			else if (CurLevel == 4)
			obj[kn.1][kn.2][kn.3][k] := v
			else if (CurLevel == 5)
			obj[kn.1][kn.2][kn.3][kn.4][k] := v
			else if (CurLevel == 6)
			obj[kn.1][kn.2][kn.3][kn.4][kn.5][k] := v
			else if (CurLevel == 7)
			obj[kn.1][kn.2][kn.3][kn.4][kn.5][kn.6][k] := v
		}
		else
		{
			if (CurLevel == 1)
			obj.Insert(kn.1,Object())
			else if (CurLevel == 2)
			obj[kn.1].Insert(kn.2,Object())
			else if (CurLevel == 3)
			obj[kn.1][kn.2].Insert(kn.3,Object())
			else if (CurLevel == 4)
			obj[kn.1][kn.2][kn.3].Insert(kn.4,Object())
			else if (CurLevel == 5)
			obj[kn.1][kn.2][kn.3][kn.4].Insert(kn.5,Object())
			else if (CurLevel == 6)
			obj[kn.1][kn.2][kn.3][kn.4][kn.5].Insert(kn.6,Object())
		}
	}
	return obj
}
; 指定オブジェクトの内容を文字列に変換
_ObjToStr(Obj, Rows="`n", Equal=" = ", Indent="`t", Depth=7, CurIndent=""){
	For k,v in Obj
		ToReturn .= CurIndent . k . (IsObject(v) && depth>1 ? Rows . _ObjToStr(v, Rows, Equal, Indent, Depth-1, CurIndent . Indent) : Equal . v) . Rows
	return RTrim(ToReturn, Rows)
}
; 指定オブジェクトの内容をファイルに書き出す
_ObjToFile(Obj, FilePath, BackUp="", Rows="`n", Equal=" = ", Indent="`t", Depth=7, CurIndent=""){
	If ( BackUp != "" and IfExist(FilePath) ) {
		backup := _FileGetName(FilePath) ".bak"
		_FileRename(FilePath, backup, 1)
	}
	_FileNewAppend( _ObjToStr(Obj, Rows, Equal, Indent, Depth, CurIndent), FilePath, "UTF-8" )
	return ErrorLevel
}
; 指定ファイルの内容をオブジェクトに読み込む
_ObjFromFile(FilePath, Rows="`n", Equal="=", Indent="`t"){
	If ( !FileExist(FilePath) )
		return
	String := FileRead(FilePath)
	return _ObjFromStr(String, Rows, Equal, Indent)
}
; 指定文字列を、指定区切り記号ごとに分解後オブジェクト配列に格納
_StringSplit(InputVar, Delimiters, OmitChars=""){
	StringSplit, Array, InputVar, %Delimiters%, %OmitChars%
	Obj := []
	Loop {
		p := Array%A_Index%
		If (p == "")
			Break
		Obj.Insert(p)
	}
	return Obj
}
; 指定文字列(複数)を指定区切り記号で繋げて連結
_StringCombine(Delimiters, args*){
	str    := ""
	Length := StrLen(Delimiters)
	
	For key, value in args
		str .= value . Delimiters
	str := StringTrimRight(str, Length)
	return str
}
; 指定文字列から無駄なスペースや改行を取り除く
_RemoveSpace(str){
	str := StringReplace(str, """""", , "All")
	str := StringReplace(str, "　", A_Space, "All")
	str := StringReplace(str, "`r", , "All")
	str := StringReplace(str, "`n", , "All")
	str := RegExReplace(str, " +", " ")
	str := RegExReplace(str, "^\s", "")
	str := RegExReplace(str, "\s$", "")
	return str
}
; 指定文字列を指定リスト(置換前と後の文字列一覧)を参照して一括置換
_ListReplace(str, list="", ruleText=""){
	replaceRule := ruleText
	If ( replaceRule=="" ) {
		list        := IfExist(list) ? list : _FileSelectFile("replace", "置換リストを開く", "テキストドキュメント (*.txt)")
		replaceRule := FileRead(list)
	}
	
	Loop, parse, replaceRule, `n, `r
	{
		If (A_LoopField == "")
			continue
		
		StringSplit, rule, A_LoopField, %A_Tab%
		l := rule1, r := rule2
		
		str := StringReplace(str, l, r, "All")
	}
	return str
}
; 指定文字列を指定リスト(置換前と後の正規表現一覧)を参照して一括正規置換
_ListReplaceRegex(str, list="", ruleText=""){
	replaceRule := ruleText
	If ( replaceRule=="" ) {
		list        := IfExist(list) ? list : _FileSelectFile("regex", "正規表現リストを開く", "テキストドキュメント (*.txt)")
		replaceRule := FileRead(list)
	}
	
	Loop, parse, replaceRule, `n, `r
	{
		If (A_LoopField == "")
			continue
		
		StringSplit, rule, A_LoopField, %A_Tab%
		l := rule1, r := rule2
		
		l   := "imXS)" l
		r   := StringReplace(r, "\n", "`n", "All")
		r   := StringReplace(r, "\t", A_Tab, "All")
		str := RegExReplace(str, l, r)
	}
	return str
}
; 指定オブジェクトの最後尾に要素を挿入
_AddToObj(obj, target*){
	obj.Insert(target*)
}
; 指定オブジェクトの格納要素数を取得
_GetMaxIndex(obj){
	key := obj.MaxIndex() ? obj.MaxIndex() : 0y
	return key
}
; 指定オブジェクトをNativeCOMでラッピング、操作可能にする
_NativeCom(ByRef obj){
	If ( !IsObject(obj) )
		return
	
	ComObjError(false)
	If ( !ComObjType(obj,"iid") )
		obj := ComObjEnwrap(COM_Unwrap(obj))
	ComObjError(true)
}
; 指定ウィンドウのIDを取得
_WinGetId(twnd="A"){
	WinGet, id, ID, %twnd%
	return id
}
; 指定ウィンドウを常に最前面表示
_WinAlwaysTop(twnd="A"){
	WinSet, AlwaysOnTop, ON, %twnd%
}
; 指定ウィンドウがアクティブでなければアクティブにする
_WinActivate(twnd="A"){
	IfWinNotActive, %twnd%
		WinActivate, %twnd%
}
; 指定ウィンドウを最小化(タスクトレイに収納可)
_WinMinimize(twnd="A"){
	PostMessage, 0x112, 0xF020,,, %twnd%
}
; 指定ウィンドウを無理やりトレイアイコンに最小化
_WinMinimizeTray(twnd="A"){
	_WinActivate(twnd)
	Send, !#w
}
; 指定ウィンドウを最大化
_WinMaximize(twnd="A"){
	WinMaximize, %twnd%
}
; 指定ウィンドウの最大化・最小化を解除
_WinRestore(twnd="A"){
	WinRestore, %twnd%
}
; 指定コマンドを実行
_Run(runapp, option="", runcmd=""){
	Run, %runapp% %option%,, %runcmd%
}
; モニタの電源をOFFにする
_MonitorOff(){
	SendMessage, 0x112, 0xF170, 2,, ahk_id 0xFFFF
}
; システムの再起動を実行
_Reboot(){
	Shutdown, 2
}
; システムをスリープ・休止状態へ移行
_Hybernate(){
	DllCall("PowrProf\SetSuspendState", "int", 0, "int", 0, "int", 0)
}
; Google検索バー
_SearchBox(){
	Input := _Inputbox("Search for Google",,, 60, "Repeat",, 230, 100)
	Run, C:\Program Files\Internet Explorer\iexplore.exe "https://www.google.co.jp/search?hl=ja&q=%Input%"
}
; 指定のAHK関数を実行
_ExecFunc(function, args*){
	
	; 関数が不存在ならエラー
	If ( !IsFunc(function) )
		throw Exception("指定した関数は存在しません : _ExecFunc(""" function """).")
	
	; クラス関数なら引数の先頭にダミーを追加
	If ( IfInString(function, ".") )
		args.Insert(1, "")
	
	; 引数の数量チェック
	receivedArgs := args.maxIndex() ? args.maxIndex() : 0
	needArgs     := IsFunc(function) - 1
	If ( receivedArgs < needArgs )
		throw Exception("引数の指定が不足しています : _ExecFunc(""" function """).")
	
	; 関数の実行結果を返り値に
	return Func(function).(args*)
}
