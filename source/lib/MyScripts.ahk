; 指定文字列から無駄なスペースや改行を取り除く
_RemoveSpace(text){
	text := StringReplace(text, """""", , "All")
	text := StringReplace(text, "　", A_Space, "All")
	text := StringReplace(text, "`r", , "All")
	text := StringReplace(text, "`n", , "All")
	text := RegExReplace(text, " +", " ")
	text := RegExReplace(text, "^\s", "")
	text := RegExReplace(text, "\s$", "")
	return text
}
; 指定文字列をダブルクォーテーション[""]で囲む
_WQ(str){
	;コマンドラインの場合は空白でも囲むべき
	;If (str="")
	;	return
	
	StringGetPos, start, str, ", L
	If (ErrorLevel = 0) {
		StringLen, Length, str
		StringGetPos, end, str, ", R
		If (start = 0 && end = Length-1)
			return str
	}
	str = "%str%"
	return str
}
; 指定文字列のダブルクォーテーション[""]の囲みを削除
_DeWQ(str){
	StringGetPos, start, str, ", L
	If (ErrorLevel = 0) {
		StringLen, Length, str
		StringGetPos, end, str, ", R
		If (start = 0 && end = Length-1) {
			str := StringTrimLeft(str, 1)
			str := StringTrimRight(str, 1)
		}
	}
	return str
}
; 指定文字列を、指定区切り記号ごとに分解後オブジェクト配列に格納
_StringSplit(InputVar, Delimiters, OmitChars=""){
	StringSplit, Array, InputVar, %Delimiters%, %OmitChars%
	Obj := Object()
	Loop {
		p := Array%A_Index%
		If (p = "")
			Break
		Obj.Insert(p)
	}
	return Obj
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
; パスから現在ドライブを取得
_FileGetDrive(path){
	SplitPath(path, name, dir, ext, noext, drive)
	return _RemoveSpace(drive)
}
; 指定ファイルが存在すればリネームしてフラグを返す
_FileRename(fname, newName="", flag=0, PassExt=0){
	Attrib := FileExist(fname)
	If (Attrib)
	{
		dir     := _FileGetDir(fname)
		ext     := _FileGetExt(fname)
		default := (PassExt = 1) ? _FileGetNoExt(fname) : _FileGetName(fname)
		newName := (newName != "") ? newName : _Inputbox("ファイル名の変更", "変更後のファイル名を入力", default)
		
		If (Attrib = "D")
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
; 指定ファイルが存在すれば指定先へコピーしてフラグを返す
_FileCopy(fname, fname2, flag=0){
	Attrib := FileExist(fname)
	If (Attrib)
	{
		If (Attrib = "D")
			FileCopyDir, %fname%, %fname2%, %flag%
		Else
			FileCopy, %fname%, %fname2%, %flag%
		return ErrorLevel
	}
	Else
		return 0
}

; 指定オブジェクトの最後尾に要素を挿入
_AddToObj(obj, target*){
	obj.Insert(target*)
}
; 指定フォーマットの文字列をオブジェクトに格納
_ObjFromStr(String, Rows="`n", Equal="=", Indent="`t"){
	obj := Object(), kn := Object()
	IndentLen := StrLen(Indent)
	Loop, parse, String, %Rows%
	{
		if A_LoopField is space
			continue
		Field := RTrim(A_LoopField, " `t`r")
		
		CurLevel := 1, k := "", v := ""
		While (SubStr(Field,1,IndentLen) = Indent) {
			StringTrimLeft, Field, Field, %IndentLen%
			CurLevel++
		}
		
		EqualPos := InStr(Field, Equal)
		if (EqualPos = 0)
			k := Field
		else
			k := SubStr(Field, 1, EqualPos-1), v := SubStr(Field, EqualPos+1)
		
		k := Trim(k, " `t`r"), v := Trim(v, " `t`r")
		kn[CurLevel] := k
		if !(EqualPos = 0)
		{
			if (CurLevel = 1)
			obj[kn.1] := v
			else if (CurLevel = 2)
			obj[kn.1][k] := v
			else if (CurLevel = 3)
			obj[kn.1][kn.2][k] := v
			else if (CurLevel = 4)
			obj[kn.1][kn.2][kn.3][k] := v
			else if (CurLevel = 5)
			obj[kn.1][kn.2][kn.3][kn.4][k] := v
			else if (CurLevel = 6)
			obj[kn.1][kn.2][kn.3][kn.4][kn.5][k] := v
			else if (CurLevel = 7)
			obj[kn.1][kn.2][kn.3][kn.4][kn.5][kn.6][k] := v
		}
		else
		{
			if (CurLevel = 1)
			obj.Insert(kn.1,Object())
			else if (CurLevel = 2)
			obj[kn.1].Insert(kn.2,Object())
			else if (CurLevel = 3)
			obj[kn.1][kn.2].Insert(kn.3,Object())
			else if (CurLevel = 4)
			obj[kn.1][kn.2][kn.3].Insert(kn.4,Object())
			else if (CurLevel = 5)
			obj[kn.1][kn.2][kn.3][kn.4].Insert(kn.5,Object())
			else if (CurLevel = 6)
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
; 指定オブジェクトの格納要素数を取得
_GetMaxIndex(obj){
	key := obj.MaxIndex() ? obj.MaxIndex() : 0y
	return key
}

; 複数の文字列をハイフン"-"で連結してメッセージボックスに表示
_MsgBox(params*){
	Msg := ""
	For key, value in params
		Msg .= value . " - "
	Msg := StringTrimRight(Msg, 3)
	MsgBox, % Msg
}
; 指定ウィンドウを閉じる
_WinClose(twnd="A"){
	WinClose, %twnd%
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
; 指定ウィンドウを常に最前面表示
_WinAlwaysTop(twnd="A"){
	WinSet, AlwaysOnTop, ON, %twnd%
}
; 指定ウィンドウのIDを取得
_WinGetId(twnd="A"){
	WinGet, id, ID, %twnd%
	return id
}

; 指定コマンドを実行
_Run(runapp, option="", runcmd=""){
	Run, %runapp% %option%,, %runcmd%
}
; 指定コマンドを作業フォルダを指定して実行
_RunIn(runapp, option="", runcmd="", dir=""){
	If (dir="")
		dir := _FileGetDir(runapp)
	Run, %runapp% %option%, %dir%, %runcmd%
}
; 指定ウィンドウが存在しなければ指定コマンドを実行、存在すればアクティブに
_RunOrActive(twnd, runapp, option="", runcmd=""){
	IfWinNotExist, %twnd%
	{
		_RunIn(runapp, option, runcmd)
		return 1
	}
	else
	{
		_WinActivate(twnd)
		return 0
	}
}
; IMEのオンオフ状態を取得 (戻り値 1=on 0=off)
_IME_GET(WinTitle="A"){
	VarSetCapacity(stGTI, 48, 0)
	NumPut(48, stGTI,  0, "UInt")   ;	DWORD   cbSize;
	hwndFocus := DllCall("GetGUIThreadInfo", Uint,0, Uint,&stGTI)
				 ? NumGet(stGTI,12,"UInt") : WinExist(WinTitle)

	return DllCall("SendMessage"
		, UInt, DllCall("imm32\ImmGetDefaultIMEWnd", Uint,hwndFocus)
		, UInt, 0x0283  ;Message : WM_IME_CONTROL
		,  Int, 0x0005  ;wParam  : IMC_GETOPENSTATUS
		,  Int, 0)      ;lParam  : 0
}
; IMEのオンオフ切替 (SetSTsの値 1=on 0=off)
_IME_SET(SetSts, WinTitle="A"){
	VarSetCapacity(stGTI, 48, 0)
	NumPut(48, stGTI,  0, "UInt")   ;	DWORD   cbSize;
	hwndFocus := DllCall("GetGUIThreadInfo", Uint,0, Uint,&stGTI)
					? NumGet(stGTI,12,"UInt") : WinExist(WinTitle)
	return DllCall("SendMessage"
		, UInt, DllCall("imm32\ImmGetDefaultIMEWnd", Uint,hwndFocus)
		, UInt, 0x0283  ;Message : WM_IME_CONTROL
		,  Int, 0x006   ;wParam  : IMC_SETOPENSTATUS
		,  Int, SetSts) ;lParam  : 0 or 1
}
; IMEの入力状態を取得 (戻り値 1=入力中 2=変換中 0=入力なし)
_IME_GetConverting(WinTitle="A",ConvCls="",CandCls=""){

	;IME毎の 入力窓/候補窓Class一覧 ("|" 区切りで適当に足してけばOK)
	ConvCls .= (ConvCls ? "|" : "")                 ;--- 入力窓 ---
		.  "ATOK\d+CompStr\d*"                  ; ATOK系
		.  "|imejpstcnv\d+"                     ; MS-IME系
		.  "|WXGIMEConv"                        ; WXG
		.  "|SKKIME\d+\.*\d+UCompStr"           ; SKKIME Unicode
		.  "|MSCTFIME Composition"              ; Google日本語入力

	CandCls .= (CandCls ? "|" : "")                 ;--- 候補窓 ---
		.  "ATOK\d+Cand"                        ; ATOK系
		.  "|imejpstCandList\d+|imejpstcand\d+" ; MS-IME 2002(8.1)XP付属
		.  "|mscandui\d+\.candidate"            ; MS Office IME-2007
		.  "|WXGIMECand"                        ; WXG
		.  "|SKKIME\d+\.*\d+UCand"              ; SKKIME Unicode
   CandGCls := "GoogleJapaneseInputCandidateWindow" ;Google日本語入力

	VarSetCapacity(stGTI, 48, 0)
	NumPut(48, stGTI,  0, "UInt")   ;	DWORD   cbSize;
	hwndFocus := DllCall("GetGUIThreadInfo", Uint,0, Uint,&stGTI)
				? NumGet(stGTI,12,"UInt") : WinExist(WinTitle)

	WinGet, pid, PID,% "ahk_id " hwndFocus
	tmm:=A_TitleMatchMode
	SetTitleMatchMode, RegEx
	ret := WinExist("ahk_class " . CandCls . " ahk_pid " pid) ? 2
		:  WinExist("ahk_class " . CandGCls                 ) ? 2
		:  WinExist("ahk_class " . ConvCls . " ahk_pid " pid) ? 1
		:  0
	SetTitleMatchMode, %tmm%
	return ret
}
; モニタの電源をOFFにする
_MonitorOff(){
	SendMessage, 0x112, 0xF170, 2,, ahk_id 0xFFFF
}
; システムのシャットダウンを実行
_ShutDown(){
	Shutdown, 1
}
; システムの再起動を実行
_Reboot(){
	Shutdown, 2
}
; システムをスリープ・休止状態へ移行
_Hybernate(){
	DllCall("PowrProf\SetSuspendState", "int", 0, "int", 0, "int", 0)
}

; Google検索実行
_Google(word=""){
	If (word)
		_UE( "google", _WQ(word) )
}
; 対象ワード(複数可)でぐぐる
_SearchText(target){
	Loop, parse, target, `n, `r
		If (A_LoopField != "")
			_Google(A_LoopField)
}
; 指定パス(複数可)を適切なアプリで開く
_OpenPath(target){
	Loop, parse, target, `n, `r
	{
		If (A_LoopField = "")
			continue
		
		Path  := _RemoveSpace( _DeWQ(A_LoopField) )
		drive := _FileGetDrive(Path)
		pos   := InStr(drive, "ttp")
		If pos between 1 and 2
		{
			URLs := % (pos == 1) ? ("h" . Path) : Path
			_Run("..\..\cpt\FirefoxPortable\FirefoxPortable.exe", URLs)
		}
		Else IfExist, %Path%
		{
			Path := _RelToAbs(Path)
			ext  := _FileGetExt(Path)
			If (RegExMatch(ext, "^(txt)$"))
				_Run("..\X-Finder\XF.exe", _WQ(Path) " ..")
			Else
				_Run("..\X-Finder\XF.exe", _WQ(Path) )
		}
		Else
		{
			pos2 := InStr(Path, "HKEY_")
			If (pos2 = 1)
				_Run("..\GekiOreRegEdit\GekiOreRegEdit.exe", Path)
			Else If ( RegExMatch(Path, "^(\d|[01]?\d\d|2[0-4]\d|25[0-5])\.(\d|[01]?\d\d|2[0-4]\d|25[0-5])\.(\d|[01]?\d\d|2[0-4]\d|25[0-5])\.(\d|[01]?\d\d|2[0-4]\d|25[0-5])$") )
			{
				_Run("..\..\cpt\FirefoxPortable\FirefoxPortable.exe", "http://www.iphiroba.jp/index.php")
				Sleep, 4000
				MouseClick, Left, 750, 397, 1, 0
				Sleep 100
				SendInput, %Path%
				Sleep 100
				Send, {Enter}
			}
		}
	}
}
; 指定サイト＆指定ワードでWEB検索
_UE(SearchE, target, encoded=""){
	url := _UE_Get(SearchE, target, encoded)
	_OpenPath(url)
}
; 指定サイト＆指定ワードでWEB検索URLを生成
_UE_Get(SearchE, target, encoded=""){
	obj := A_Init_Object["UE"][SearchE]
	str := IsObject(encoded) ? encoded[obj["encode"]] : _URLEncode(target, obj["encode"])
	
	url := (obj["urlsuf"] != "") ? obj["urlpre"] . str . obj["urlsuf"] : obj["urlpre"] . str
	return url
}
; 指定文字列をURLエンコード (enc=文字コード指定[1=Shift_JIS 2=EUC-JP 3=UTF-8 4=JIS 空=すべて])
_URLEncode(Str, Enc=""){
	option := _OptionCombine(str, enc)
	;stdout := _PHP("PHP\urlencode.php", "STDOUT", option )
	stdout := _RunStdOut("PHP\urlencode.exe", option )
	return stdout
}
; 指定コマンドラインオプション(複数)をダブルクォーテーション[""]で囲み連結
_OptionCombine(params*){
	option := ""
	For key, value in params
		option .= _WQ(value) . " "
	option := StringTrimRight(option, 1)
	return option
}
; 指定コマンドを作業フォルダを指定して実行、標準出力(STDOUT)を取得
_RunStdOut(runapp, option="", dir=""){
	runapp := _RelToAbs(runapp)
	If (dir="")
		dir := _FileGetDir(runapp)
	WDir := A_WorkingDir
	SetWorkingDir, %dir%
	
	stdout := _WSHExec(runapp . " " . option)
	
	SetWorkingDir, %WDir%
	return, stdout
}
; 指定WScriptShellをExecで実行
_WSHExec(command){
	exec := ComObjCreate("WScript.Shell").Exec(command)
	While !exec.Status
		Sleep 100
	strLine := exec.StdOut.ReadAll()
	return strLine
}
; Google検索バー
_SearchBox(){
	_AHK_SA("_SearchBox_Main")
}
; SearchBox実行
_SearchBox_Main(){
	Input := _Inputbox("Search for Google",,, 60, "Repeat",, 230, 100)
	_Google(Input)
}
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
	While (OutputVar = "" && Repeat!="") {
		InputBox, OutputVar, %Title%, %Prompt%, %HIDE%, %Width%, %Height%, %X%, %Y%,, %Timeout%, %Default%
		If (ErrorLevel)
			Exit
	}
	return OutputVar
}
; 指定時間経過後にアラートを表示
_StopAlert(time, unit, sound=""){
	If (unit = "H")
		u := "時間", sleeptime := time * 1000 * 60 * 60
	Else If (unit = "M")
		u := "分",   sleeptime := time * 1000 * 60
	Else
		u := "秒",   sleeptime := time * 1000
	
	_Notify("=== STOP ALERT ===", time u "後にアラート", 3)
	Sleep, %sleeptime%
	_Notify("=== STOP ALERT ===", time u "が経過しました",,,, sound)
}
; 指定文字列で付箋を作り画面に貼る
_Notify(Title="!!!",Message="",Duration=3,Options="",Image="",Sound=""){
	_AHK("Notify",, Title, Message, Duration, Options, Image, Sound)
}
; AHKスクリプト実行
_AHK(fname, wait="", params*){
	IfNotInString, fname, \
		script := "util\" fname ".ahk"
	Else
		script := _RelToAbs(fname)
	
	option := _OptionCombine(params*)
	If (wait)
		_RunWait("AutoHotKey.exe", script " " option)
	Else
		_Run("AutoHotKey.exe", script " " option)
}
; スタンドアロンでAHK関数を実行
_AHK_SA(function, params*){
	_AHK("AHK_StandAlone",, function, params*)
}
; 指定ディレクトリを基準とした相対パスを絶対パスに変換
_RelToAbs_From(root, dir, s = "\"){
	; 既に絶対パスなら処理せず返す
	If ( _FileGetDrive(dir) )
		return dir

	pr := SubStr(root, 1, len := InStr(root, s, "", InStr(root, s . s) + 2) - 1)
		, root := SubStr(root, len + 1)
	If InStr(root, s, "", 0) = StrLen(root)
		root := StringTrimRight(root, 1)
	If InStr(dir, s, "", 0) = StrLen(dir)
		dir := StringTrimRight(dir, 1)
	sk := 0
	Loop, Parse, dir, %s%
	{
		If A_LoopField = ..
		{
			StringLeft, root, root, InStr(root, s, "", 0) - 1
			sk += 3
		}
		Else If A_LoopField = .
			sk += 2
		Else If A_LoopField =
		{
			root =
			sk++
		}
	}
	dir := StringTrimLeft(dir, sk)
	
	Abs := pr . root . s . dir
	If InStr(Abs, s, "", 0) = StrLen(Abs)
		Abs := StringTrimRight(Abs, 1)
	
	Return, Abs
}
; AHK本体を基準とした相対パスを絶対パスに変換
_RelToAbs(Path){
	return % _RelToAbs_From(A_WorkingDir, Path, "\")
}
; 指定コマンドを実行、終了までウェイト
_RunWait(runapp, option="", runcmd=""){
	RunWait, %runapp% %option%,, %runcmd%
}
; 指定文字列を指定リスト(置換前と後の文字列一覧)を参照して一括置換
_ListReplace(str, list="", objName=""){
	target := A_Init_Object[objName]
	If ( target="" ) {
		list   := IfExist(list) ? list : _FileSelectFile("replace", "置換リストを開く", "テキストドキュメント (*.txt)")
		target := FileRead(list)
	}
	
	Loop, parse, target, `n, `r
	{
		If (A_LoopField = "")
			continue
		
		StringSplit, rule, A_LoopField, %A_Tab%
		l := rule1, r := rule2
		
		str := StringReplace(str, l, r, "All")
	}
	return str
}
; 指定文字列を指定リスト(置換前と後の正規表現一覧)を参照して一括正規置換
_ListReplaceRegex(str, list="", objName=""){
	target := A_Init_Object[objName]
	If ( target="" ) {
		list   := IfExist(list) ? list : _FileSelectFile("regex", "正規表現リストを開く", "テキストドキュメント (*.txt)")
		target := FileRead(list)
	}
	
	Loop, parse, target, `n, `r
	{
		If (A_LoopField = "")
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
; 指定文字列を正規表現検索し、一致した先頭マッチ文字列を取得
_RegExMatch_Get(Target, Pattern){
	RegExMatch(Target, Pattern, $)
	return $1
}
; 指定ファイルをEmEditorで開く [EmEditor]
_EmEditor(target, line=""){
	Loop, parse, target, `n, `r
	{
		if (A_LoopField = "")
			continue
		
		IfExist, %A_LoopField%
		{
			path := _WQ( _RelToAbs(A_LoopField) )
			If (line)
				line := "/l " . line
			option := _OptionCombine( path, line )
			_RunIn("..\EmEditor_Portable\EmEditor.exe", option )
		}
	}
}
