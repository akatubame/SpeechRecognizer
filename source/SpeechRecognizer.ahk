﻿#Persistent
#SingleInstance, Force
;#NoTrayIcon

#Include *i <CommonHeader>
#Include *i <SpeechRecognition>

; このアプリ専用のグローバル変数の格納オブジェクト生成
global A_Init_Object
global io := A_Init_Object["SpeechRecognizer"] := Object()

;-------------------------------------------
; 初期設定
;-------------------------------------------

; 関連ファイルのパスを指定
io.IniFile       := A_ScriptDir "\" "SpeechRecognizer.ini"
io.HotkeyFile    := A_ScriptDir "\" "Hotkey.ini"
Menu, Tray, Icon, % A_ScriptDir "\" "SpeechRecognizer.ico"

; 各種グローバル変数の設定
io.tbl := Object()             ; 音声認識対応表
io.s   := new SpeechRecognizer ; 音声認識スクリプトの呼出し

; 動作順序の定義
Gosub, Init
SetTimer, MainTimer, 100
OnExit, ExitSub
return

; 必ず後で読み込む
#Include *i <GUIFunctions>

;-------------------------------------------
; プログラム開始処理
;-------------------------------------------

; 初期動作
Init:
	io.Gui     := _ObjFromFile(io.IniFile)
	io.HK      := _ObjFromFile(io.HotkeyFile)
	io.ctr     := GetCtrAll()
	io.thisGui := ""
	
	Gosub, Hotkey_Build
	For key in io.Gui {
		io.thisGui := io.Gui[key]
		GoSub, GUI_Build
	}
	io.tbl := io.ctr.SpeechList.ItemObj.ItemList ; 音声認識対応表の同期
return

;-------------------------------------------
; 制御ルーチン
;-------------------------------------------

; 音声認識処理（タイマー）
MainTimer:
	Main()
return
Main(){
	; 挙動選択のGUI窓アクティブ時
	IfWinActive, ahk_group GuiGroup
	{
		return
	}
	; それ以外
	Else
	{
		; GUI消去
		For i,thisGui in io.Gui
			GUI_Hide(thisGui)
		
		; 音声認識
		io.s.Recognize(True)
		io.text := io.s.Prompt()
		
		; 誤作動を無視 (１文字だけ、語頭に促音などの音声認識は不正と見なす)
		If ( StrLen(io.text) == 1 or SubStr(io.text, 1, 1) == "っ" )
			return
		
		; 認識ワードで指定動作を実行
		matchFlag := false
		For i,thisItem in io.tbl {
			
			; 認識ワードが音声認識対応表のワードと一致するまで検索
			For j in thisItem["keyword"] {
				If ( io.text == thisItem["keyword"][j] ){
					matchFlag := true
					Break
				}
			}
			; 一致すれば対応する関数を実行
			If (matchFlag) {
				_ExecFunc(thisItem["func"], thisItem["option"]*)
				Break
			}
		}
		
		; 見つからない場合、該当する挙動を選択
		If (!matchFlag) {
			io.thisGui.Title := "挙動の選択 - 「" io.text "」"
			GUI_Show(io.thisGui)
			;_RunWaitClose("挙動の選択 ahk_class AutoHotkeyGUI")
		}
		SoundPlay, *64
	}
}

; イベント振分け処理
Event:
	Gosub, SetUp
	; 項目をクリックした時のイベント
	If (A_GuiEvent == "Normal") {
		ID := GetFocusItem(io.thisCtr, 1)
		_AddToObj(io.tbl[ID].keyword, io.text)
		GUI_Hide(io.thisGui)
	}
return

; 単純終了サブルーチン
Exit:
ExitApp

; 終了時の処理
ExitSub:
	io.Exit := 1
	Gosub, GUI_Save
ExitApp

; セーブ処理
Save:
	io.Exit := 0
	Gosub, GUI_Save
	GoSub, GUI_Load
return

; ウィンドウのサイズ変更時のイベント
GuiSize:
	; 最小化
	If (A_EventInfo == 1) {
		return
	}
	;; それ以外
	Else {
		For i,thisCtr in GetGui().Ctr
			CTL_Size(thisCtr)
	}
return

; ウィンドウを閉じた時のイベント
GuiClose:
GuiEscape:
	GUI_Hide(io.thisGui)
return
