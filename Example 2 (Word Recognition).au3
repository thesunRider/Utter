#AutoIt3Wrapper_UseX64=n
#include "./Include/Utter.au3"

_Example()

Func _Example()
	;  here "|" is default GUIDataSeparatorChar delimiter so string will be split from the delimiter
	Local $recognize = _Utter_Speech_StartEngine() ;  initializes the engine
	_Utter_Speech_CreateGrammar($recognize, "Love|Car|Exit|Follow|Autoit") ;  Creates a grammar with the words
	_Utter_Speech_CreateTokens($recognize) ;  Creates a token for registering speech recognition
	_Utter_Speech_GrammarRecognize($recognize, "", 0, _spchin) ;  Starts the recognition and calls the word recognized to the _spchin function
	While 1
		Sleep(50)
	WEnd
	_Utter_Speech_ShutdownEngine() ;  shutdowns the function
EndFunc   ;  ==>_Example


Func _spchin($dummy)
	ConsoleWrite($dummy)
	Switch $dummy
		Case "exit" ;  Exit the script when the user says exit
			Exit

	EndSwitch
EndFunc   ;  ==>_spchin
