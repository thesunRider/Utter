#include "./Include/Utter.au3"

_Example()

Func _Example()
	Local $mhm[4] = ["jarvis", "Autoit", "Surya", "Exit"] ;Creates an array for recognition
	Local $mgm[3] = ["good", "bad", "out"] ;Creates an array for after recognition

	; The second array words should be said with the first array word to get the result
	; examples : "jarvis good","Exit out","Autoit good"

	Local $recognize = _Utter_Speech_StartEngine()
	_Utter_Speech_CreateGrammar($recognize, $mhm, $mgm) ;Recognize using the arrays
	_Utter_Speech_CreateTokens($recognize)
	_Utter_Speech_GrammarRecognize($recognize, "", 3000, _spchin) ; start the recognition lasting for 3 seconds
	_Utter_Speech_ShutdownEngine() ;Shutdown the engine
EndFunc   ;==>_Example

Func _spchin($dummy)
	ConsoleWrite($dummy)
	Switch $dummy
		Case "exit out" ;exit when user says exit out
			Exit

	EndSwitch
EndFunc   ;==>_spchin
