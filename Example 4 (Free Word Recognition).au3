#include "./Include/Utter.au3"

_Utter_Speech_CreateFreeGrammar(5000,"_spchin") ; Create free grammar and load words from system dictionary lasting for 3 seconds

Func _spchin($dummy)
ConsoleWrite($dummy)
Switch $dummy
Case "exit" ;exit the script when the user says exit
   Exit

EndSwitch
EndFunc