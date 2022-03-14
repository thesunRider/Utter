#AutoIt3Wrapper_UseX64=n
#include "./Include/Utter.au3"

_Example()

Func _Example()
	Local $txtout = "This is an example of Speech synthesis using utter"

	Local $voice = _Utter_Voice_StartEngine() ;Initializes Voice Engine
	_Utter_Voice_Setvolume($voice, 80) ;Changes Volume to 80% (Default is 100)
	Local $save = FileSaveDialog("Save speech as mp3 to", @ScriptDir, "All wav (*.wav)", 0, "Speech.wav")
	_Utter_Voice_Speak($voice, $txtout, True) ;Speaks the text and pauses the script until speaking is finished
	_Utter_Voice_Transcribe($voice, $save, $txtout) ; Converts speech to wav and saves it to $save
	_Utter_Voice_Shutdown($voice) ;Shutdown the engine
EndFunc   ;==>_Example
