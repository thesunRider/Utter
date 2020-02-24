#include "./Include/Utter.au3"

$txtout = "This is an example of Speech synthesis using utter"

$voice = _Utter_Voice_StartEngine() ;Initializes Voice Engine
_Utter_Voice_Setvolume($voice, 80) ;Changes Volume to 80% (Default is 100)
$save = FileSaveDialog("Save speech as mp3 to", @ScriptDir, "All wav (*.wav)", 0, "Speech.wav")
_Utter_Voice_Speak($voice, $txtout, True) ;Speaks the text and pauses the script until speaking is finished
_Utter_Voice_Transcribe($voice, $save, $txtout) ; Converts speech to wav and saves it to $save
_Utter_Voice_Shutdown($voice) ;Shutdown the engine
