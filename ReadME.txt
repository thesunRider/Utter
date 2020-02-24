------------------------------- Utter [3.0.0.1] -------------------------------------

Modified ......: 11/04/2016
Version .......: 3.0.0.1
Author ........: Surya
------------------------[Introduction]----------------------------
Utter is a free ware windows API automation script.It can do most of the sapi dll
functions."SAPI" stands for Windows Speech Reconition API,SAPI.dll is the file
which manages the speech recognition of windows Utter utilises most of the SAPI 
functions making use of the best potential of SAPI.dll,You can include speech 
recognition to your project by using utter

-------------------[CAUTION!]------------------------------
REMEMBER TO SHUTDOWN THE INSTANCE OF CREATED RECOGNITION 
ENGINE BEFORE STARTING ANOTHER INSTANCE IF YOU START ANOTHER
WITHOUT SHUTTING THE PREVIOUS ONE DOWN IT WILL LEAD TO AN ERROR!

REMEMBER THAT "|" IS THE DEFAULT GUIDataSeparatorChar CHANGE IT 
ACCORDING TO YOUR NEEDS AND GRAMMAR DELIMITER IS GUIDataSeparatorChar
IF NO GUIDataSeparatorChar IS FOUND IN THE INPUT STRING THEN THE
ENTIRE STRING WOULD BE CONSIDERED AS ONE WORD!

DO NOT CALL THE INTERNAL FUNCTIONS THEY ARE TO BE CALLED INSIDE THE FUNCTION AND DO NOT 
CHANGE THE VALUE OF VARIABLES USED IN THE FUNCTION!

THE RECIEVING FUNCTIONS SHOULD HAVE ATLEAST ONE PARAMETER TO ACCEPT THE SPEECH COMMANDS
FROM THE _Utter_Speech_GrammarRecognize() FUNCTION
------------[Functions]---------------
  _Utter_Speech_StartEngine()
  _Utter_Speech_CreateGrammar()
  _Utter_Speech_CreateTokens()
  _Utter_Speech_GrammarRecognize()
  _Utter_Speech_Pass()
  _Utter_Speech_ShutdownEngine()
  _Utter_Speech_CreateFreeGrammar()
  _Utter_Speech_FreegrammarPass()
  _Utter_Speech_ShutdownFreeGrammar()
  _Utter_Voice_Setvolume()
  _Utter_Voice_Speak()
  _Utter_Voice_IsSpeaking()
  _Utter_Voice_GetCurrentvoice()
  _Utter_Voice_Setvoice ()
  _Utter_Voice_Getvoice ()
  _Utter_Voice_StartEngine()
  _Utter_Voice_SetRate()
  _Utter_Voice_Transcribe()
  _Utter_Voice_Shutdown()
  _Utter_GetSpeechData()
  _Utter_StartTraining()
  _Utter_MIC_GetLevel()
  _Utter_DebugOut()

------------[Features Available includes]-----------------
- Adding specific words that the computer can recognise
   words are delimited or seperated by the GUIDataSeparatorChar
   string that is currently set the default is "|"

- Pausing and resuming of speech recognition is also available

- Creating Free grammar without any word input recognises
   anything you say with the system dictonary

- Speech synthesis helps to convert text to speech

- Speech transcriber helps to convert speech to a wav file

- Voice customisations include increasing or decreasing volume
   setting pitch of the voice,selecting a slang for speech
   Checking wether the computer is speaking

- Get details about the speech recognition configurations of system
 
- Direct access to microphone input volume stream
-------[Change log]---------
Version [3.0.0.1]:-
* Added more examples.
* Improvised algorithms.
- Added new two data gathering function.
- Added MIC_Getlevel function so as to obtain microphone 
  input level.
- Detailed up explanation of functions.
 
Version [2.1.0.3]:-
* Fixed some bugs
* Removed unnecessary locals and Globals
- Added parameter to _Utter_Voice_Shutdown() for
  making it more accurate
- Added _Utter_Voice_IsSpeaking() function
- Added extra parameters for resuming and 
  pausing in speaking in _Utter_Voice_Speak() function
* Changed delimiter string to GUIDataSeparatorChar
* Added few more examples to zip

Version [1.0.0.9]:-
- Beta version
**********************************************************************************
The last file has been updated with a compressed zip to reduce file size 
and the UDF has also been updated and various examples have been included
-----------------------------------------------------------------------------------