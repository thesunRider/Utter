#cs --------------------------------------------------------------------------
                  (Utter UDF)

   Utter is a free ware windows API automation script.It can do most of the sapi dll
functions."SAPI" stands for Windows Speech Reconition API,SAPI.dll is the file
which manages the speech recognition of windows Utter utilises most of the SAPI
functions making use of the best potential of SAPI.dll,You can include speech
recdognition to your project by using utter

   Utter UDF (SAPI.au3)needs SAPI installed (mostly pre-installed)on your computer
to work properly.You can use and manage SAPI (Speech Recognition API).Please

--------- CAUTION! ----------
REMEMBER TO SHUTDOWN THE INSTANCE OF CREATED RECOGNITION ENGINE BEFORE STARTING ANOTHER
INSTANCE IF YOU START ANOTHER WITHOUT SHUTTING DOWN IT WILL LEAD TO AN ERROR!

REMEMBER THAT "|" IS THE DEFAULT GUIDataSeparatorChar CHANGE IT ACCORDING TO YOUR NEEDS
AND GRAMMAR DELIMITER IS GUIDataSeparatorChar IF NO GUIDataSeparatorChar IS FOUND IN THE
INPUT STRING THEN THE ENTIRE STRING WOULD BE CONSIDERED AS ONE WORD!

DO NOT CALL THE INTERNAL FUNCTIONS THEY ARE TO BE CALLED INSIDE THE FUNCTION AND DO NOT
CHANGE THE VALUE OF VARIABLES USED INSIDE THE FUNCTION!
----------------------------

Current Functions(USER CALLABLE):-
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

Internal Functions ----------------------------------------------------
                    _getpath($file)
                    _Fetchextension($file)
                    _checkgrammar($atb)
                    _CheckObject($ark)
                    MyErrFunc()
                    RecoContext_Recognition($iStr, $vStr, $iRec, $iRe)
                    SpRecEvent_Recognition($iStr, $vStr, $iRec, $iRe)
					_mciSendUtterError($mciError)
					_mciSendUtter()
					_whoami_currentid()
					Readcmd($pid)

Variables (USER MODIFIABLE) -----------------------------------------
      $UTTER_SPEECH_RECOGNIZE :- Contains currently recognised word
		   _Utter_Speech_CreateFreeGrammar() function can also set this
		   variable as well as the _Utter_Speech_GrammarRecognize()
		   function this variable is changed when each time the user
		   speaks a different word.Changes to empty string when you
		   shutdown the function or start a new instance of the engine
	  $UTTER_SPEECH_VAR :- The word to be recognized the recognition
		   variable changes to empty string each time you create a
		   new instance of the speech engine
	  $UTTER_SPEECH_RECOSTATE :- This variable is set to 0 If the
	      recognition verb matches the currently recognised verb
	  $UTTER_SPEECH_ERROR :- Shows the @error instance of the
	      _Utter_Speech_GrammarRecognize() function
	  $UTTER_SPEECH_RESPONSE :- Sets variable to 0 if any thing is recognised
	      and 1 if Not
      $UTTER_SPEECH_FUNC :- Function to call for after recognising a word
	      and to send that word to a function as paramter.NOTE THE FUNCTION To
	      BE CALLED MUST ACCEPT ONLY ONE PARAMETER.
	  $UTTER_SPEECH_PASS :- Sets the value to 1 if pause parameter (1) is issued
		  to _Utter_Speech_Pass() function and To 0 if resume parameter (1)
		  is issued to _Utter_Speech_Pass()
      $UTTER_SPEECH_DEBUG :- Debug Variable change to true by calling _Utter_DebugOut(True)
	      to print out output to console


The Examples are included within the zip file
Thankyou to all the forum members to help me make up this UDF.
Any update will be notified and released on the forum.

Version .........: 3.0.0.1
Author ..........: SuryaSaradhi.B
Modified ........: 24/08/2015
#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here
#include-once
#include <Constants.au3>
#include <String.au3>

Global $UTTER_SPEECH_RECOGNIZE
Global $UTTER_SPEECH_VAR
Global $UTTER_SPEECH_ERROR
Global $UTTER_SPEECH_RECOSTATE
Global $UTTER_SPEECH_RESPONSE
Global $UTTER_SPEECH_FUNC
Global $UTTER_SPEECH_PASS = 0
Global $UTTER_SPEECH_DEBUG = False
Global Const $SAFT48kHz16BitMono = 38
Global Const $SSFMCreateForWrite = 3
Global Const $SVSFPurgeBeforeSpeak = 2
Global Const $SVSFNLPSpeakPunc = 64
Global $UTTER_MIC_STR[8]
Global $oMyError ;Error registration variable

#Region :User-Functions
;#FUNCTION# ;===============================================================================
; Function Name:    _Utter_Speech_StartEngine()
; Description:		Creates a new instance of Speech Recognition Grammar Object.
;
; Parameter(s):     Null
; Requirement(s):   SAPI Installed on your computer
; Return Value(s):  Returns an empty array of handles to Recognition objects
;                   this array handle is given to other functions in the
;                   group for adding grammar
; Author ........: Surya Saradhi.B
; Version .......: 3.0.0.1
; Modified.......: 11/04/16 by Surya
;============================================================================================
Func _Utter_Speech_StartEngine()
$h_Context = ObjCreate("SAPI.SpInProcRecoContext")
$comer = @error
$h_Recognizer = $h_Context.Recognizer
;Global $h_Grammar = $h_Context.CreateGrammar(1)
$oRecoContext = $h_Context ;ObjCreate('SAPI.SpSharedRecoContext')
$oVBS = ObjCreate("ScriptControl")
$oVBS.Language = "VBScript"
$oNothing = $oVBS.Eval("Nothing")
$oRecoContext.CreateGrammar(0)
$oGrammar = $oRecoContext.CreateGrammar()
$h_Grammar = $oGrammar
$oMyError = ObjEvent("AutoIt.Error","MyErrFunc")
_nullifyvariable()
Global $array[8]
$array[0] = $oRecoContext ;spinprorecocontext
$array[1] = $h_Recognizer ;main recognizer
$array[2] = 'null' ;Null handler
$array[3] = $oVBS ;Script
$array[4] = $oNothing ;Nothing
$array[5] = $h_Grammar ;Grammar verbs
$array[6] = "Error:" &$comer ;error notifications
$array[7] = $oMyError
Return $array
EndFunc

;#FUNCTION# ;===============================================================================
; Function Name:    _Utter_Speech_CreateGrammar($ark,$gramr="",$trn="")
; Description:		Adds Grammar to the Speech Recognition Dictionary for recognition
;                   the computer will recognize the word only making recognition more
;                   specific and accurate.
;
; Parameter(s):     $ark   :     Array handle from _Utter_Speech_StartEngine() function
;                   $gramr :-    (Main Phrase)Grammar to add to the input grammar can be a one
;                                dimensional array (two or more dimensions will return an error)
;                                word or string seperated by the current GUIDataSeparatorChar.Deafult value of
;                                Opt("GUIDataSeparatorChar") is "|" change it according to your convienience
;                   $trn   :-    (transitional Verb)Word to be said after the main verb.This word is recognised to
;                 				 be coming after the first verb
; Requirement(s):   SAPI Installed on your computer For Accurate Recognition Train your computer before
;                   adding grammar and starting the recognition procedure.
; Return Value(s):  Returns an empty array of handles to Dictionary Objects
;                   this array handle is given to other functions in the
; Author ........: Surya Saradhi.B
; Version .......: 3.0.0.1
; Modified.......: 11/04/16 by Surya
;============================================================================================
Func _Utter_Speech_CreateGrammar(ByRef $ark,$gramr,$trn="")
;run a check function
_CheckObject($ark)
$ftr = _checkgrammar($gramr)
$btr = _checkgrammar($trn)
$oGrammar = $ark[5]
$oTopRule = $oGrammar.Rules.Add('RUN', 1, 0)
$oWordsRule = $oGrammar.Rules.Add('WORDS', BitOR(1, 32), 1)
$oAfterCmdState = $oTopRule.AddState()
For $i = 0 To UBound ($ftr)-1
$oTopRule.InitialState.AddWordTransition($oAfterCmdState, $ftr[$i])
$oAfterCmdState.AddRuleTransition($ark[4], $oWordsRule)
Next
For $i = 0 To UBound ($btr) -1
$oWordsRule.InitialState.AddWordTransition($ark[4], $btr[$i])
Next
;~ $oWordsRule.InitialState.AddWordTransition($oNothing, 'whatever')
$oGrammar.Rules.Commit()
$oGrammar.CmdSetRuleState('RUN', 1)
Local $array[3]
$array[0] = $oTopRule
$array[1] = $oAfterCmdState
$array[2] = $oWordsRule
Return $array[2]
EndFunc

;#FUNCTION# ;===============================================================================
; Function Name:    _Utter_Speech_CreateTokens($ark)
; Description:		Creates tokens for redirecting the input of Microphone into the
;                   main recognition channel.
;
; Parameter(s):     $ark      :- Handle from the _Utter_Speech_StartEngine() function
; Requirement(s):   SAPI Installed on your computer,Microphone installed on your computer
; Return Value(s):  Returns an empty array of handles to token objects
;                   this array handle is given to other functions in the
;                   group for adding grammar
; Author ........: Surya Saradhi.B
; Version .......: 3.0.0.1
; Modified.......: 11/04/16 by Surya
;============================================================================================
Func _Utter_Speech_CreateTokens(ByRef $ark)
$h_Recognizer = $ark[1]
$h_Category = ObjCreate("SAPI.SpObjectTokenCategory")
$h_Category.SetId("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\AudioInput\TokenEnums\MMAudioIn\")
$h_Token = ObjCreate("SAPI.SpObjectToken")
$h_Token.SetId("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\AudioInput\TokenEnums\MMAudioIn\")
$h_Recognizer.AudioInput = $h_Token
Local $array[2]
$array[0] = $h_Category
$array[1] = $h_Token
Return $array[1]
EndFunc

;#FUNCTION# ;===============================================================================
; Function Name:    _Utter_Speech_GrammarRecognize($ark,$grm="",$delay=1000,$func=False)
; Description:		Starts the Speech recognition procedure.$UTTER_SPEECH_ERROR is set To
;					the @error of the Object Handler and $UTTER_SPEECH_VAR is set to The
;					recognising verb.
;
; Parameter(s):     $ark      :- Handle from the _Utter_Speech_StartEngine() function
;                   $grm      :- Grammar to Recognise,if what the user said matches
;                                the grammar you entered it will set the $UTTER_SPEECH_RECOSTATE
;                                to 0 else to 1 if it doesnt matches.if the input string is ""
;                                then the variable will always be 1.the parameter simply sets
;                                the $UTTER_SPEECH_RECOSTATE variable only.
;								 (Default is "" that is variable will always be 1)
;				   $delay(ms) :- The time to stop the script for the object handler to
;								 recognise phrases and word said by the user.you can issue
;								 0 to the parameter to continue until speech shutdown is called
;                                THE SCRIPT WILL PAUSE UNTIL THE TIME OF DELAY PARAMETER
;					$func     :- The function to be called after recognising the word.Defaulty
;								 it will not call any function but will only set a recognized
;								 verb $UTTER_SPEECH_RECOGNIZE variable.If a parameter or a function
;								 name is issued when the user speaks and when it is recognised the
;								 function will be called with the recognised verb as the parameter
;								 NOTE:THE FUNCTION MUST EXIST AND THE FUNCTION SHOULD ONLY ACCEPT
;                                ONE PARAMETER.Issuing this parameter will also set the function
;                    			 variable $UTTER_SPEECH_FUNC to the function name.Function name
;								 must be given in quotes
; Requirement(s):   SAPI Installed on your computer,For Better acuuracy train your computers
;                   speech recognition module.
; Return Value(s):  Returns empty variable containing handle to event handler module.
; Author ........: Surya Saradhi.B
; Version .......: 3.0.0.1
; Modified.......: 11/04/16 by Surya
;============================================================================================
Func _Utter_Speech_GrammarRecognize(ByRef $ark,$grm="",$delay=1000,$func="")
$hk = ObjEvent($ark[0], 'RecoContext_')
$UTTER_SPEECH_ERROR = @error
$UTTER_SPEECH_VAR = $grm
$UTTER_SPEECH_FUNC = ""
If Not $func = "" Then $UTTER_SPEECH_FUNC = $func
Sleep ($delay)
Return $hk
EndFunc

;#FUNCTION# ;===============================================================================
; Function Name:    _Utter_Speech_Pass($in = 0)
; Description:		Pauses and resumes the recognition instance.
;
; Parameter(s):     $in = Defaultly 0 if issued 1 Then recognition instance will stop
;						  giving output which will resume on giving 0 to the function
;						  as a parameter.This will also set the  $UTTER_SPEECH_PASS variable.
; Return Value(s):  Null
; Author ........: Surya Saradhi.B
; Version .......: 3.0.0.1
; Modified.......: 11/04/16 by Surya
;============================================================================================
Func _Utter_Speech_Pass($in = 0)
If $in = 0 Then
   $UTTER_SPEECH_PASS = 0
ElseIf $in = 1 Then
   $UTTER_SPEECH_PASS = 1
Else
   $UTTER_SPEECH_PASS = 1
EndIf
EndFunc

;#FUNCTION# ;===============================================================================
; Function Name:    _Utter_Speech_ShutdownEngine()
; Description:		Shuts down the handles created by _Utter_Speech_StartEngine() function
;                   and nullifies every object.
;
; Parameter(s):     Null
; Requirement(s):   Already called the _Utter_Speech_StartEngine() function
; Return Value(s):  Null
; Author ........: Surya Saradhi.B
; Version .......: 3.0.0.1
; Modified.......: 11/04/16 by Surya
;============================================================================================
Func _Utter_Speech_ShutdownEngine()
_nullifyvariable()
 $hk = 0
 $h_Context = 0
 $h_Recognizer = 0
 $oRecoContext = 0
 $oVBS = 0
 $oGrammar = 0
 $oMyError = 0
 $oTopRule = 0
 $oWordsRule = 0
 $oAfterCmdState = 0
 $h_Category = 0
 $h_Token = 0
 $hk = 0
 $oNothing = 0
EndFunc

;#FUNCTION# ;===============================================================================
; Function Name:    _Utter_Speech_CreateFreeGrammar($delay = 3000,$func = "")
; Description:		Starts the Speech recognition Module with system dictionary loaded
;                   no grammar has to be specified by the user.Delay time can be set to
;					0 if called before a While...WEnd loop.All variables are changed by
;                   the function.Accuracy is very low for this type of recognition
;
; Parameter(s):     $delay       :- The delay time for the script to wait for the verb
;									recognition you can issue 0 if the script is continued
;                                   and a the delay time to delay the script for recognition.
;					$func        :- The function to be called after recognising the word.Defaulty
;								    it will not call any function but will only set a recognized
;								    verb $UTTER_SPEECH_RECOGNIZE variable.If a parameter or a function
;								    name is issued when the user speaks and when it is recognised the
;								    function will be called with the recognised verb as the parameter
;								    NOTE:THE FUNCTION MUST EXIST AND THE FUNCTION SHOULD ONLY ACCEPT
;                                   ONE PARAMETER.Issuing this parameter will also set the function
;                    			    variable $UTTER_SPEECH_FUNC to the function name.Function name
;									must be given in quotes
; Requirement(s):   SAPI installed in your computer it is also highly recommended to train
;					your computer before starting the procedure.
; Return Value(s):  @error of The main Event Handler.
; Author ........: Surya Saradhi.B
; Version .......: 3.0.0.1
; Modified.......: 11/04/16 by Surya
;============================================================================================
Func _Utter_Speech_CreateFreeGrammar($delay = 3000,$func = "")
_nullifyvariable()
If Not $func = "" Then $UTTER_SPEECH_FUNC = $func
 $h_Contextm = ObjCreate("SAPI.SpInProcRecoContext")
 $h_Recognizerm = $h_Contextm.Recognizer
 $h_Grammarm = $h_Contextm.CreateGrammar(1)
$h_Grammarm.Dictationload
$h_Grammarm.DictationSetState(1)

;Create a token for the default audio input device and set it
 $h_Categorym = ObjCreate("SAPI.SpObjectTokenCategory")
$h_Categorym.SetId("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\AudioInput\TokenEnums\MMAudioIn\")
 $h_Tokenm = ObjCreate("SAPI.SpObjectToken")
$h_Tokenm.SetId("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\AudioInput\TokenEnums\MMAudioIn\")
$h_Recognizerm.AudioInput = $h_Tokenm

 $i_ObjInitializedm = 0

 $h_ObjectEventsm = ObjEvent($h_Contextm, "SpRecEvent_")
If @error Then
    $i_ObjInitializedm = 0
Else
    $i_ObjInitializedm = 1
 EndIf
$UTTER_SPEECH_ERROR = $i_ObjInitializedm
 Sleep ($delay)
 Return $i_ObjInitializedm
EndFunc

;#FUNCTION# ;===============================================================================
; Function Name:    _Utter_Speech_FreegrammarPass($in = 0)
; Description:		Pauses and resumes the recognition instance.
;
; Parameter(s):     $in = Defaultly 0 if issued 1 Then recognition instance will stop
;						  giving output which will resume on giving 0 to the function
;						  as a parameter.This will also set the  $UTTER_SPEECH_PASS variable.
; Return Value(s):  Null
; Author ........: Surya Saradhi.B
; Version .......: 3.0.0.1
; Modified.......: 11/04/16 by Surya
;============================================================================================
Func _Utter_Speech_FreegrammarPass($in = 0)
If $in = 0 Then
   $UTTER_SPEECH_PASS = 0
ElseIf $in = 1 Then
   $UTTER_SPEECH_PASS = 1
Else
   $UTTER_SPEECH_PASS = 1
EndIf
EndFunc

;#FUNCTION# ;===============================================================================
; Function Name:    _Utter_Speech_ShutdownFreeGrammar()
; Description:		Nullifies all variables and shutsdown the objects created by
;					the _Utter_Speech_CreateFreeGrammar() function.
;
; Parameter(s):     Null
; Requirement(s):   Already Called _Utter_Speech_CreateFreeGrammar() function
; Return Value(s):  Null
; Author ........: Surya Saradhi.B
; Version .......: 3.0.0.1
; Modified.......: 11/04/16 by Surya
;============================================================================================
Func _Utter_Speech_ShutdownFreeGrammar()
_nullifyvariable()
 $h_Contextm = 0
 $h_Recognizerm = 0
 $h_Grammarm = 0
 $h_Categorym = 0
 $h_Tokenm = 0
 $i_ObjInitializedm = 0
 $h_ObjectEventsm = 0
EndFunc

;#FUNCTION# ;===============================================================================
; Function Name:    _Utter_Voice_Shutdown()
; Description:		Shutsdown Objects created by the _Utter_Voice_StartEngine() function
;
; Parameter(s):     $ark     :- Handle returned from _Utter_Voice_StartEngine()
; Requirement(s):   Already Called _Utter_Voice_StartEngine() function
; Return Value(s):  Null
; Author ........: Surya Saradhi.B
; Version .......: 3.0.0.1
; Modified.......: 11/04/16 by Surya
;============================================================================================
Func _Utter_Voice_Shutdown(ByRef $ark)
 $ark = 0
EndFunc

;#FUNCTION# ;===============================================================================
; Function Name:    _Utter_Voice_Setvolume(ByRef $ark,$val)
; Description:		Sets Volume of _Utter_Voice_StartEngine() function TTS Engine
;
; Parameter(s):     $ark     :- Handle returned from _Utter_Voice_StartEngine()
;					$val     :- Volume to be set (0-100)
; Requirement(s):   Already Called _Utter_Voice_StartEngine() function,SAPI installed
; Return Value(s):  Null
; Author ........: Surya Saradhi.B
; Version .......: 3.0.0.1
; Modified.......: 11/04/16 by Surya
;============================================================================================
Func _Utter_Voice_Setvolume(ByRef $ark,$val)
$oSpeech = $ark
$oSpeech.Volume = $val
EndFunc

;#FUNCTION# ;===============================================================================
; Function Name:    _Utter_Voice_GetCurrentvoice(ByRef $ark)
; Description:		Gets the current voice Activated in utter.
;
; Parameter(s):     $ark     :- Handle returned from _Utter_Voice_StartEngine()
; Requirement(s):   Already Called _Utter_Voice_StartEngine() function,SAPI installed
; Return Value(s):  Current Active Voice Profile
; Author ........: Surya Saradhi.B
; Version .......: 3.0.0.1
; Modified.......: 11/04/16 by Surya
;============================================================================================
Func _Utter_Voice_GetCurrentvoice(ByRef $ark)
$hl = $ark.getvoices.item (0)
$ml = $hl.GetDescription()
Return $ml
EndFunc

;#FUNCTION# ;===============================================================================
; Function Name:    _Utter_Voice_Setvoice (ByRef $Object, $sVoiceName)
; Description:		Sets an available voice from library as active utter voice
;
; Parameter(s):     $ark          :- Handle returned from _Utter_Voice_StartEngine()
;					$sVoiceName   :- Voice Profile Name returned from _Utter_Voice_Getvoice()
;									 function
; Requirement(s):   Already Called _Utter_Voice_StartEngine() function,SAPI installed
; Return Value(s):  Null
; Author ........: Surya Saradhi.B
; Version .......: 3.0.0.1
; Modified.......: 11/04/16 by Surya
;============================================================================================
Func _Utter_Voice_Setvoice (ByRef $Object, $sVoiceName)
	Local $VoiceNames, $VoiceGroup = $Object.GetVoices
	For $VoiceNames In $VoiceGroup
		If $VoiceNames.GetDescription() = $sVoiceName Then
			$Object.Voice = $VoiceNames
			Return
		EndIf
	Next
EndFunc

;#FUNCTION# ;===============================================================================
; Function Name:    _Utter_Voice_Getvoice(ByRef $ark,$bReturn = True)
; Description:		Gets available voice profile names installed on your Computer
;
; Parameter(s):     $ark          :- Handle returned from _Utter_Voice_StartEngine()
;					$bReturn      :- Boolean value paramater if True is issued returns an
;									 array containing the available installed profile names
;									 If False is issued returns a String consisting of profile
;									 names seperated by the '|' character
; Requirement(s):   Already Called _Utter_Voice_StartEngine() function,SAPI installed
; Return Value(s):  Returns an array or String (seperated by '|' delimeter) containing
;                   available profile names installed in your computer
; Author ........: Surya Saradhi.B
; Version .......: 3.0.0.1
; Modified.......: 11/04/16 by Surya
;============================================================================================
Func _Utter_Voice_Getvoice (ByRef $ark,$bReturn = True)
$oSpeech = $ark
Local $sVoices, $VoiceGroup = $oSpeech.GetVoices
	For $Voices In $VoiceGroup
		$sVoices &= $Voices.GetDescription() & '|'
	 Next
	If $bReturn Then Return StringSplit(StringTrimRight($sVoices, 1), '|', 2)
	Return StringTrimRight($sVoices, 1)
EndFunc

;#FUNCTION# ;===============================================================================
; Function Name:    _Utter_Voice_StartEngine()
; Description:		Starts an instance of TTS Engine Object
;
; Parameter(s):     Null
; Requirement(s):   SAPI installed
; Return Value(s):  Returns a handle to TTS Engine Object.
; Author ........: Surya Saradhi.B
; Version .......: 3.0.0.1
; Modified.......: 11/04/16 by Surya
;============================================================================================
Func _Utter_Voice_StartEngine()
    Return ObjCreate ("SAPI.spVoice")
EndFunc

;#FUNCTION# ;===============================================================================
; Function Name:    _Utter_Voice_SetRate(ByRef $ark,$val)
; Description:		Sets the rate of voice OR how fast the voice speaks.
;
; Parameter(s):     $ark    :- Handle returned by the _Utter_Voice_StartEngine() function
;					$val    :- Rate or Speed of the voice (-10 to 10)
; Requirement(s):   Already called _Utter_Voice_StartEngine(),SAPI installed
; Return Value(s):  Null
; Author ........: Surya Saradhi.B
; Version .......: 3.0.0.1
; Modified.......: 11/04/16 by Surya
;============================================================================================
Func _Utter_Voice_SetRate(ByRef $ark,$val)
$oSpeech = $ark
$oSpeech.Rate = $val
EndFunc

;#FUNCTION# ;===============================================================================
; Function Name:    _Utter_Voice_Speak(ByRef $ark,$text)
; Description:		Speaks the input text in the current TTS voice
;                   Speaking Rate and volume could be changed by
;					the other functions
;
; Parameter(s):     $ark    :- Handle returned by the _Utter_Voice_StartEngine() function
;					$text   :- text to speak
;                   $resum  :- Default (False) is to continue script execution while speaking
;                              True to pause the script execution while running it
; Requirement(s):   Already called _Utter_Voice_StartEngine(),SAPI installed
; Return Value(s):  Null
; Author ........: Surya Saradhi.B
; Version .......: 3.0.0.1
; Modified.......: 11/04/16 by Surya
;============================================================================================
Func _Utter_Voice_Speak(ByRef $ark,$text,$resum = False)
$oSpeech = $ark
If Not $resum Then
   $okr = 11
Else
   $okr = 0
EndIf
$oSpeech.Speak($text,$okr)
EndFunc

;#FUNCTION# ;===============================================================================
; Function Name:    _Utter_Voice_IsSpeaking(ByRef $ark)
; Description:		Checks wether utter is currently speaking the text specified by
;                   the user before in the _Utter_Voice_Speak() function
;
; Parameter(s):     $ark    :- Handle returned by the _Utter_Voice_StartEngine() function
; Requirement(s):   Already called _Utter_Voice_StartEngine(),SAPI installed
; Return Value(s):  True if utter is speaking,False if utter is Not speaking
; Author ........: Surya Saradhi.B
; Version .......: 3.0.0.1
; Modified.......: 11/04/16 by Surya
;============================================================================================
Func _Utter_Voice_IsSpeaking(ByRef $ark)
$oSpeech = $ark
$stt = $oSpeech.Status.RunningState()
If $stt = 1 Then
   Return False
ElseIf $stt = 2 Then
   Return True
Else
   Return False
EndIf
EndFunc

;#FUNCTION# ;===============================================================================
; Function Name:    _Utter_Voice_Transcribe(ByRef $ark,$file,$txt,$comp = 1,$lame="")
; Description:		Transcribes speech to audio OR Converts text to Speech which is
;					being saved to an mp3 or wav fileNo other formats than .wav and .mp3
;					are allowed.System can save the audio as wav only if you want the file
;					to be an mp3 then you will have to issue the lame parameter.
;
; Parameter(s):     $ark    :- Handle returned by the _Utter_Voice_StartEngine() function
;					$file   :- The path with the name of the file to be saved to if to Be
;							   saved as an mp3 then specify the file extension as .mp3
;                              otherwise as .wav
;					$txt    :- The text to convert to audio
;                   $comp   :- Without issuing the parameter (value = 1) the file will be
;							   saved as .wav issue the parameter with 0 value to save the
;							   audio as .mp3 also specify the lame location otherwise this
;							   will return an error.
;                   $lame   :- Location of the Lame.exe file for converting the .wav to .mp3
; Requirement(s):   Already called _Utter_Voice_StartEngine(),SAPI installed
; Return Value(s):  Null
; Author ........: Surya Saradhi.B
; Version .......: 3.0.0.1
; Modified.......: 11/04/16 by Surya
;============================================================================================
Func _Utter_Voice_Transcribe(ByRef $ark,$file,$txt,$comp = 1,$lame="")
$ext = _Fetchextension($file)
If $ext=0 Or $ext = 2 Then
$oStream = ObjCreate ( "SAPI.SpFileStream" )
$ark.AllowAudioOutputFormatChangesOnNextSet = True
$oStream.Format.Type = $SAFT48kHz16BitMono
If $ext = 2 Then
   $mh = ".wav"
Else
   $mh = ""
EndIf
$oStream.Open ( $file &$mh, $SSFMCreateForWrite, True )
$ark.AudioOutputStream = $oStream
$ark.Speak ( $txt, $SVSFPurgeBeforeSpeak + $SVSFNLPSpeakPunc )
$oStream.Close ( )
$ark.AudioOutput = $ark.GetAudioOutputs ( '' ).Item ( 0 )
If $ext &$comp = 20  Then
If FileExists ($lame ) Then
RunWait ( '"' &$lame  &'"' &' -q0 "' & $file &'.wav" "' &_getpath($file) & '.mp3"', '', @SW_HIDE)
FileDelete ($file &".wav")
Else
$txt = "Utter Cant Perform conversion operation because lame directory you enterd is not valid"
ConsoleWrite ($txt)
EndIf
EndIf
Else
   $msg = "Utter cant convert text to other formats than .wav lame is required to convert to .mp3"
   ConsoleWrite (@CRLF &$msg)
EndIf
EndFunc

;#FUNCTION# ;===============================================================================
; Function Name:    _Utter_DebugOut($ffl=False)
; Description:		turns on Printing recognised words to the console stream for debugging.
;
; Parameter(s):     $ffl        :-  Defaultly False if set to True will print out recognised
;									words to the console strem.
; Requirement(s):   Null
; Return Value(s):  Null
; Author ........: Surya Saradhi.B
; Version .......: 3.0.0.1
; Modified.......: 11/04/16 by Surya
;============================================================================================
Func _Utter_DebugOut($ffl=False)
$UTTER_SPEECH_DEBUG = $ffl
EndFunc

;#FUNCTION# ;===============================================================================
; Function Name:    _Utter_GetSpeechData()
; Description:		returns the speech data stored in the computer
;
; Parameter(s):     Null
;
; Requirement(s):   SAPI installed
;					The script should be run under administrative privilages so as to get this result.
; Return Value(s):  [Array]
;                   $ary[0]   :  no of times the speech recogntion engine has been trained
;                   $ary[1]   :  Current System TTS voice (not utter voice)
; Author ........: Surya Saradhi.B
; Version .......: 3.0.0.1
; Modified.......: 11/04/16 by Surya
;============================================================================================
Func _Utter_GetSpeechData()
Local $ary[1] = [RegRead(RegRead("HKEY_USERS\" & _Whoami_CurrentId() & "\Software\Microsoft\Speech\RecoProfiles\", "DefaultTokenId") & "\" & RegEnumKey(RegRead("HKEY_USERS\" & _Whoami_CurrentId() & "\Software\Microsoft\Speech\RecoProfiles\", "DefaultTokenId"), 1), "TrainingStatus"),RegRead(RegRead("HKEY_CURRENT_USER\Software\Microsoft\Speech\Voices\","DefaultTokenId"),"")]
Return $ary
EndFunc

;#FUNCTION# ;===============================================================================
; Function Name:    _Utter_StartTraining()
; Description:		Launches the speech training section
;
; Parameter(s):     Null
; Requirement(s):   SAPI installed
;                   The script should be run under administrative privilages so as to get this result
; Return Value(s):  the pid of the launched windows
; Author ........: Surya Saradhi.B
; Version .......: 3.0.0.1
; Modified.......: 11/04/16 by Surya
;============================================================================================
Func _Utter_StartTraining()
Local $pidpros = Run(@SystemDir &"\Speech\SpeechUX\SpeechUXWiz.exe" &' UserTraining,en-GB,HKEY_CURRENT_USER\SOFTWARE\Microsoft\Speech\RecoProfiles\Tokens\' &_StringBetween(RegRead("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Speech\RecoProfiles\","DefaultTokenId"),"{","}")[0] &',656860,0,""')
Return $pidpros
EndFunc

;#FUNCTION# ;===============================================================================
; Function Name:    _Utter_MIC_GetLevel()
; Description:		Retrievs the Linein microphone levels
;
; Parameter(s):     Null
; Requirement(s):   A microphone plugged in to the computer.
; Return Value(s):  A numreical value which shows the microphone
;                   level(the sound which microphone recieves)
; Author ........: Surya Saradhi.B
; Version .......: 3.0.0.1
; Modified.......: 11/04/16 by Surya
;============================================================================================
Func _Utter_MIC_GetLevel()
$UTTER_MIC_STR[0] = "new type waveaudio"
$UTTER_MIC_STR[2] = "alias mywave"
$UTTER_MIC_STR[4] = ""
$UTTER_MIC_STR[5] = StringFormat("open %s %s %s", $UTTER_MIC_STR[0], $UTTER_MIC_STR[2], $UTTER_MIC_STR[4]);
$UTTER_MIC_STR[6] = _StringRepeat(" ", 100) ; Information will return in this string
$UTTER_MIC_STR[7] = StringLen($UTTER_MIC_STR[6])
$mciError = _mciSendUtter($UTTER_MIC_STR[5], $UTTER_MIC_STR[6], $UTTER_MIC_STR[7], 0);
If $mciError[0] <> 0 Then _mciSendUtterError($mciError[0])
$UTTER_MIC_STR[1] = "mywave"
$UTTER_MIC_STR[3] = "level"
$UTTER_MIC_STR[4] = ""
$UTTER_MIC_STR[5] = StringFormat("status %s %s %s", $UTTER_MIC_STR[1], $UTTER_MIC_STR[3], $UTTER_MIC_STR[4]);
$mciError = _mciSendUtter($UTTER_MIC_STR[5], $UTTER_MIC_STR[6], $UTTER_MIC_STR[7], 0);
If $mciError[0] <> 0 Then _mciSendUtterError($mciError[0])
Return $mciError[2]
EndFunc
#EndRegion :User-Functions

;#INTERNAL FUNCTION# ;===============================================================================
; Below Functions are called by the functions of the script they should not be called
; by the User,These are support functions which provide data values for main function
;============================================================================================
#Region :Internal Functions Start
Func Readcmd($cmdin)
	$cmd = Run(@ComSpec & " /c " & $cmdin, @ScriptDir, @SW_HIDE, $STDERR_CHILD + $STDOUT_CHILD)
	Local $sLine, $sOutput
	While 1
		$sLine = StdoutRead($cmd)
		If @error Then ExitLoop
		$sOutput &= $sLine
	WEnd
	Return $sOutput
EndFunc   ;==>Readcmd

Func _Whoami_CurrentId()
	$read = Readcmd("WHOAMI /USER /FO CSV")
	$string = _StringBetween($read, '"' & @ComputerName & "\" & @UserName & '","', '"' & @CRLF)
	Return $string[0]
EndFunc   ;==>_Whoami_CurrentId

Func _nullifyvariable()
 $UTTER_SPEECH_RECOGNIZE = ""
 $UTTER_SPEECH_VAR = ""
 $UTTER_SPEECH_ERROR = ""
 $UTTER_SPEECH_RECOSTATE = ""
 $UTTER_SPEECH_RESPONSE = ""
 $UTTER_SPEECH_FUNC = ""
 $UTTER_SPEECH_PASS = ""
EndFunc

Func _getpath($file)
$spk = StringRegExp($file, "^\h*((?:\\\\\?\\)*(\\\\[^\?\/\\]+|[A-Za-z]:)?(.*[\/\\]\h*)?((?:[^\.\/\\]|(?(?=\.[^\/\\]*\.)\.))*)?([^\/\\]*))$", 1)
Return $spk[1] &$spk[2] &$spk[3]
EndFunc

Func _Fetchextension($file)
$aArray = StringRegExp($file, "^\h*((?:\\\\\?\\)*(\\\\[^\?\/\\]+|[A-Za-z]:)?(.*[\/\\]\h*)?((?:[^\.\/\\]|(?(?=\.[^\/\\]*\.)\.))*)?([^\/\\]*))$", 1)
If $aArray[4] = ".wav" Then
   Return 0
ElseIf $aArray[4] = ".mp3" Then
   Return 2
Else
   Return 1
EndIf
EndFunc

Func _checkgrammar($atb)
$char = Opt("GUIDataSeparatorChar")
If IsArray($atb) Then
   Return $atb
ElseIf StringInStr ($atb,$char) Then
   $split = StringSplit ($atb,$char,2)
   Return $split
Else
   Local $ark[1]
   $ark[0] = $atb
   Return $ark
EndIf
EndFunc

Func _CheckObject($ark)
ObjName ($ark[0])
If @error Then MsgBox ("","Utter","Handler not valid")
EndFunc

Func MyErrFunc()
Msgbox(0,"Utter COM Test","We intercepted a COM Error !"      & @CRLF  & @CRLF & _
             "err.description is: "    & @TAB & $oMyError.description    & @CRLF & _
             "err.windescription:"     & @TAB & $oMyError.windescription & @CRLF & _
             "err.number is: "         & @TAB & hex($oMyError.number,8)  & @CRLF & _
             "err.lastdllerror is: "   & @TAB & $oMyError.lastdllerror   & @CRLF & _
             "err.scriptline is: "     & @TAB & $oMyError.scriptline     & @CRLF & _
             "err.source is: "         & @TAB & $oMyError.source         & @CRLF & _
             "err.helpfile is: "       & @TAB & $oMyError.helpfile       & @CRLF & _
             "err.helpcontext is: "    & @TAB & $oMyError.helpcontext _
            )
    Local $err = $oMyError.number
    If $err = 0 Then $err = -1
    $g_eventerror = $err  ; to check for after this function returns
Endfunc

Func SpRecEvent_SoundStart($StreamNumber, $StreamPosition)
    $UTTER_SPEECH_RESPONSE = 0
EndFunc

Func SpRecEvent_SoundEnd($StreamNumber, $StreamPosition)
    $UTTER_SPEECH_RESPONSE = 1
EndFunc

Func RecoContext_SoundStart($StreamNumber, $StreamPosition)
    $UTTER_SPEECH_RESPONSE = 0
EndFunc

Func RecoContext_SoundEnd($StreamNumber, $StreamPosition)
    $UTTER_SPEECH_RESPONSE = 1
EndFunc

Func RecoContext_Recognition($iStreamNumber, $vStreamPosition, $iRecognitionType, $iResult)
    $said = $iResult.PhraseInfo.GetText()
   If $UTTER_SPEECH_PASS = 0 Then
	If $UTTER_SPEECH_DEBUG Then ConsoleWrite( @CRLF &"--UTTER NOTIFIED:" &$said)
    If Not $UTTER_SPEECH_FUNC = "" Then
	   Call ($UTTER_SPEECH_FUNC,$said)
    EndIf
	If $UTTER_SPEECH_VAR = $UTTER_SPEECH_RECOGNIZE Then
	   $UTTER_SPEECH_RECOSTATE = 0
    Else
	   $UTTER_SPEECH_RECOSTATE = 1
    EndIf
	$UTTER_SPEECH_RECOGNIZE = $said
 Else
	If $UTTER_SPEECH_DEBUG Then ConsoleWrite( @CRLF &"--UTTER PAUSED RECOGNITION STATE")
   EndIf
EndFunc

Func SpRecEvent_Recognition($StreamNumber, $StreamPosition, $RecognitionType, $iResult)
    $said = $iResult.PhraseInfo.GetText()
	If $UTTER_SPEECH_PASS = 0 Then
	If Not $UTTER_SPEECH_FUNC = "" Then
	   Call ($UTTER_SPEECH_FUNC,$said)
    EndIf
    If $UTTER_SPEECH_DEBUG Then ConsoleWrite( @CRLF &"--UTTER NOTIFIED:" &$said)
	$UTTER_SPEECH_VAR = $said
	$UTTER_SPEECH_RECOGNIZE = $said
	 Else
	If $UTTER_SPEECH_DEBUG Then ConsoleWrite( @CRLF &"--UTTER PAUSED RECOGNITION STATE")
   EndIf
EndFunc

Func _mciSendUtter($lpszCommand, $lpszReturnString, $cchReturn, $hwndCallback)
	Return DllCall("winmm.dll", "long", "mciSendStringA", "str", $lpszCommand, "str", $lpszReturnString, "long", $cchReturn, "long", 0)
EndFunc   ;==>_mciSendUtter

Func _mciSendUtterError($mciError)
	Dim $errStr ; Error message
	$errStr = _StringRepeat(" ", 100) ; Reserve some space for the error message
	$Result = DllCall("winmm.dll", "long", "mciGetErrorStringA", "long", $mciError, "string", $errStr, "long", StringLen($errStr))
EndFunc   ;==>_mciSendUtterError
#EndRegion :Internal Functions End
