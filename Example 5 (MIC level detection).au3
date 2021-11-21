#include <GUIConstants.au3>
#include "./Include/Utter.au3"

_Example()

Func _Example()
	GUICreate("MIC level Detection test", 400, 150)
	Local $prgrs = GUICtrlCreateProgress(20, 20, 350, 35)
	GUICtrlCreateLabel("Sound level:", 140, 70)
	Local $idLabel = GUICtrlCreateLabel("0%", 220, 70)
	Local $exit = GUICtrlCreateButton("Exit", 150, 90, 80)
	GUISetState(@SW_SHOW)

	Local $percent
	While 1
		Switch GUIGetMsg()
			Case $GUI_EVENT_CLOSE, $exit
				Exit
		EndSwitch

		$percent = _Utter_MIC_GetLevel()
		GUICtrlSetData($idLabel, $percent)
		GUICtrlSetData($prgrs, $percent)
	WEnd

EndFunc   ;==>_Example
