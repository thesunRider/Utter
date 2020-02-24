#include <GUIConstants.au3>
#include "./Include/Utter.au3"

GUICreate("MIC level Detection test", 400, 150)
$prgrs = GUICtrlCreateProgress(20, 20, 350, 35)
GUICtrlCreateLabel("Sound level:", 140, 70)
$idLabel = GUICtrlCreateLabel("0%", 220, 70)
$exit = GUICtrlCreateButton("Exit", 150, 90, 80)
GUISetState(@SW_SHOW)

While (1)
	Switch GUIGetMsg()
		Case $GUI_EVENT_CLOSE, $exit
			Exit
	EndSwitch

	$percent = _Utter_MIC_GetLevel()
	GUICtrlSetData($idLabel, $percent)
	GUICtrlSetData($prgrs, $percent)
WEnd
