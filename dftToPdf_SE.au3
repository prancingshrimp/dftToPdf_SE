#include <GUIConstantsEx.au3>
#include <MsgBoxConstants.au3>
#include <File.au3>
#include <Array.au3>

Opt('MustDeclareVars', 1)

Func _ErrFunc($oError)
	MsgBox(BitOR(0,16), "Error", "There is a problem with Solid Edge!")
	Exit
EndFunc

Local $gui, $guiwidth, $guiheight, $buttonwidth, $buttonheight
Local $button_convert, $button_quit
Local $path, $txt_path, $label
Local $signal

$buttonwidth = 120
$buttonheight = $buttonwidth/4
$guiwidth = 2*$buttonwidth+20
$guiheight = 2*$buttonheight+25

$gui = GUICreate("dftToPdf",$guiwidth,$guiheight,2,2)
$label = GUICtrlCreateLabel("Please select folder with dft files.",5,1*$buttonheight+15,400)
GUICtrlCreateLabel("v0.21",5,2*$buttonheight+5,400)
$button_convert = GUICtrlCreateButton("Select folder",5,5,$buttonwidth,$buttonheight)
$button_quit = GUICtrlCreateButton("Quit",$buttonwidth+15,5,$buttonwidth,$buttonheight)

GUISetState(@SW_SHOW)

While 1
   $signal = GUIGetMsg()
   Select
	  Case $signal = $GUI_EVENT_CLOSE
		 ExitLoop

	  Case $signal = $button_quit
		 ExitLoop

	   Case $signal = $button_convert

		 $txt_path = "T:\16_Technik"
		 $path = FileSelectFolder("Choose the destination folder", $txt_path)
		 If @error Then
			MsgBox(BitOR(0,16), "Error", "No folder has been selected!")
			ContinueLoop
		 EndIf

		 convert($path, $label)
	EndSelect
WEnd
GUIDelete($gui)


Func convert($path, $label)
;~ get all link-files
	Local $sizeLnk, $fileArrayLnk
	Local $proofLnk = 1
	$fileArrayLnk = _FileListToArrayRec($path, "*.lnk", $FLTAR_FILES, $FLTAR_NORECUR, $FLTAR_SORT)
	If @error Then
		If @extended = 9 Then
			$proofLnk = 0
		EndIf
	EndIf

;~ get all dft-files
	Local $sizeDft, $fileArrayDft
	Local $proofDft = 1
	$fileArrayDft = _FileListToArrayRec($path, "*.dft", $FLTAR_FILES, $FLTAR_NORECUR, $FLTAR_SORT)
	If @error Then
		If @extended = 9 Then
			$proofDft = 0
		EndIf
	EndIf

	If $proofLnk = 0 And $proofDft = 0 Then
		MsgBox(BitOR(0,16), "Error", "No drawings found!")
		Return
	EndIf

	If $proofDft = 1 Then
		Local $fileArrayDft_[$fileArrayDft[0]+1][3]
		$fileArrayDft_[0][0] = $fileArrayDft[0]
		For $i = 1 To $fileArrayDft[0]
			$fileArrayDft_[$i][0] = $fileArrayDft[$i]
			$fileArrayDft_[$i][1] = $path & "\" & $fileArrayDft_[$i][0]
			$fileArrayDft_[$i][2] = $path & "\" & StringReplace($fileArrayDft_[$i][0], ".dft", ".pdf")
		Next
	EndIf

	If $proofLnk = 1 Then
		Local $fileArrayLnk_[$fileArrayLnk[0]+1][3]
		$fileArrayLnk_[0][0] = $fileArrayLnk[0]

		Local $tmp
		For $i = 1 To $fileArrayLnk[0]
			$tmp = FileGetShortcut($path & "\" & $fileArrayLnk[$i])
			$fileArrayLnk_[$i][0] = $fileArrayLnk[$i]
			$fileArrayLnk_[$i][1] = $tmp[0]
			$fileArrayLnk_[$i][2] = $path & "\" & StringReplace($fileArrayLnk_[$i][0], ".lnk", ".pdf")
		Next
	EndIf

	If $proofLnk = 1 And $proofDft = 1 Then
		Local $Array[$fileArrayDft_[0][0] + $fileArrayLnk_[0][0] +1] [3]
		$Array[0][0] = $fileArrayDft_[0][0] + $fileArrayLnk_[0][0]

		Local $j = 0
		For $i = 1 To $fileArrayLnk_[0][0]
			$Array[$i][0] =	$fileArrayLnk_[$i][0]
			$Array[$i][1] =	$fileArrayLnk_[$i][1]
			$Array[$i][2] =	$fileArrayLnk_[$i][2]
			$j = $j + 1
		Next
			$j = $j + 1
		For $i = 1 To $fileArrayDft_[0][0]
;~ 			_DebugPrint($j & " " & $i)
			$Array[$j][0] =	$fileArrayDft_[$i][0]
			$Array[$j][1] =	$fileArrayDft_[$i][1]
			$Array[$j][2] =	$fileArrayDft_[$i][2]
			$j = $j + 1
		Next


	ElseIf $proofLnk = 1 And $proofDft = 0 Then
		Local $Array = $fileArrayLnk_

	ElseIf $proofLnk = 0 And $proofDft = 1 Then
		Local $Array = $fileArrayDft_

	EndIf

	GUICtrlSetData($label, "Starting Solid Edge. Please wait... ")

	Local $oErrorHandler = ObjEvent("AutoIt.Error", "_ErrFunc")

	Local $oEdge = ObjCreate("SolidEdge.Application")
	If @error Then
		MsgBox(BitOR(0,16), "_Error_", "There is a problem with Solid Edge!")
		Exit
	EndIf

	$oEdge.Visible = True
	$oEdge.DisplayAlerts = False

	For $i = 1 To $Array[0][0]
		GUICtrlSetData($label, "Working on no. " & $i & " out of " & $Array[0][0] & " dft files.")
		Local $objDoc = $oEdge.Documents.Open($Array[$i][1])
;~ 		$oEdge.DoIdle()
		$objDoc.SaveAs($Array[$i][2])
		$objDoc.Close(False)
	Next

	$oEdge.Quit()
	MsgBox(64, "", "Done!")
	GUICtrlSetData($label, "Please select folder with dft files.")
EndFunc

Func _DebugPrint($s_Text, $line = @ScriptLineNumber)
    ConsoleWrite( _
            "!===========================================================" & @LF & _
            "+======================================================" & @LF & _
            "-->Line(" & StringFormat("%04d", $line) & "):" & @TAB & $s_Text & @LF & _
            "+======================================================" & @LF)
EndFunc   ;==>_DebugPrint