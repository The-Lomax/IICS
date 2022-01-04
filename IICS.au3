#include <INet.au3>
#Include <WinAPI.au3>
#include <Date.au3>
#include <Array.au3>
#include <GuiButton.au3>
#include <GuiEdit.au3>
#include <GuiToolBar.au3>
#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>

Global $trc, $fname, $txtbox, $agent, $tm, $lob, $fault, $ltestresult, $ltest, $trcda
Global $note, $ticket, $nocpe, $dpa, $fun, $txt, $tlob, $fres, $hwnd, $formres, $formreset
Global $cpe[10], $cpename[10]
$fname="IICS - Internal Industrial Communication Solution"
$Form1 = GUICreate($fname, 620, 480, 192, 124)
GUISetBkColor(0x008080)
$txtbox = GUICtrlCreateEdit("", 16, 128, 321, 289, BitOR($ES_AUTOVSCROLL,$ES_WANTRETURN,$WS_VSCROLL,$WS_BORDER))
$loginbtn = GUICtrlCreateButton("Login", 8, 8, 161, 33)
$agntlbl = GUICtrlCreateLabel("Agent:", 8, 48, 48, 20, $SS_RIGHT)
GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
$tmlbl = GUICtrlCreateLabel("TM:", 8, 72, 48, 20, $SS_RIGHT)
GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
$loblbl = GUICtrlCreateLabel("LoB:", 8, 96, 48, 20, $SS_RIGHT)
GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
$agentname = GUICtrlCreateLabel("", 56, 48, 277, 20)
GUICtrlSetFont(-1, 10, 400, 0, "MS Sans Serif")
$tmname = GUICtrlCreateLabel("", 56, 72, 277, 20)
GUICtrlSetFont(-1, 10, 400, 0, "MS Sans Serif")
$lobname = GUICtrlCreateLabel("", 56, 96, 277, 20)
GUICtrlSetFont(-1, 10, 400, 0, "MS Sans Serif")
$Group1 = GUICtrlCreateGroup("Have you done a line test?", 368, 112, 233, 49)
GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
$ltestyes = GUICtrlCreateRadio("yes", 416, 136, 49, 17)
$ltestno = GUICtrlCreateRadio("no", 520, 136, 41, 17)
GUICtrlCreateGroup("", -99, -99, 1, 1)
$Group2 = GUICtrlCreateGroup("CPE Checks you've completed:", 368, 176, 233, 130)
GUICtrlSetFont(-1, 9, 800, 0, "MS Sans Serif")
$sparephone = GUICtrlCreateCheckbox("Spare Phone", 384, 190, 105, 17)
$sparerouter = GUICtrlCreateCheckbox("Spare Router", 490, 190, 105, 17)
$filter = GUICtrlCreateCheckbox("Spare Filter", 384, 214, 105, 17)
$dslcab = GUICtrlCreateCheckbox("DSL Cable", 490, 214, 105, 17)
$msocket = GUICtrlCreateCheckbox("Master Socket", 384, 238, 105, 17)
$tsocket = GUICtrlCreateCheckbox("Test Socket", 384, 262, 105, 17)
$phoneonly = GUICtrlCreateCheckbox("Phone Only", 384, 286, 105, 17)
$routeronly = GUICtrlCreateCheckbox("Router Only", 490, 262, 105, 17)
$manconf = GUICtrlCreateCheckbox("Manual Config", 490, 238, 105, 17)
$dport = GUICtrlCreateCheckbox("Data Port", 490, 286, 105, 17)
GUICtrlCreateGroup("", -99, -99, 1, 1)
$faultname = GUICtrlCreateCombo("", 368, 20, 233, 25, BitOR($CBS_DROPDOWNLIST,$CBS_AUTOHSCROLL,$CBS_SORT))
GUICtrlSetData(-1, " |Authentication Failure|Browsing Issue|Call Features Issue|Crossed Line Fault|Damaged Wiring|Destination Fault|Device Issue|Email Fault|Go Live Query|High Open Fault|Intermittent Sync Problem|IP Address Issue|IPS Migration Problem|No Sync Issue|Outage|Partial Loss Inbound|Partial Loss Outbound|Quality of Service / Noisy Line|Router Replacement|SafeGuard Issue|Self Care Portal Fault|Slow Speeds Issue|Ticket Update|TLOS / No Dial Tone|WiFi Problem")
GUICtrlSetFont(-1, 10, 400, 0, "MS Sans Serif")
$Group3 = GUICtrlCreateGroup("Have you raised a ticket?", 368, 312, 233, 49)
GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
$ticketyes = GUICtrlCreateRadio("yes", 416, 336, 57, 17)
$ticketno = GUICtrlCreateRadio("no", 520, 336, 65, 17)
GUICtrlCreateGroup("", -99, -99, 1, 1)
$Group4 = GUICtrlCreateGroup("DPA", 368, 56, 233, 49)
GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
$dpayes = GUICtrlCreateRadio("yes", 416, 80, 57, 17)
$dpano = GUICtrlCreateRadio("no", 520, 80, 65, 17)
GUICtrlCreateGroup("", -99, -99, 1, 1)
$header = GUICtrlCreateLabel("Fault Report", 440, 0, 89, 20)
GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
$notegen = GUICtrlCreateButton("Generate", 400, 386, 171, 49)
GUICtrlSetFont(-1, 14, 800, 0, "MS Sans Serif")
$freset = GUICtrlCreateButton("Reset", 362, 438, 120, 34)
GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
$ftrans = GUICtrlCreateButton("Transfer", 485, 438, 120, 34)
GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
$formres = GUICtrlCreateCheckbox("Resolved", 450, 368, 105, 15)
GUICtrlCreateLabel("Copyright © 2016 by Krzysztof Beszterda", 8, 458, 196, 17)
GUISetState(@SW_SHOW)

Func WriteNote()
	If WinExists("Untitled - Notepad") Then
		WinActivate("Untitled - Notepad")
		WinWaitActive("Untitled - Notepad")
		$hwnd = ControlGetHandle("", "", "Edit1")
		ControlSetText($hwnd, "", "", $txt)
	Else
		$fun = ShellExecute ("Notepad.rdp")
		WinWait("Untitled - Notepad")
		WinActivate("Untitled - Notepad")
		WinWaitActive("Untitled - Notepad")
		$hwnd = ControlGetHandle("", "", "Edit1")
		ControlSetText($hwnd, "", "", $txt)
	EndIf
EndFunc


While 1
	$r = GUIGetMsg()
	Switch $r
		Case $GUI_EVENT_CLOSE
			Exit
		Case $loginbtn
			Do
				$vn=InputBox($fname,"Enter Your Name:")
				$err=@error
			Until $err=0
			GUICtrlSetData($agentname,$vn)
			$agent=$vn
			Do
				$vn=InputBox($fname,"Enter your TM's Name:")
				$err=@error
			until $err=0
			GUICtrlSetData($tmname,$vn)
			$tm=$vn
			Do
				$vn=InputBox($fname,"Enter Your Line of Business:")
				$err=@error
			until $err=0
			GUICtrlSetData($lobname,$vn)
			GUIctrlSetState($loginbtn,32)
			$lob=$vn
		Case $ltestyes
			Do
				$vn=InputBox($fname,"Did it find a NETWORK fault? (y / n)")
				$err=@error
			until $err=0 Or $err=1 And $vn="y" Or $vn="n" Or $vn=""
			if ($vn="y") Then
				$ltestresult="it came back with a network fault. "
			ElseIf ($vn="n") Then
				$ltestresult="it found no network fault. "
			ElseIf ($err=1 Or $vn="") Then
				GUICtrlSetState($ltestyes, $GUI_UNCHECKED)
			EndIf
		Case $dpayes
			Do
				$vn=InputBox($fname,"Customer or DA? (c / d)")
				$err=@error
			until $err=0 Or $err=1 And $vn="c" Or $vn="d" Or $vn=""
			if ($vn="c") Then
				$dpa="DPA confirmed with the Customer. "
				$trcda="Customer"
			ElseIf ($vn="d") Then
				$dpa="DPA confirmed with a DA. "
				$trcda="DA"
			ElseIf ($err=1 Or $vn="") Then
				GUICtrlSetState($dpayes, $GUI_UNCHECKED)
			EndIf
		Case $ticketyes
			Do
				$vn=InputBox($fname,"TRC & AVC confirmed? (y / n)")
				$err=@error
			until $err=0 Or $err=1 And $vn="y" Or $vn="n" Or $vn=""
			If($vn="y") Then
				$trc="TRC & AVC confirmed with the " & $trcda & ". "
			ElseIf($vn="n") Then
				$trc="TRC & AVC not required for this fault. "
			ElseIf ($err=1 Or $vn="") Then
				GUICtrlSetState($ticketyes, $GUI_UNCHECKED)
			EndIf
		Case $notegen
			$fault = GUICtrlRead($faultname)
			$data = _Now()
			
			$ltest1 = GUICtrlRead($ltestyes)
			if ($ltest1 = 1) Then
				$ltest = "I've done a line test and " & $ltestresult & ". "
			else
				$ltest = "Line test not performed. "
			EndIf
			
			$dpa1 = GUICtrlRead($dpayes)
			if ($dpa1 = 4) Then
				$dpa = "DPA not confirmed. "
				$trcda = "Customer"
			EndIf
			
			$ticket1 = GUICtrlRead($ticketyes)
			If ($ticket1 = 1) Then
				$ticket = "Ticket has been raised to the network. " & $trc
			Else
				$ticket = "No ticket has been raised. "
			EndIf
			
			$formreset = GUICtrlRead($formres)
			
			$nocpe = 0
			$cpe[0] = GUICtrlRead($sparephone)
			if ($cpe[0] = 1) Then
				$cpename[0]="- Tried spare phone."
				$nocpe += 1
			EndIf
			$cpe[1] = GUICtrlRead($sparerouter)
			if ($cpe[1] = 1) Then
				$cpename[1]="- Tried spare router."
				$nocpe += 1
			EndIf
			$cpe[2] = GUICtrlRead($msocket)
			if ($cpe[2] = 1) Then
				$cpename[2]="- Tried connecting into the Master Socket."
				$nocpe += 1
			EndIf
			$cpe[3] = GUICtrlRead($tsocket)
			if ($cpe[3] = 1) Then
				$cpename[3]="- Tried connecting into the Test Socket."
				$nocpe += 1
			EndIf
			$cpe[4] = GUICtrlRead($phoneonly)
			if ($cpe[4] = 1) Then
				$cpename[4]="- Tried connecting only the phone."
				$nocpe += 1
			EndIf
			$cpe[5] = GUICtrlRead($manconf)
			if ($cpe[5] = 1) Then
				$cpename[5]="- Router manually re-configured."
				$nocpe += 1
			EndIf
			$cpe[6] = GUICtrlRead($filter)
			if ($cpe[6] = 1) Then
				$cpename[6]="- Tried spare microfilter."
				$nocpe += 1
			EndIf
			$cpe[7] = GUICtrlRead($dport)
			if ($cpe[7] = 1) Then
				$cpename[7]="- Resetted Data Port."
				$nocpe += 1
			EndIf
			$cpe[8] = GUICtrlRead($routeronly)
			if ($cpe[8] = 1) Then
				$cpename[8]="- Tried connecting only the router."
				$nocpe += 1
			EndIf
			$cpe[9] = GUICtrlRead($dslcab)
			if ($cpe[9] = 1) Then
				$cpename[9]="- Tried spare DSL cable."
				$nocpe += 1
			EndIf
			
			if $nocpe=0 Then
				$nocpe = "No CPE checks done. "
			Else
				$nocpe = "CPE checks done so far:"
			EndIf
			
			$header = "----------------------------------------------" & @CRLF & "Agent: " & $agent & @CRLF & "TM: " & $tm & @CRLF & "LoB: " & $lob & @CRLF & $data & @CRLF & $dpa & @CRLF & @CRLF
			$core = $trcda & " called in to report " & $fault & ". " & $ltest & @CRLF & $nocpe & @CRLF
			For $i=0 To 9
				If ($cpe[$i]=1) Then
					$core=$core & $cpename[$i] & @CRLF
				EndIf
			Next
			If $formreset=1 Then
				$fres = "Fault Resolved."
				$core = $core & $ticket & $fres & @CRLF
			Else
				$core = $core & $ticket & @CRLF
			EndIf
			$footer = "----------------------------------------------"
			
			Switch $fault
				Case $fault="" Or $fault=" "
					$txt = $header & @CRLF & $footer
					ControlSetText($fname, "", $txtbox, $header & @CRLF & $footer)
					;WriteNote()
				Case $fault="Go Live Query"
					$txt = $header & $fault & $footer
					ControlSetText($fname, "", $txtbox, $txt)
					;WriteNote()
				Case $fault="Ticket Update"
					$txt = $header & $trcda & " called in for a " & $fault & @CRLF & $footer
					ControlSetText($fname, "", $txtbox, $txt)
					;WriteNote()
				Case $fault="Outage"
					$txt = $header & $trcda & " called in reporting issue with the service. Informed about the " & $fault & ". " & @CRLF & $footer
					ControlSetText($fname, "", $txtbox, $txt)
					;WriteNote()
				Case Else
					$txt = $header & $core & $footer
					ControlSetText($fname, "", $txtbox, $txt)
					;WriteNote()
			EndSwitch
		Case $ftrans
			$data = _Now()
			
			$dpa1 = GUICtrlRead($dpayes)
			if ($dpa1 = 4) Then
				$dpa = "DPA not confirmed. "
				$trcda = "Customer"
			EndIf
			
			Do
				$tlob1 = InputBox($fname,"Transfer to?" & @CRLF & "C - CS" & @CRLF & "R - Retentions" & @CRLF & "S- Sales" & @CRLF & "P- Provisioning")
				$err=@error
			until $err=0 Or $err=1 And $tlob1="c" Or $tlob1="r" Or $tlob1="s" Or $tlob1="p" Or $tlob1=""
			If($tlob1="c") Then
				$tlob="Customer Services Department."
			ElseIf($tlob1="r") Then
				$tlob="Retentions Department."
			ElseIf($tlob1="s") Then
				$tlob="Sales Department."
			ElseIf($tlob1="p") Then
				$tlob="Provisioning Department."
			ElseIf ($err=1 Or $tlob="") Then
				ContinueLoop
			EndIf
			
			$header = "----------------------------------------------" & @CRLF & "Agent: " & $agent & @CRLF & "TM: " & $tm & @CRLF & "LoB: " & $lob & @CRLF & $data & @CRLF & $dpa & @CRLF & @CRLF
			$core = "Customer called in regarding non-tech issue. Call transferred to " & $tlob & @CRLF
			$footer = "----------------------------------------------"
			$txt = $header & $core & $footer
			
			ControlSetText($fname, "", $txtbox, $txt)
			;WriteNote()
		Case $freset
			GUICtrlSetState($sparephone, $GUI_UNCHECKED)
			GUICtrlSetState($sparerouter, $GUI_UNCHECKED)
			GUICtrlSetState($msocket, $GUI_UNCHECKED)
			GUICtrlSetState($tsocket, $GUI_UNCHECKED)
			GUICtrlSetState($filter, $GUI_UNCHECKED)
			GUICtrlSetState($dport, $GUI_UNCHECKED)
			GUICtrlSetState($dslcab, $GUI_UNCHECKED)
			GUICtrlSetState($phoneonly, $GUI_UNCHECKED)
			GUICtrlSetState($manconf, $GUI_UNCHECKED)
			GUICtrlSetState($ltestyes, $GUI_UNCHECKED)
			GUICtrlSetState($ltestno, $GUI_UNCHECKED)
			GUICtrlSetState($dpayes, $GUI_UNCHECKED)
			GUICtrlSetState($routeronly, $GUI_UNCHECKED)
			GUICtrlSetState($dpano, $GUI_UNCHECKED)
			GUICtrlSetState($ticketyes, $GUI_UNCHECKED)
			GUICtrlSetState($ticketno, $GUI_UNCHECKED)
			GUICtrlSetState($formres, $GUI_UNCHECKED)
			GUICtrlSetData($faultname, " ", " ")
			ControlSetText($fname, "", $txtbox, "")
	EndSwitch
WEnd
