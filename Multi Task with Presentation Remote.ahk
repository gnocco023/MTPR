#Requires AutoHotkey v2.0

; The idea behind of this script
; https://www.reddit.com/r/powerpoint/comments/1gyr508/remote_powerpoint_without_focus/


; Dave Denzel Ytac
; 2025/07/09
; v1
; Google Slides and Canva - Requires presenter view
; Ms PowerPoint - does not work if presenter view is present


navkeys := ["{PgUp}", "{PgDn}"]

PgUp::
{		
	PresentationNav(navkeys[1])
}


PgDn::
{	
	PresentationNav(navkeys[-1])
}




PresentationNav(key)
{
	; Google Slides
	slides := WinExist("Presenter view")
	
	; Canva
	canva := WinExist("Presenter Window")
	
	; MS PowerPoint
	mspowerpoint := WinExist("PowerPoint")
	
	
	if slides{
		ExecuteNavCommand(slides, key)
	}
	else if canva{
		ExecuteNavCommand(canva, key)
	}
	else if mspowerpoint{
		ExecuteNavCommand(mspowerpoint, key)
	}
	else{
		Send(key)
	}
	
}



ExecuteNavCommand(hwnd, key)
{
	; Get the PID of the powerpoint or the presentation
	pid := WinGetPID(hwnd)
		
	; Get the active window youre using. e.g. Adobe Photoshop
	title := WinGetTitle("A")
		
	; Activate the presentation
	WinActivate("ahk_pid " pid)
	
	Sleep(100)
		
	; Send the pressed key e.g. PgDn or PgUp
	Send(key)
	
	Sleep(100)
		
	; Return to the active window youre using. e.g. Adobe Photoshop
	WinActivate(title)
}

