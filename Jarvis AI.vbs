Set Sapi = Wscript.CreateObject("SAPI.SpVoice")
set wshshell = wscript.CreateObject("wscript.shell")
dim Input
wshshell.run "%windir%\Speech\Common\sapisvr.exe -SpeechUX"
Sapi.speak "Please speak, or type, what you want to open?"
Input=inputbox ("Please speak, or type, what you want to open.")

if Input = "youtube" OR Input = "Youtube"
thenSapi.speak "Opening youtube"
wshshell.run "www.youtube.com"

elseif Input = "Github" OR Input = "github" 
thenSapi.speak "Opening GitHub"
wshshell.run "www.github.com"

elseif Input = "google" OR Input = "Google"
 thenSapi.speak "Opening google"
wshshell.run "www.google.com"

elseif Input = "command prompt" OR Input = "Command prompt" 
thenSapi.speak "Opening command prompt"
wshshell.run "cmd"

elseif Input = "calculator" OR Input = "Calculator" 
thenSapi.speak "Opening calculator"
wshshell.run "calc"

elseif Input = "notepad" OR Input = "Notepad
" thenSapi.speak "Opening notepad"
wshshell.run "notepad"

elseif Input = ""
 then
else
 Sapi.speak "I don't recognize your input, Please try something else"
 end if
end if
end if
end if
end if
end if
end if
