Set wshshell = wscript.CreateObject("WScript.Shell")
hack=InputBox("Pls message to spam :D")
times=InputBox("How many times? :P")

msgbox("Starting spamming: " & hack & " " & times & " times in 3 seconds!")
wscript.sleep 3000

For i = 1 To times
wshshell.sendkeys hack
wshshell.sendkeys "{ENTER}"
wscript.sleep 1000
Next