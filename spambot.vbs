Dim intResult
Set wshshell = wscript.CreateObject("WScript.Shell")
usermsg=InputBox("Pls message to spam :D")
times=InputBox("How many times? :P")
delaytime=InputBox("Delay Between Messages? ._. (For discord use '1000')")
intAnswer = Msgbox("Wanna use anti-spam bypass? xd",vbYesNo,"ANTI-SPAM BYPASS")

If intAnswer = vbYes Then
    antispam=InputBox("Enter a number, and the bot will chose a random one from 0 to your input! It's for antispam bypass. I recommend using 10M")
    msgbox("Starting spamming: '" & usermsg & "' '" & times & "' times with '" & delaytime & "' ms delay in 3 seconds!")
    wscript.sleep 3000
    For i = 1 To times
    Randomize
    intResult = Int((antispam * Rnd) +1)
    wshshell.sendkeys usermsg
    wshshell.sendkeys " | "
    wshshell.sendkeys intResult
    wshshell.sendkeys "{ENTER}"
    wscript.sleep delaytime
    Next
Elseif intAnswer = vbNo Then
    msgbox("Starting spamming: '" & usermsg & "' '" & times & "' times with '" & delaytime & "' ms delay in 3 seconds!")
    wscript.sleep 3000
    For i = 1 To times  
    wshshell.sendkeys usermsg
    wshshell.sendkeys " "
    wshshell.sendkeys "{ENTER}"
    wscript.sleep delaytime
    Next
End If
