Dim intResult
Set wshshell = wscript.CreateObject("WScript.Shell")
usermsg=InputBox("Enter The Message You Want To Spam!")
times=InputBox("Enter The Amount Of Messages You Want To Send!")
delaytime=InputBox("Enter The Delay Between Sending Each Message In Millisecond! (Use '1000' For Discord)")
intAnswer = Msgbox("Do You Want To Use Anti-Spam Bypasser?",vbYesNo,"ANTI-SPAM BYPASS")

If intAnswer = vbYes Then
    antispam=InputBox("Enter A Value For Anti-Spam Bypasser! The Higher Is The Bett")
    msgbox("Spam Message: '" & usermsg & "'" & vbcrlf & "Amount Of Messages: '" & times & "'" & vbcrlf & "Delay Between Messages: '" & delaytime & "'" & vbcrlf & "Is Anti-Spam Bypass On: YES" & vbcrlf & "Anti-Spam Bypass Value: '" & antispam & "'")
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
    msgbox("Spam Message: '" & usermsg & "'" & vbcrlf & "Amount Of Messages: '" & times & "'" & vbcrlf & "Delay Between Messages: '" & delaytime & "'" & vbcrlf & "Is Anti-Spam Bypass On: NO")
    wscript.sleep 3000
    For i = 1 To times  
    wshshell.sendkeys usermsg
    wshshell.sendkeys " "
    wshshell.sendkeys "{ENTER}"
    wscript.sleep delaytime
    Next
End If
