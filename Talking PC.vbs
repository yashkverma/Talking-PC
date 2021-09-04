Call mainsub
Sub controlsub
Dim a
a=InputBox("Enter 1 to re-execute the program or 0 to exit","Talking PC")
if a=1 then Call mainsub
End Sub
Sub mainsub
Dim message, sapi
message=InputBox("What do you want me to say?","Talking PC")
Set sapi=CreateObject("sapi.spvoice")
sapi.Speak message
Call controlsub
End Sub
