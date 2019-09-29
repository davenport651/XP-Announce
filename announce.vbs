strAgentName = "Merlin"
strAgentPath = "c:\windows\msagent\chars\" & strAgentName & ".acs"
Set objAgent = CreateObject("Agent.Control.2")

objAgent.Connected = TRUE
objAgent.Characters.Load strAgentName, strAgentPath
Set objCharacter = objAgent.Characters.Character(strAgentName)

objCharacter.Show

objCharacter.Balloon.Style = 0

objCharacter.MoveTo 500,400
objCharacter.Play "Announce"
objCharacter.Speak "SysAdmin has a special announcement for User!"
objCharacter.Play "think"
objCharacter.Think "Hmmm, now what was it again?"
wscript.sleep 500
objCharacter.think "Oh yes."
objCharacter.play "read"
objCharacter.Speak "All your base are belong to us."
objCharacter.Speak "Stay Frosty!"
onjCharacter.Speak "--SysAdmin"
objCharacter.play "readreturn"


objCharacter.Play "Pleased"
objCharacter.Speak "Good day."
objCharacter.Hide

Do While objCharacter.Visible = TRUE
    Wscript.Sleep 100
Loop

