Dim objShell, name, sett, msg
Set objShell = WScript.CreateObject( "WScript.Shell" )

msg =       "0: Best (default)" & vbCR
msg = msg & "1: 1080p60" & vbCR
msg = msg & "2: 1080p" & vbCR
msg = msg & "3: 720p60" & vbCR
msg = msg & "4: 720p" & vbCR
msg = msg & "5: 480p" & vbCR
msg = msg & "6: 360p" & vbCR
msg = msg & "7: 160p" & vbCR
msg = msg & "8: audio_only" & vbCR

sett = InputBox(msg,"Make your selection")

Select Case sett
    Case "8"
        sett = "audio_only"
    Case "7"
        sett = "160p"
    Case "6"
        sett = "360p"
    Case "5"
        sett = "480p"
    Case "4"
        sett = "720p"
    Case "3"
        sett = "720p60"
    Case "2"
        sett = "1080p"
    Case "1"
        sett = "1080p60"
    Case Else
        sett = "best"
End Select

name = Inputbox("Streamer?")
If Len(name) = 0 Then WScript.Quit

objShell.Run("livestreamer --twitch-oauth-token <-https://twitchapps.com/tmi-> twitch.tv/"&name&" "&sett)
Set objShell = Nothing
