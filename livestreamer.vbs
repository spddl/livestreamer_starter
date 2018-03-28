Dim name, sett, msg
Dim objShell : Set objShell = WScript.CreateObject("WScript.Shell")
Dim FileSystemObject : Set FileSystemObject = CreateObject("Scripting.FileSystemObject")
Dim dict : Set dict = CreateObject("Scripting.Dictionary")
Dim token : token = "rm0hl2lkyga4aq9bziy8x46uoe3ykq"

msg =       "0: Best (default)" & vbCR
msg = msg & "1: 1080p60" & vbCR
msg = msg & "2: 1080p" & vbCR
msg = msg & "3: 900p60" & vbCR
msg = msg & "4: 720p60" & vbCR
msg = msg & "5: 720p" & vbCR
msg = msg & "6: 480p" & vbCR
msg = msg & "7: 360p" & vbCR
msg = msg & "8: 160p" & vbCR
msg = msg & "9: audio_only" & vbCR

sett = InputBox(msg, "Make your selection")
If Len(sett) = 0 Then WScript.Quit

Select Case sett
    Case "9"
        sett = "audio_only"
    Case "8"
        sett = "160p"
    Case "7"
        sett = "360p"
    Case "6"
        sett = "480p"
    Case "5"
        sett = "720p"
    Case "4"
        sett = "720p60"
    Case "3"
        sett = "900p60"
    Case "2"
        sett = "1080p"
    Case "1"
        sett = "1080p60"
    Case Else
        sett = "best"
End Select

ls = "Streamer? or pick one Number:" & vbCR & vbCR
Set file = FileSystemObject.OpenTextFile(".\LastStreamer.ini", 1, true)
LineCount = 0

Do Until file.AtEndOfStream
    Dim line : line = file.ReadLine()
	if not len(line) = 0 then
		LineCount = LineCount + 1
		ls = ls & "(" & LineCount & ") " & line & vbCR
		dict.Add CStr(LineCount), line
	end if
Loop

name = Inputbox(ls)
If Len(name) = 0 Then WScript.Quit

If dict.Exists(name) Then
    objShell.Run("livestreamer --twitch-oauth-token " & token & " twitch.tv/" & dict.Item(name) & " " & sett)
Else
    objShell.Run("livestreamer --twitch-oauth-token " & token & " twitch.tv/" & name & " " & sett)
    Set objFile = FileSystemObject.CreateTextFile(".\LastStreamer.ini", True)
    objFile.WriteLine name
    a = dict.Items
    For i = 0 To 3
        wscript.echo i
        if dict.Count > i then
            If not name = a(i) Then objFile.WriteLine a(i)
		else
			objFile.WriteLine vbNewLine
		end if
    Next
    objFile.Close
End If

Set objShell = Nothing
