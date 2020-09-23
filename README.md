<div align="center">

## Determine the Capabilities of an AVI file


</div>

### Description

MCI Multimedisa control NOT NEEDED! Determine if an AVI (movie) file has AUDIO, VIDEO, REVERSE, TOTAL NUMBER of FRAMES, STRETCH, etc... This is good information to know about an AVI before playing it in your program. You can use this information to help you display a "meter" or a scroll bar to quickly move around in an AVI file. This is easy code, enjoy.
 
### More Info
 
Create a new project with a form (Form1)

Add a command control to the form (Command1)

Have a few AVI (*.avi) files on hand for testing


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Patrick K\. Bigley](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/patrick-k-bigley.md)
**Level**          |Unknown
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/patrick-k-bigley-determine-the-capabilities-of-an-avi-file__1-1917/archive/master.zip)





### Source Code

```
Private Sub Command1_Click()
'NOTE: Some of the routines below obviously do not
'apply to an AVI, such as "Can Eject", but the routines
'within this code applies ALL multimedia (WAV, MIDI, AVI,
'CD Audio, Scanner, DAT, etc...)
Dim mssg As String * 255
Dim Rslt As String
Rslt = "Capabilities of this AVI file:" & vbCrLf & vbCrLf
'We must "open" the AVI file first
 ComStr = "open c:\shut.avi type avivideo alias video1"
 x% = mciSendString(ComStr, 0&, 0, 0&)
'---Can it be played?
x% = mciSendString("capability video1 can play", mssg, 255, 0)
If Left$(mssg, 4) = "true" Then
 Rslt = Rslt & "- Can be played" & vbCrLf
Else
 Rslt = Rslt & "- Cannot be played" & vbCrLf
End If
'---Does it have audio?
x% = mciSendString("capability video1 has audio", mssg, 255, 0)
If Left$(mssg, 4) = "true" Then
 Rslt = Rslt & "- Has audio" & vbCrLf
Else
 Rslt = Rslt & "- Has no audio" & vbCrLf
End If
'---Does it have video?
x% = mciSendString("capability video1 has audio", mssg, 255, 0)
If Left$(mssg, 4) = "true" Then
 Rslt = Rslt & "- Has video" & vbCrLf
Else
 Rslt = Rslt & "- Has no video" & vbCrLf
End If
'---Can it be played in reverse?
x% = mciSendString("capability video1 can reverse", mssg, 255, 0)
If Left$(mssg, 4) = "true" Then
 Rslt = Rslt & "- Can reverse" & vbCrLf
Else
 Rslt = Rslt & "- Cannot reverse" & vbCrLf
End If
'---Can it be stretched?
x% = mciSendString("capability video1 can stretch", mssg, 255, 0)
If Left$(mssg, 4) = "true" Then
 Rslt = Rslt & "- Can stretch" & vbCrLf
Else
 Rslt = Rslt & "- Cannot stretch" & vbCrLf
End If
'---Can it record?
x% = mciSendString("capability video1 can record", mssg, 255, 0)
If Left$(mssg, 4) = "true" Then
 Rslt = Rslt & "- Can record" & vbCrLf
Else
 Rslt = Rslt & "- Cannot record" & vbCrLf
End If
'---Can it eject?
x% = mciSendString("capability video1 can eject", mssg, 255, 0)
If Left$(mssg, 4) = "true" Then
 Rslt = Rslt & "- Can eject" & vbCrLf
Else
 Rslt = Rslt & "- Cannot eject" & vbCrLf
End If
'---Compound Device?
x% = mciSendString("capability video1 compound device", mssg, 255, 0)
If Left$(mssg, 4) = "true" Then
 Rslt = Rslt & "- Compound device = TRUE" & vbCrLf
Else
 Rslt = Rslt & "- Compound device = FALSE" & vbCrLf
End If
'---Uses file(s)?
x% = mciSendString("capability video1 uses files", mssg, 255, 0)
If Left$(mssg, 4) = "true" Then
 Rslt = Rslt & "- Uses file(s)" & vbCrLf
Else
 Rslt = Rslt & "- Does not use file(s)" & vbCrLf
End If
'---Does this use palettes?
x% = mciSendString("capability video1 uses palettes", mssg, 255, 0)
If Left$(mssg, 4) = "true" Then
 Rslt = Rslt & "- Uses palettes" & vbCrLf
Else
 Rslt = Rslt & "- Does not use palettes" & vbCrLf
End If
'---Can it save?
x% = mciSendString("capability video1 can save", mssg, 255, 0)
If Left$(mssg, 4) = "true" Then
 Rslt = Rslt & "- Can be saved" & vbCrLf
Else
 Rslt = Rslt & "- Cannot be saved" & vbCrLf
End If
'Close the AVI file
x% = mciSendString("close video1", 0&, 0, 0&)
 MsgBox Rslt, , "Results"
End Sub
```

