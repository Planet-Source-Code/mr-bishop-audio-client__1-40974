Attribute VB_Name = "ElapTm"
Public Function ElapsedTime(tStart, tStop) As String

' *******************************************************************
' Function Name : ElapsedTime *
' Created By    : Herry Hariry Amin *
' Email         : h2arr@yahoo.com *
' Language      : VB4, VB5, VB6 *
' Example       : sYourVariable = ElapsedTime(tStartTime,tStopTime) *
' *******************************************************************

Dim dtr, dtl, jml As Long

dtl = (Hour(tStart) * 3600) + _
(Minute(tStart) * 60) + (Second(tStart))

dtr = (Hour(tStop) * 3600) + _
(Minute(tStop) * 60) + (Second(tStop))

If tStop < tStart Then
  jml = 86400
Else
  jml = 0
End If

jml = jml + (dtr - dtl)

ElapsedTime = Format(Str(Int((Int((jml / 3600)) Mod 24))), "00") _
+ ":" + Format(Str(Int((Int((jml / 60)) Mod 60))), "00") + ":" + _
Format(Str(Int((jml Mod 60))), "00")

End Function

