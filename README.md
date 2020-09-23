<div align="center">

## Sunset \- SunRise


</div>

### Description

It returns the time of the sunset or sunrise when you supply the latitude and longtitude
 
### More Info
 
I didnt have time to clean it. just add a text box command button and date picker


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[JoelL](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/joell.md)
**Level**          |Advanced
**User Rating**    |4.0 (8 globes from 2 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) , VBA MS Access, VBA MS Excel
**Category**       |[Complete Applications](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/complete-applications__1-27.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/joell-sunset-sunrise__1-35484/archive/master.zip)





### Source Code

```

  ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
  ':::                                     :::
  '::: These functions calculate sunrise and sunset times for any given    :::
  '::: latitude and longitude. They may also be used to calculate such    :::
  '::: things as astronomical twilight, nautical twilight and civil      :::
  '::: twilight.                               :::
  ':::                                     :::
  '::: SPECIAL NOTES: This code is valid for dates from 1901 to 2099, and   :::
  ':::         will not calculate sunrise/set times for latitudes   :::
  ':::         above/below 63/-63 degrees.               :::
  ':::                                     :::
  '::: This code is based on the work of several others, including Jean    :::
  '::: Meeus, Todd Guillory, Christophe David, Kieth Burnett and Roger W.  :::
  '::: Sinnott (credit where due!)                      :::
  ':::                                     :::
  '::: Converted to VBScript and cleaned-up by Mike Shaffer.         :::
  ':::                                     :::
  ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
  Const pi = 3.14159265358979
 Public degrees, radians As Variant
  ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
  ':::  Returns an angle in range of 0 to (2 * pi)              :::
  ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
  Function GetRange(x)
   Dim temp1
   Dim temp2
   temp1 = x / (2 * pi)
   temp2 = (2 * pi) * (temp1 - Fix(temp1))
   If temp2 < 0 Then
     temp2 = (2 * pi) + temp2
   End If
   GetRange = temp2
  End Function
  ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
  ':::  Returns 24 hour time from decimal time                :::
  ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
  Function GetMilitaryTime(DecimalTime, GMTOffset)
   Dim temp1
   Dim temp2
   ' Handle 24-hour time wrap
   If DecimalTime + GMTOffset < 0 Then DecimalTime = DecimalTime + 24
   If DecimalTime + GMTOffset > 24 Then DecimalTime = DecimalTime - 24
   temp1 = Abs(DecimalTime + GMTOffset)
   temp2 = Int(temp1)
   temp1 = 60 * (temp1 - temp2)
   temp1 = Right("0000" & CStr(Int(temp2 * 100 + temp1 + 0.5)), 4)
   GetMilitaryTime = Left(temp1, 2) & ":" & Right(temp1, 2)
  End Function
  ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
  ':::  This routine does all the real work                  :::
  ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
  Function GetSunRiseSet(latitude, ByVal longitude, ZoneRelativeGMT, RiseOrSet, Year, Month, Day)
   If Abs(latitude) > 63 Then
     GetSunRiseSet = "{invalid latitude}"
     Exit Function
   End If
y = Year
m = Month
d = Day
   ' An altitude of -0.833 is generally accepted as the angle of
   ' the sun at which sunrise/sunset occurs. It is not exactly
   ' zero because of refraction effects of the earth's atmosphere.
   altitude = -0.833
   Select Case UCase(RiseOrSet)
     Case "S"
      RS = -1
     Case Else
      RS = 1
   End Select
   Ephem2000Day = 367 * y - 7 * (y + (m + 9) \ 12) \ 4 + 275 * m \ 9 + d - 730531.5
   utold = pi
   utnew = 0
   sinalt = CDbl(Sin(altitude * radians))  ' solar altitude
   sinphi = CDbl(Sin(latitude * radians))  ' viewer's latitude
   cosphi = CDbl(Cos(latitude * radians))  '
   longitude = CDbl(longitude * radians) ' viewer's longitude
   Err.Clear
   On Error Resume Next
   Do While (Abs(utold - utnew) > 0.001) And (ct < 35)
    ct = ct + 1
    utold = utnew
    days = Ephem2000Day + utold / (2 * pi)
    t = days / 36525
    ' These 'magic' numbers are orbital elements of the sun, and should not be changed
    L = GetRange(4.8949504201433 + 628.331969753199 * t)
    G = GetRange(6.2400408 + 628.3019501 * t)
    ec = 0.033423 * Sin(G) + 0.00034907 * Sin(2# * G)
    lambda = L + ec
    E = -1 * ec + 0.0430398 * Sin(2# * lambda) - 0.00092502 * Sin(4# * lambda)
    obl = 0.409093 - 0.0002269 * t
    ' Obtain ASIN of (SIN(obl) * SIN(lambda))
    delta = Sin(obl) * Sin(lambda)
    delta = Atn(delta / (Sqr(1 - delta * delta)))
    GHA = utold - pi + E
    cosc = (sinalt - sinphi * Sin(delta)) / (cosphi * Cos(delta))
    Select Case cosc
    Case cosc > 1
     correction = 0
    Case cosc < -1
     correction = pi
    Case Else
     correction = Atn((Sqr(1 - cosc * cosc)) / cosc)
    End Select
    utnew = GetRange(utold - (GHA + longitude + RS * correction))
   Loop
   If Err = 0 Then
     GetSunRiseSet = GetMilitaryTime(utnew * degrees / 15, ZoneRelativeGMT)
   Else
     GetSunRiseSet = "{err}"
   End If
  End Function
Private Sub Command1_Click()
If IsDate(MyMaskDate) = False Then Exit Sub
y = Year(MyMaskDate)
m = Month(MyMaskDate)
d = Day(MyMaskDate)
If IsNumeric(Text1.Text) = True Then
  ' Set these to the latitude/longitude of the observer
  Set dbs = OpenDatabase(App.Path & "/Zipcodes.Mdb")
  Set rst = dbs.OpenRecordset("Select * From Zip Where Zip_Code = '" & Text1.Text & "'")
  If rst.RecordCount > 0 Then
    MyLatitude = rst!latitude
    MyLongitude = rst!longitude
  End If
End If
If MyLatitude = "" Then 'This is the setting for brooklyn NY
  MyLatitude = 40.633157
  MyLongitude = -73.996953
End If
' Set this to your offset from GMT (e.g. for Dallas is -6)
' NOTE: The routine does NOT handle switches to/from daylight savings
'    time, so beware!
MyTimeZone = -5
' Note:Set RiseOrSet to "R" for sunrise, "S" for sunset
RiseOrSet = "R"
Rise = GetSunRiseSet(MyLatitude, MyLongitude, MyTimeZone, _
  RiseOrSet, y, m, d)
SSEt = GetSunRiseSet(MyLatitude, MyLongitude, MyTimeZone, _
  "s", y, m, d)
Label1.Caption = Format(Rise, "H:nn AMPM") & vbNewLine & Format(SSEt, "H:nn AMPM")
End Sub
Private Sub Form_Load()
degrees = 180 / pi
  radians = pi / 180
End Sub
```

