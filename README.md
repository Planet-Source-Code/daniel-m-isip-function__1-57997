<div align="center">

## IsIP Function


</div>

### Description

Checks if string is a valid IP or not. I saw this submitted earlier and made my own better version of it. =)

OK there is now an updated version with fixes so go check it out here:

http://planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=57999&lngWId=1&txtForceRefresh=123020041845765523
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Daniel M](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/daniel-m.md)
**Level**          |Beginner
**User Rating**    |3.5 (14 globes from 4 users)
**Compatibility**  |VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/daniel-m-isip-function__1-57997/archive/master.zip)





### Source Code

```
Private Function IsIP(strIP As String) As Boolean
Dim splitIP() As String, i As Long
IsIP = True 'Starts out as true
splitIP$ = Split(strIP$, ".", -1, 1) 'Split IP to check value
'==============================================================
'Things we must check to verify IP
'1. Make sure each section of IP is not greater than 255
'2. Make sure each section of IP does not contain a negative
'3. Make sure each section of IP is numeric
'==============================================================
 For i = 0 To UBound(splitIP$) 'loop through array and check 3 things
  If IsNumeric(splitIP(i)) = False Then
   IsIP = False
   Exit For
  Else
   If splitIP(i) > 255 Or splitIP(i) < 0 Then
    IsIP = False
    Exit For
   End If
  End If
 Next i
End Function
```

