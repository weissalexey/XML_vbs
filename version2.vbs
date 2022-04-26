


Set con = CreateObject("ADODB.Connection")
	With con
		.Provider = "SQLOLEDB"
		.Properties("Data Source") = "192.168.114.15"
		.ConnectionString = "user id = gobabygo; password=comeback"
		.Open
		.DefaultDatabase = "WinSped"
	End With
   sql = "select * from V_Schaden_VTL_941"
   Set result = con.Execute(sql)


    If Not result.EOF  Then
      conter = 0 
	  result.MoveFirst
	  While Not result.EOF
	  	dim LN, naim, FE, DOCDAT
	  	  LN = result.Fields("LiefNr").Value 
	  	  naim = result.Fields("SW_NVE").Value
          FE = result.Fields("FileExtension").Value 
	  	  DOCDAT = result.Fields("DocumentData").Value
               '  MsgBox DOCDAT
            ITOg = _ 
            "<?xml version=""1.0"" encoding=""utf-8""?>"& chr(13) & chr(10) _
            &"<DamageReports>"& chr(13) & chr(10) _
            &"    <Sender>04245</Sender>"& chr(13) & chr(10) _
            &"    <DamageReport>"& chr(13) & chr(10) _
            &"        <Author>04245</Author>"& chr(13) & chr(10) _
            &"        <Contact>kj</Contact>"& chr(13) & chr(10) _
            &"        <Fon>0461 95707 0</Fon>"& chr(13) & chr(10) _
            &"        <Email>jm@carstensen.eu</Email>"& chr(13) & chr(10) _
            &"        <NOC>"& LN &"</NOC>"& chr(13) & chr(10) _
            &"        <Details>"& chr(13) & chr(10) _
            &"            <Detail>"& chr(13) & chr(10) _
            &"                <NVE>"& naim &"</NVE>"& chr(13) & chr(10) _
            &"                <Description>Siehe passende NVE Statusmeldungen</Description>"& chr(13) & chr(10) _
            &"                <Documents>"& chr(13) & chr(10) _
            &"                    <Document>"& chr(13) & chr(10) _
            &"                        <File>" 
           
                              
 	    writelog ITOg, conter
	    
        writelog str_to_base64 (DOCDAT) & "</File>" & chr(13) & chr(10) , conter

	     ITOg = "                        <FileType>"& FE &"</FileType>"& chr(13) & chr(10) _
            &"                    </Document>"& chr(13) & chr(10) _
            &"                </Documents>"& chr(13) & chr(10) _
            &"            </Detail>"& chr(13) & chr(10) _
            &"        </Details>"& chr(13) & chr(10) _
            &"    </DamageReport>"& chr(13) & chr(10) _
            &"</DamageReports>"& chr(13) & chr(10) 

            
  
      
       writelog ITOg, conter
       conter = conter + 1
	   result.movenext
	  wend
	end if

sub WriteLog(  logstr , conter )

Const ForAppending = 8
A = year(Now)
B=Month(Now)
if len(B) < 2  then b="0" & B end if
C = day (Now)
if len(C)<2 then C ="0" & C end if
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objLogFile = objFSO.OpenTextFile("" & A & B & C & "_"& conter &".xml" , ForAppending, TRUE)
objLogFile.Write(logstr)

End Sub

Function btoa(sourceStr)
    Dim i, j, n, carr, rarr(), a, b, c
    carr = Array("A", "B", "C", "D", "E", "F", "G", "H", _
            "I", "J", "K", "L", "M", "N", "O" ,"P", _
            "Q", "R", "S", "T", "U", "V", "W", "X", _
            "Y", "Z", "a", "b", "c", "d", "e", "f", _
            "g", "h", "i", "j", "k", "l", "m", "n", _
            "o", "p", "q", "r", "s", "t", "u", "v", _
            "w", "x", "y", "z", "0", "1", "2", "3", _
            "4", "5", "6", "7", "8", "9", "+", "/")
    n = Len(sourceStr)-1
    ReDim rarr(n\3)
    For i=0 To n Step 3
        a = AscW(Mid(sourceStr,i+1,1))
        If i < n Then
            b = AscW(Mid(sourceStr,i+2,1))
        Else
            b = 0
        End If
        If i < n-1 Then
            c = AscW(Mid(sourceStr,i+3,1))
        Else
            c = 0
        End If
        rarr(i\3) = carr(a\4) & carr((a And 3) * 16 + b\16) & carr((b And 15) * 4 + c\64) & carr(c And 63)
    Next
    i = UBound(rarr)
    If n Mod 3 = 0 Then
        rarr(i) = Left(rarr(i),2) & "=="
    ElseIf n Mod 3 = 1 Then
        rarr(i) = Left(rarr(i),3) & "="
    End If
    btoa = Join(rarr,"")
End Function


Function char_to_utf8(sChar)
    Dim c, b1, b2, b3
    c = AscW(sChar)
    If c < 0 Then
        c = c + &H10000
    End If
    If c < &H80 Then
        char_to_utf8 = sChar
    ElseIf c < &H800 Then
        b1 = c Mod 64
        b2 = (c - b1) / 64
        char_to_utf8 = ChrW(&HC0 + b2) & ChrW(&H80 + b1)
    ElseIf c < &H10000 Then
        b1 = c Mod 64
        b2 = ((c - b1) / 64) Mod 64
        b3 = (c - b1 - (64 * b2)) / 4096
        char_to_utf8 = ChrW(&HE0 + b3) & ChrW(&H80 + b2) & ChrW(&H80 + b1)
    Else
    End If
End Function

Function str_to_utf8(sSource)
    Dim i, n, rarr()
    n = Len(sSource)
    ReDim rarr(n - 1)
    For i=0 To n-1
        rarr(i) = char_to_utf8(Mid(sSource,i+1,1))
    Next
    str_to_utf8 = Join(rarr,"")
End Function

Function str_to_base64(sSource)
    str_to_base64 = btoa(str_to_utf8(sSource))
    'str_to_base64 = btoa(str_to_utf8(sSource))
End Function

'test

'msgbox btoa("Hello")   'SGVsbG8=
'msgbox btoa("Hell")    'SGVsbA==

'msgbox str_to_base64("中文한국어")  '5Lit5paH7ZWc6rWt7Ja0
