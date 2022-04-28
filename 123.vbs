


Set con = CreateObject("ADODB.Connection")
	With con
		.Provider = "SQLOLEDB"
		.Properties("Data Source") = ""
		.ConnectionString = "user id =; password="
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
strHexValue = TextToHex(DOCDAT)
' Convert hexadecimal bytes into Base64 characters.
strBase64 = HexToBase64(strHexValue)
 
	    writelog  strBase64 & "</File>" & chr(13) & chr(10) , conter
        'writelog str_to_base64 (DOCDAT) & "</File>" & chr(13) & chr(10) , conter
        'writelog Base64Encode (DOCDAT, True) & "</File>" & chr(13) & chr(10) , conter

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
'
end Sub

Function TextToHex(strText)
    ' Function to convert a text string into a string of hexadecimal bytes.
    Dim strChar, k

    TextToHex = ""
    For k = 1 To Len(strText)
        strChar = Mid(strText, k, 1)
        TextToHex = TextToHex & Right("00" & Hex(Asc(strChar)), 2)
    Next
End Function

Function HexToBase64(strHex)
    ' Function to convert a hex string into a base64 encoded string.
    ' Constant B64 has global scope.
    Dim lngValue, lngTemp, lngChar, intLen, k, j, strWord, str64, intTerm

    intLen = Len(strHex)

    ' Pad with zeros to multiple of 3 bytes.
    intTerm = intLen Mod 6
    If (intTerm = 4) Then
        strHex = strHex & "00"
        intLen = intLen + 2
    End If
    If (intTerm = 2) Then
        strHex = strHex & "0000"
        intLen = intLen + 4
    End If

    ' Parse into groups of 3 hex bytes.
    j = 0
    strWord = ""
    HexToBase64 = ""
    For k = 1 To intLen Step 2
        j = j + 1
        strWord = strWord & Mid(strHex, k, 2)
        If (j = 3) Then
            ' Convert 3 8-bit bytes into 4 6-bit characters.
            lngValue = CCur("&H" & strWord)

            lngTemp = Fix(lngValue / 64)
            lngChar = lngValue - (64 * lngTemp)
            str64 = Mid(B64, lngChar + 1, 1)
            lngValue = lngTemp

            lngTemp = Fix(lngValue / 64)
            lngChar = lngValue - (64 * lngTemp)
            str64 = Mid(B64, lngChar + 1, 1) & str64
            lngValue = lngTemp

            lngTemp = Fix(lngValue / 64)
            lngChar = lngValue - (64 * lngTemp)
            str64 = Mid(B64, lngChar + 1, 1) & str64

            str64 = Mid(B64, lngTemp + 1, 1) & str64

            HexToBase64 = HexToBase64 & str64
            j = 0
            strWord = ""
        End If
    Next
    ' Account for padding.
    If (intTerm = 4) Then
       ' HexToBase64 = Left(HexToBase64, Len(HexToBase64) - 1) & "="
    End If
    If (intTerm = 2) Then
        'HexToBase64 = Left(HexToBase64, Len(HexToBase64) - 2) & "=="
    End If

End Function
