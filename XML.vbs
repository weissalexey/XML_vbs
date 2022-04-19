dim filNum

filNum = "36"
filDateStart = "20-12-2006"
	
	Set con = CreateObject("ADODB.Connection")
	With con
		.Provider = "SQLOLEDB"
		.Properties("Data Source") = "itan2"
		.Properties("Integrated Security") = "SSPI"
		.Open
		.DefaultDatabase = "contour"
	End With
	 


	
	Dim fso, ts
	Const ForWriting = 2
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set ts = fso.OpenTextFile("fstart_load.0"& filNum, ForWriting, True) 
	
	 ' ��������� ������
	 ts.Write _ 
 "<?xml version=""1.0"" encoding=""windows-1251""?>"& chr(13) & chr(10) _
 &"<?Contour version=""1.2"" ?>"& chr(13) & chr(10) _
 &"<�����>"& chr(13) & chr(10) _
 &"<���������������������>"& chr(13) & chr(10) _
 &" <������>"& filNum &"</������>"& chr(13) & chr(10) _
 &" <����>" & filDateStart & "</����>"& chr(13) & chr(10) _
 &" <�����C�����>0</�����C�����>"& chr(13) & chr(10) _
 &" <��������������������>���</��������������������>"& chr(13) & chr(10) _
 &" <�������������������>0</�������������������>"& chr(13) & chr(10) _
 &" <������������>09-08-2006</������������>"& chr(13) & chr(10) _
 &" <�������������>13:59:09</�������������>"& chr(13) & chr(10) _
 &" <��������>2.0.02.01</��������>"& chr(13) & chr(10) _
 &" <�����������������������������>01-01-2004</�����������������������������>"& chr(13) & chr(10) _
 &" <����������������������������>0</����������������������������>"& chr(13) & chr(10) _
 &" <����������������������>N</����������������������>"& chr(13) & chr(10) _
 &" <��������������������>���</��������������������>"& chr(13) & chr(10) _
 &" <����������������������>���</����������������������>"& chr(13) & chr(10) _
 &"</���������������������>"& chr(13) & chr(10)

   ' ������� �������������
    ts.Write _ 	
 "<������������ ���="""& filNum &"_�������"">"& chr(13) & chr(10) _
 &"  <��������>����������� ������������ ������� " & filNum & "</��������>"& chr(13) & chr(10) _
 &"  <������>"& filNum &"</������>"& chr(13) & chr(10) _
 &"  <������������>E" & filNum & "1000001</������������>"& chr(13) & chr(10) _
 &"</������������>"& chr(13) & chr(10)

   ts.WriteLine "<������������������������ ���������=""ST""/>"
	
	
  ' ������� �������
   ts.Write _ 	
"<������� ���=""�����_���"" ���="""& filNum &"_���������"">"& chr(13) & chr(10) _
&"  <��������>����������� ������ �������   "& filNum &" </��������>"& chr(13) & chr(10) _
&"  <��������>�</��������>"& chr(13) & chr(10) _
&"  <�����_���>5</�����_���>"& chr(13) & chr(10) _
&"  <����������>10</����������>"& chr(13) & chr(10) _
&"  <�������������>"& chr(13) & chr(10) _
&"    <�������������>"& chr(13) & chr(10) _
&"      <���>������</���>"& chr(13) & chr(10) _
&"      <���_��>�</���_��>"& chr(13) & chr(10) _
&"      <����>�</����>"& chr(13) & chr(10) _
&"    </�������������>"& chr(13) & chr(10) _
&"  </�������������>"& chr(13) & chr(10) _
&"</�������>"& chr(13) & chr(10) 
	
	ts.WriteLine "<������������������������ ���������=""SB""/>" & chr(13) & chr(10) & chr(13) & chr(10) & chr(13) & chr(10) 
	
	'����� ���������
	dim sql,result 
	sql = "select * from dacleval where ev_type = 13 and ev_itcode like '[0123456789][0123456789][0123456789]'"
	Set result = con.Execute(sql)
	
	If Not result.EOF  Then
	  result.MoveFirst
	  While Not result.EOF
	  	dim cod,naim
	  	  cod = result.Fields("ev_itcode").Value 
	  	  naim = result.Fields("ev_name").Value 
	  	  
	  	  ts.WriteLine GenegateConvAcc(cod,naim,"�",filNum,filDateStart)
	  	  ts.WriteLine GenegateConvAcc(cod,naim,"�",filNum,filDateStart)
	  	  ts.WriteLine GenegateConvAcc(cod,naim,"�",filNum,filDateStart)
	  	  ts.WriteLine GenegateConvAcc(cod,naim,"�",filNum,filDateStart)
	  	  if cod = "810" then
	  	    ts.WriteLine GenegateConvAcc(cod,naim,"�",filNum,filDateStart)
	  	  end if
	   result.movenext
	  wend
	end if
	
	
	ts.WriteLine "<������������������������ ���������=""FA""/>"
	

   ts.Write _
"<��������������>"& chr(13) & chr(10) _
&"  <������>1</������>"& chr(13) & chr(10) _
&"</��������������>"& chr(13) & chr(10) _
&"</�����>"& chr(13) & chr(10)

   
	
	ts.Close
	
	
function GenegateConvAcc (val,val_n, gl, fil, dat_open)
	dim tp 
	tp = "�"
	if val = "810" then 
	  tp = "�"
	end if
dim str
  str = _
"<����������� �����=""00000/"& val & "_" & gl& """>"& chr(13) & chr(10) _
&"  <������>"& fil &"</������>"& chr(13) & chr(10) _
&"  <������������>"& fil &"_�������</������������>"& chr(13) & chr(10) _
&"  <������>"& fil &"_���������</������>"& chr(13) & chr(10) _
&"  <������>"& val &"</������>"& chr(13) & chr(10) _
&"  <��������>�����_����</��������>"& chr(13) & chr(10) _
&"  <��������>"& val_n & "</��������>"& chr(13) & chr(10) _
&"  <����������������>A�</����������������>"& chr(13) & chr(10) _
&"  <���>��</���>"& chr(13) & chr(10) _
&"  <����������������>"& chr(13) & chr(10) _
&"    <����������������>"& chr(13) & chr(10) _
&"      <������������>"& dat_open & "</������������>"& chr(13) & chr(10) _
&"    </����������������>"& chr(13) & chr(10) _
&"  </����������������>"& chr(13) & chr(10) _
&"  <�������������� ����������=""1"">"& chr(13) & chr(10) _
&"   <��������������>"& chr(13) & chr(10) _
&"      <����������>��_��_"& gl & "</����������>"& chr(13) & chr(10) _
&"      <�����������������>00000_"& tp &"</�����������������>"& chr(13) & chr(10) _
&"    </��������������>"& chr(13) & chr(10) _
&"  </��������������>"& chr(13) & chr(10) _
&"</�����������>"& chr(13) & chr(10)

  GenegateConvAcc = str	
end function