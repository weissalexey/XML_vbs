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
	
	 ' Заголовок сеанса
	 ts.Write _ 
 "<?xml version=""1.0"" encoding=""windows-1251""?>"& chr(13) & chr(10) _
 &"<?Contour version=""1.2"" ?>"& chr(13) & chr(10) _
 &"<Сеанс>"& chr(13) & chr(10) _
 &"<КонтрольныйЛистСеанса>"& chr(13) & chr(10) _
 &" <Филиал>"& filNum &"</Филиал>"& chr(13) & chr(10) _
 &" <Дата>" & filDateStart & "</Дата>"& chr(13) & chr(10) _
 &" <НомерCеанса>0</НомерCеанса>"& chr(13) & chr(10) _
 &" <ЕстьДокументыПозднее>Нет</ЕстьДокументыПозднее>"& chr(13) & chr(10) _
 &" <КонтрольноеЗначение>0</КонтрольноеЗначение>"& chr(13) & chr(10) _
 &" <ДатаВыгрузки>09-08-2006</ДатаВыгрузки>"& chr(13) & chr(10) _
 &" <ВремяВыгрузки>13:59:09</ВремяВыгрузки>"& chr(13) & chr(10) _
 &" <ВерсияПО>2.0.02.01</ВерсияПО>"& chr(13) & chr(10) _
 &" <ДатаПервогоАрхивногоДокумента>01-01-2004</ДатаПервогоАрхивногоДокумента>"& chr(13) & chr(10) _
 &" <КоличествоАрхивныхДокументов>0</КоличествоАрхивныхДокументов>"& chr(13) & chr(10) _
 &" <ПервоначальнаяВыгрузка>N</ПервоначальнаяВыгрузка>"& chr(13) & chr(10) _
 &" <ПолнаяВыгрузкаСчетов>Нет</ПолнаяВыгрузкаСчетов>"& chr(13) & chr(10) _
 &" <ПолнаяВыгрузкаКлиентов>Нет</ПолнаяВыгрузкаКлиентов>"& chr(13) & chr(10) _
 &"</КонтрольныйЛистСеанса>"& chr(13) & chr(10)

   ' Заводим операциониста
    ts.Write _ 	
 "<Операционист Код="""& filNum &"_ТехОпер"">"& chr(13) & chr(10) _
 &"  <Название>Технический операционист филиала " & filNum & "</Название>"& chr(13) & chr(10) _
 &"  <Филиал>"& filNum &"</Филиал>"& chr(13) & chr(10) _
 &"  <КодВИерархии>E" & filNum & "1000001</КодВИерархии>"& chr(13) & chr(10) _
 &"</Операционист>"& chr(13) & chr(10)

   ts.WriteLine "<СинхронизаторНавигаторов Категория=""ST""/>"
	
	
  ' заводим клиента
   ts.Write _ 	
"<Субъект Тип=""Предп_орг"" Код="""& filNum &"_ТехКлиент"">"& chr(13) & chr(10) _
&"  <Название>Технический клиент филиала   "& filNum &" </Название>"& chr(13) & chr(10) _
&"  <Резидент>Д</Резидент>"& chr(13) & chr(10) _
&"  <Связь_БНК>5</Связь_БНК>"& chr(13) & chr(10) _
&"  <СектЭконом>10</СектЭконом>"& chr(13) & chr(10) _
&"  <ГруппыОбъекта>"& chr(13) & chr(10) _
&"    <ГруппаОбъекта>"& chr(13) & chr(10) _
&"      <Код>Клиент</Код>"& chr(13) & chr(10) _
&"      <Вид_кл>Ю</Вид_кл>"& chr(13) & chr(10) _
&"      <Банк>Н</Банк>"& chr(13) & chr(10) _
&"    </ГруппаОбъекта>"& chr(13) & chr(10) _
&"  </ГруппыОбъекта>"& chr(13) & chr(10) _
&"</Субъект>"& chr(13) & chr(10) 
	
	ts.WriteLine "<СинхронизаторНавигаторов Категория=""SB""/>" & chr(13) & chr(10) & chr(13) & chr(10) & chr(13) & chr(10) 
	
	'Счета конверсии
	dim sql,result 
	sql = "select * from dacleval where ev_type = 13 and ev_itcode like '[0123456789][0123456789][0123456789]'"
	Set result = con.Execute(sql)
	
	If Not result.EOF  Then
	  result.MoveFirst
	  While Not result.EOF
	  	dim cod,naim
	  	  cod = result.Fields("ev_itcode").Value 
	  	  naim = result.Fields("ev_name").Value 
	  	  
	  	  ts.WriteLine GenegateConvAcc(cod,naim,"А",filNum,filDateStart)
	  	  ts.WriteLine GenegateConvAcc(cod,naim,"Б",filNum,filDateStart)
	  	  ts.WriteLine GenegateConvAcc(cod,naim,"В",filNum,filDateStart)
	  	  ts.WriteLine GenegateConvAcc(cod,naim,"Г",filNum,filDateStart)
	  	  if cod = "810" then
	  	    ts.WriteLine GenegateConvAcc(cod,naim,"Д",filNum,filDateStart)
	  	  end if
	   result.movenext
	  wend
	end if
	
	
	ts.WriteLine "<СинхронизаторНавигаторов Категория=""FA""/>"
	

   ts.Write _
"<ЗавершитьСеанс>"& chr(13) & chr(10) _
&"  <Быстро>1</Быстро>"& chr(13) & chr(10) _
&"</ЗавершитьСеанс>"& chr(13) & chr(10) _
&"</Сеанс>"& chr(13) & chr(10)

   
	
	ts.Close
	
	
function GenegateConvAcc (val,val_n, gl, fil, dat_open)
	dim tp 
	tp = "И"
	if val = "810" then 
	  tp = "Р"
	end if
dim str
  str = _
"<ЛицевойСчет Номер=""00000/"& val & "_" & gl& """>"& chr(13) & chr(10) _
&"  <Филиал>"& fil &"</Филиал>"& chr(13) & chr(10) _
&"  <Операционист>"& fil &"_ТехОпер</Операционист>"& chr(13) & chr(10) _
&"  <Клиент>"& fil &"_ТехКлиент</Клиент>"& chr(13) & chr(10) _
&"  <Валюта>"& val &"</Валюта>"& chr(13) & chr(10) _
&"  <Операция>Неопр_опер</Операция>"& chr(13) & chr(10) _
&"  <Название>"& val_n & "</Название>"& chr(13) & chr(10) _
&"  <БалансовыйСтатус>AП</БалансовыйСтатус>"& chr(13) & chr(10) _
&"  <Тип>ЛС</Тип>"& chr(13) & chr(10) _
&"  <ОперацииОткрытия>"& chr(13) & chr(10) _
&"    <ОперацияОткрытия>"& chr(13) & chr(10) _
&"      <ДатаОткрытия>"& dat_open & "</ДатаОткрытия>"& chr(13) & chr(10) _
&"    </ОперацияОткрытия>"& chr(13) & chr(10) _
&"  </ОперацииОткрытия>"& chr(13) & chr(10) _
&"  <ОтнесенияСчета Количество=""1"">"& chr(13) & chr(10) _
&"   <ОтнесениеСчета>"& chr(13) & chr(10) _
&"      <ПланСчетов>ЦБ_РФ_"& gl & "</ПланСчетов>"& chr(13) & chr(10) _
&"      <СинтетическийСчет>00000_"& tp &"</СинтетическийСчет>"& chr(13) & chr(10) _
&"    </ОтнесениеСчета>"& chr(13) & chr(10) _
&"  </ОтнесенияСчета>"& chr(13) & chr(10) _
&"</ЛицевойСчет>"& chr(13) & chr(10)

  GenegateConvAcc = str	
end function