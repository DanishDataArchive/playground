<%
	Response.Expires = 0	
	Dim objComAjax
%>
	<!-- #include virtual="incDbConn.asp" -->
<%
	Set objComAjax = Server.CreateObject("ADODB.Command")
	objComAjax.ActiveConnection = datconstr
	
	If Request("action") & "x" <> "x" Then
		Dim strAction, strReturn
		Dim strCounty, strHerred, strParish
		Dim strYear, strKipNr, strStednavn,strNavn,strFsted,strHsstandst,strErhverv
		Dim strSql, rs, strSql2
		strAction = Request("action")
		
		If strAction = "getHerreds" Then ' herred
			strCounty = Request.Form("county")
			
			strSql = "SELECT DISTINCT ISNULL(Herred,'-') AS Herred FROM kipdata WHERE Amt = '" & strCounty & "' ORDER BY Herred"
			objComAjax.CommandText = strSql
			Set rs = objComAjax.Execute
			While NOT rs.EOF
				strReturn = strReturn & "<option value=""" & Replace(rs("Herred") & "", """", "'") & """>" & rs("Herred") & "</option>"
				rs.MoveNext
			Wend
			rs.Close
			Set rs = Nothing
			
			Response.Write "<option value="""">Vælg</option>" & strReturn
		ElseIf strAction = "getParishes" Then ' sogn
			strHerred = Request("herred")
			strCounty = Request("county")
			
			strSql = "SELECT DISTINCT Enhed AS Sogn FROM kipdata WHERE Amt = '" & strCounty & "' "
			If strHerred & "x" <> "x" Then
				strSql = strSql & " AND Herred = '" & strHerred & "' ORDER BY Enhed"
			End If
			
			objComAjax.CommandText = strSql
			Set rs = objComAjax.Execute
			While NOT rs.EOF
				strReturn = strReturn & "<option value=""" & Replace(rs("Sogn") & "", """", "'") & """>" & rs("Sogn") & "</option>"
				rs.MoveNext
			Wend
			rs.Close
			Set rs = Nothing
			
			Response.Write "<option value="""">Vælg</option>" & strReturn
		ElseIf strAction = "search" Then
          'her starter tilpasningen til de enkelte søgesider.NC
			strCounty = Request("county")
			strHerred = Request("herred")
			strParish = Request("parish")
			strStednavn = Request("stednavn")
			strKipNr = Request("kipnr")
			strNavn = Request("navn")
			strKilde= Request("kilde")
			
			If strCounty = "null" Then
				strCounty = "" ' don't know, but sometimes 'null' is passed from javascript
			End If
			If strHerred = "null" Then
				strHerred = "" ' don't know, but sometimes 'null' is passed from javascript
			End If
			If strParish = "null" Then
				strParish = "" ' don't know, but sometimes 'null' is passed from javascript
			End If
			
			
			'Returnerer resultatet fra kipfolder
			 strSql = "SET ROWCOUNT 0 SELECT  * FROM kipdata "

			 If strHerred <> "" then
			  strSql = strSql & " where herred like '%" & strHerred & "%'"
			  udfyldt = 1
			 End If

			 if  "alle" <> strCounty and 1 = udfyldt THEN
			  strSql = strSql & " AND  amt like '" & strCounty & "'"
			 elseif  "alle" <> strCounty and "X" <> amt  then
			  strSql = strSql & " where amt like '" & strCounty & "'"
			  udfyldt=1
			 End If

			

			 if "" <> strParish and 1=udfyldt  THEN
			  strSql = strSql & " AND enhed like '%" & strParish & "%'"
			 elseif ""<> strParish then
			  strSql = strSql & " where enhed like '%" & strParish & "%'"
			  udfyldt = 1
			End If
			
			'response.write strSql & "<BR>"
			
			' <MP> åben første tabel her
			Dim rs2
			objComAjax.CommandText = strSql
			Set rs2 = objComAjax.Execute
			' nu må strSql gerne genbruges
			' </MP>
			
			' her skal der søges i en anden tabel!!! NC. 
	         strSql = " ; select a.indtastningsnr,a.navn,a.erhverv,a.stilling_i_husstanden,a.stednavn,a.fødested  from " & strCounty
			 strSql = strSql & " a inner join kipdonorer b on a.Indtastningsnr=b.kipnr inner join enhed c on b.sognekode=c.Kode "
			 udfyldt = 0
			 
			 if "" <> strKipNr and 1=udfyldt  THEN
			  strSql = strSql & " AND indtastningsnr like '%" & strKipNr & "%'"
			 elseif ""<> strKipNr then
			  strSql = strSql & " where indtastningsnr like '%" & strKipNr & "%'"
			  udfyldt = 1
			End If

			 if "" <> strStednavn and 1 = udfyldt THEN
			  strSql = strSql & " AND stednavn like '%" & strStednavn & "%'"
			 elseif ""<> strStednavn then
			  strSql = strSql & " where stednavn like '%" & strStednavn & "%'"
			  udfyldt = 1
			 End If

			 if "" <> strnavn and 1 = udfyldt THEN
			  strSql = strSql & " AND a.navn like '%" & strNavn & "%'"
			 elseif ""<> strNavn then
			  strSql = strSql & " where a.navn like '%" & strNavn & "%'"
			  udfyldt = 1
			 End If

			 if "" <> strKilde and 1 = udfyldt  THEN
			  strSql = strSql & " AND kilde like '%" & strKilde & "%'"
			 elseif "" <> strKilde  then
			  strSql = strSql & " where kilde like '%" & strKilde & "%'"
			  udfyldt=1
			 End If
			 
			 

			if "alle" =strCounty and 1=udfyldt THEN
				strSql=strSql
			elseif strCounty="alle" then
				strSql="set rowcount 0 select * from kipdata "
			End If


			strSql = strSql & "  and c.Navn = '" & strParish & "' ORDER BY  kilde,a.navn "

			Response.Write strSql 
			'Response.Flush
			'Response.End
			' ...and then we search
			
						
			'<MP> og her læses anden tabel
			'Open the Recordsets
			objComAjax.CommandText = strSql
			Set rs = objComAjax.Execute
			'</MP>
	'i strReturn samles alle data fra database og sendes ud til html

			strReturn = "<h2>Udtræk fra DDD: " & strCounty & "</h2>"
			
			On Error Resume Next
			
			If RS.EOF Then
				Response.Write "Ingen poster fundet..."
				Response.End
			End If
			rs.MoveFirst
			Count = 0
			
			While Not rs.eof
				Count = Count + 1
				rs.MoveNext
			Wend
			strReturn = strReturn & "<h3>" & Count & " poster fundet</h3>"
			
			rs.MoveFirst
			
			While Not rs.EOF
				 Response.Write "<p></p>"
				' Test af ny udskrift

				 strReturn = strReturn & "Navn: " & rs("navn") & "; "
				 IF rs("erhverv") <>"" THEN
				  strReturn = strReturn & "Erhverv: " & rs("erhverv") &"; "
				 END IF
				 strReturn = strReturn & rs("stilling_i_husstanden") &"; "
				 IF rs("stednavn") <>"" THEN
				  strReturn = strReturn & "Stednavn: " & rs("stednavn") & "; "
				 ELSE
				  strReturn = strReturn & "; "
				 END IF
				 IF rs("fsted") <>"" THEN
				  strReturn = strReturn & "Fødested: " & rs("fsted") & "; "
				 ELSE
				  strReturn = strReturn & "; "
				 END IF
				 strReturn = strReturn & rs("indtastningsnr") & "; "
				 strReturn = strReturn & "<br>"

				 rs.MoveNext
			 Wend

			'Close the recordset
			 rs.Close
			 Set rs = Nothing
			 
			 Response.Write strReturn
			 Response.End
		End If
		Set objComAjax = Nothing
		Response.End ' stop script execution and return and buffered data
	End If
	
	Function CountyGetCollection()
		Dim rs, strReturn
		objComAjax.CommandText = "SELECT distinct amt FROM kipdata ORDER BY Amt"
		Set rs = objComAjax.Execute
		While NOT rs.EOF
			strReturn = strReturn & "<option value=""" & rs("Amt") & """>" & rs("Amt") & "</option>"
			rs.MoveNext
		Wend
		rs.Close
		Set rs = Nothing
		CountyGetCollection = "<option value=""alle"">Vælg</option>" & strReturn
	End Function
	
%>