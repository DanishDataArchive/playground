<%
	Response.Expires = 0
	Dim datconstr
	Dim objCom
	Set objCom = Server.CreateObject("ADODB.Command")
	

%>
	<!-- #include virtual="incDbConn.asp" -->
	<!-- #include file="tekstvalidering.asp" -->
<% 
	objCom.ActiveConnection = datconstr
%>
	<!-- #include file="soegpersonenkelajax.asp" -->


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
   "http://www.w3.org/TR/html4/loose.dtd">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
	<script type="text/javascript" src="jquery1.8.3.min.js"></script>
<meta name="description" content="Dansk Data Arkiv indhenter, opbevarer og udleverer forskningsdata fra samfundsvidenskab, sundhedsvidenskab og historie.">
<link rel="stylesheet" type="text/css" href="style.css">

	<script type="text/javascript" src="soegpersonenkelajax.js"></script>


<title>Dansk Demografisk Database - Find personer i Folket?lingerne</title>
<script type="text/javascript">

  var _gaq = _gaq || [];
  _gaq.push(['_setAccount', 'UA-27693674-1']);
  _gaq.push(['_setDomainName', 'dda.dk']);
  _gaq.push(['_trackPageview']);

  (function() {
    var ga = document.createElement('script'); ga.type = 'text/javascript'; ga.async = true;
    ga.src = ('https:' == document.location.protocol ? 'https://ssl' : 'http://www') + '.google-analytics.com/ga.js';
    var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(ga, s);
  })();

</script>
</head>

<body>

<div align="center" style="border-right-style: solid; border-right-width: 1px; padding-right: 4px">
	<table border="0" width="1100" id="table1" cellspacing="0" cellpadding="0" class="frametable">
		<tr>
			<td valign="top" class="header">
			<p align="center">
			<img border="0" src="gfx/topbanner.gif" width="760" height="109"></td>
		</tr>
		<tr>
			<td valign="top" class="globalnavigation">
			<!--webbot bot="Include" U-Include="../../includes/globalmenu.htm" TAG="BODY" startspan -->

<div align="center">

<table border="0" width="" id="table1" cellspacing="0" cellpadding="0" height="24">
	<tr>
		<td align="center">&nbsp;</td>
		<td align="center">
		<a class="globalnav" target="_blank" href="http://www.dis-danmark.dk/kort/kort.htm">Kort over amt og sogne</a></td>
		<td align="center" width="35">&nbsp;</td>
		<td align="center">
		<a class="globalnav" id="kiplink" href="http://ddd.dda.dk/kiplink1.htm">Folketællinger</a> </td>
		<td align="center" width="35">&nbsp;</td>
		<td align="center">
		<a class="globalnav" id="bestilcd" href="../../bestil_cd.asp">Bestil CD-ROM</a>&nbsp; </td>
		
		<td align="center" width="35"><p align="center">&nbsp;</td>
		<td align="center">
		<a class="globalnav" id="ddd" href="http://ddd.dda.dk/ddd.htm">Andre databaser</a></td>
		<td align="center" width="35"><p align="center">&nbsp;</td>

		<td align="center"> <a class="globalnav" id="ddd" href="http://ddd.dda.dk/kiip/om_kirkebog.asp">Kirkebøger</a></td>
		
	</tr>
</table>

</div>

<!--webbot bot="Include" i-checksum="24362" endspan --></td>
		</tr>
		<tr>
			<td valign="top" mainframe>
			<table border="0" width="100%" id="table2" cellspacing="0" cellpadding="0" class="main">
				<tr>
					<td align="left" valign="top" >
					  <table class="mainleft"> <tr><td width="200" class="mainleftborder" valign="top">
						<!--webbot bot="Include" U-Include="../../includes/blivindtastertip.htm" TAG="BODY" startspan -->

<table border="0" width="100%" id="table1" cellpadding="0" class="news" background="../../gfx/menubg.gif" style="border-collapse: collapse">
	<tr>
		<td>
                    <b><font face="Verdana" size="2">Mangler du et område?</font></b><p>
                    <font face="Verdana" size="2">Der mangler stadig meget at 
					blive indtastet.</font></p></td>
	</tr>
	<tr>
		<td>
                    <font face="Verdana" size="2">Du kan altid selv vælge, hvad 
					du vil indtaste.</font></td>
	</tr>
	<tr>
		<td>
                <font face="Verdana" size="2">Men lige nu på arbejder vi på at 
				få indtastet folketællingerne for 1860 og 1901.</font></td>
	</tr>
	<tr>
		<td><b><font face="Verdana" size="2">Hvor langt er vi kommet</font></b><font face="Verdana" size="2">?<br>
		Du kan følge projektet på den <a href="../../kipoversigt.htm">samlede oversigt</a> 
		eller på <a target="_blank" href="http://www.dis-danmark.dk/kipkort/">
		kortet</a> <br>
		over de forskellige årgange og steder.</font></td>
	</tr>
	<tr>
		<td height="20">&nbsp;</td>
	</tr>
	<tr>
		<td>
					<b><font face="Verdana" size="2">Hvordan kan du bidrage?</font></b><font face="Verdana" size="2"><br>
                    Du kan melde dig til DDA via
					<a href="mailto:mailbox@dda.sa.dk?subject=Vil gerne være indtaster">
					email</a> og melde dig til at indtaste<br>
&nbsp;eller til at læse korrektur på det, der er indtastet.</font></td>
	</tr>
	<tr>
		<td><font face="Verdana" size="2">Du kan også ringe til os på telefon: 
		66 11 30 10</font></td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>
	<tr>
		<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>
	</table>

<!--webbot bot="Include" i-checksum="17620" endspan --></td></tr></table>
					</td>
					<td width="2%">&nbsp;</td> 
					<td width="580" valign="top">
					<div align="center">
						<table border="0" width="900" id="printContent" cellspacing="0" cellpadding="0" align="left">	
						 <tr>
								<td valign="top">
								<h1>Søg efter person</h1>
								

								<table border="0" width="100%" id="table7" cellspacing="0" cellpadding="0">
									<tr>
										<td width="59%">
										<form id = "formkipfolder" name="formkipfolder" method="post">
										
				<p><table width="562" cellspacing="2" cellpadding="1" border="0">
			 
				<tr>
			 		 <td  colspan="6"align ="left">
					 		 <h2>Bopælsoplysninger:</h2>
						</td>
			 </tr>
			 
				<tr>
			 		 <td width="108">
					 		 Amt:
					 </td>
			 		 <td width="151" colspan="3"><select id="ddlCounty" name="county" class='ddlMedium'>
							<%= CountyGetCollection %>
						</select>					
			 </td><td width="289" colspan="2">
									
						
						</td>
			 </tr>
				<tr>
			 		 <td width="108">Herred:</td>
					 <td width="151" colspan="3"><select name="herred" id="ddlHerred">
						</select>
						
</td>
					 <td width="104">Sogn:</td>
					 <td width="182" align ="left"><select name="parish" id="ddlParish">
						</select>	
						
						
</td>
			 </tr>
				<tr>
			 		 <td width="108" >KIPnr:</td>
					 <td width="151" colspan="3" >
						<input type="text" name="kipnr" value ="<%=kipnr%>" class="txt" style="width:80px;" > </td>
					 <td width="104" >Stednavn:</td>
					 <td width="182" align="left" >
						<input type="text" name="stednavn" class="txt" style="width:100;height:19" value = "<%=stednavn%>">
					 </td>
			 </tr>

				<tr>
			 		 <th width="263" colspan="4" align="left">
						<h2>Personoplysninger:</h2>
						</th>
					 <td width="289" colspan="2">&nbsp;</td>
			 	</tr>
				<tr>
			 		 <td width="263" colspan="4">Navn:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						</td>
					 <td width="289" colspan="2">
						<input type="text" name="navn" class="txt" value = "<%=navn%>" size="60%"></td>
			 	</tr>
				<tr>
			 		 <td width="263" colspan="4">FTår</td>
					 <td width="289" colspan="2">
						<select name="kilde" size="1" > 
						  <option value="1">Alle år    </option> 
						  <option value="1769">1769</option> 
						  <option value="1787">1787</option> 
						  <option value="1801">1801</option> 
						  <option value="1803">1803</option> 
						  <option value="1834">1834</option> 
						  <option value="1835">1835</option> 
						  <option value="1840">1840</option> 
						  <option value="1845">1845</option> 
						  <option value="1850">1850</option> 
						  <option value="1855">1855</option> 
						  <option value="1860">1860</option> 
						  <option value="1864">1864</option>
						  <option value="1870">1870</option> 
						  <option value="1880">1880</option> 
						  <option value="1885">1885</option> 
						  <option value="1890">1890</option> 
						  <option value="1901">1901</option> 
						  <option value="1906">1906</option> 
						  <option value="1911">1911</option> 
						  <option value="1916">1916</option> 
						  <option value="1921">1921</option> 
						  <option value="1925">1925</option> 
						  <option value="1930">1930</option> 
						</select></td>
			 	</tr>
				<tr>
			 		 <td width="263" colspan="4">&nbsp;</td>
					 <td width="289" colspan="2">
						&nbsp;</td>
			 	</tr>
				<tr>
			 		 <td width="263" colspan="4">&nbsp;</td>
					 <td width="289" colspan="2">
						&nbsp;</td>
			 	</tr>
				<tr>
			 		 <td width="108">&nbsp;</td>
					 <td width="130">&nbsp;</td>
					 <td width="106">&nbsp;</td>
					 <td width="4">
  					 		&nbsp;</td>
			 	</tr>
				<tr>
			 		 <td width="420" colspan="3" class="butRow">
			 		<input type="reset" value="Nulstil" onclick="$('#searchResults, #ddlHerred, #ddlParish').html('');">
							&nbsp;
							 <input type="button" class="greenButton greenBuText" value="Søg" id="btnSearch" style="width:106px;">
			 		 </td>
			 	</tr>
</table>

											
</p>
										</form>
&nbsp;</td>
										<td width="4">&nbsp;</td>
										<td width="50%">
										&nbsp;</td>
									</tr>
									<tr>
										<td><div id="searchResults"></div></td>
										<td>&nbsp;</td>
										<td>&nbsp;</td>
									</tr>
									</table>
					</td>
					<td width="200" class="mainrightborder" valign="top">
					<table width="100%" id="table6" class="mainright">
						<tr>
							
							<td valign="top">
							<!--webbot bot="Include" U-Include="../../includes/soegpersontip.htm" TAG="BODY" startspan -->

<table border="0" width="100%" id="table1" cellpadding="0" class="news" background="../../gfx/menubg.gif" style="border-collapse: collapse">
	<tr>
		<td>
                    <font face="Verdana" size="2">
                    <b>Generelt om søgningen</b></font><p>
                    <font face="Verdana" size="2">Der skal altid udfyldes mindst tre tegn i navnefeltet.</font></p></td>
	</tr>
	<tr>
		<td>
                    <font face="Verdana" size="2">Fødested er først oplyst fra 1845</font></td>
	</tr>
	<tr>
		<td>
                <font face="Verdana" size="2">Ikke alle indtastninger 
					har oplysninger om køn.<br>
					KIPnr kan erstatte sogn, herred og år</font></td>
	</tr>
	<tr>
		<td><font face="Verdana" size="2"><b>Visning af herreder og sogne</b><br>Når du har valgt et amt, 
		vises felterne for herred og sogn. Der vises kun betegnelser for steder, 
		der er indtastet.</font></td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>
	<tr>
		<td>
					<font face="Verdana" size="2">
					<b>Usikker på 
					stavemåde?</b><br>
                    Du kan søge med præcis stavemåde eller du kan anvende 
		'joker' tegn, En _&nbsp; erstatter et tegn og en % erstatter flere tegn.</font></td>
	</tr>
	<tr>
		<td><font face="Verdana" size="2">Du kan søge med indeholder, begynder med eller lig med ved søgning 
		på navn</font></td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>
	<tr>
		<td><font face="Verdana" size="2"><b>Har du fundet en fejl?</b></font></td>
	</tr>
	<tr>
		<td><font face="Verdana" size="2">Hvis du finder fejl i en indtastning, 
		der er læst korrektur på, kan du udfylde et
		<a target="_blank" href="http://ddd.dda.dk/fejlmelding.asp">
		fejlmeldingsskema</a>.&nbsp; <br>
		Fejl bliver ikke rettet med det samme, men alle jeres indberetninger 
		bliver gemt og vil blive brugt.</font></td>
	</tr>
	<tr>
		<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>
	</table>

<!--webbot bot="Include" i-checksum="55393" endspan --></td>
						</tr>
					</table>
					</td>
				</tr>
			</table>
			</td>
		</tr>
	
		<tr>
			<td valign="top">&nbsp;</td>
		</tr>
		<tr>
			<td valign="top"></td>
		</tr>
	</table>
</div>

</body>

</html>
<%
	Set objCom = Nothing
%>