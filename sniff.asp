<%  
	'=== define where comlist.xml will reside.  Directory must allow read/write.	
	comXML = "comlist.xml"
	thisPath = Request.ServerVariables("PATH_TRANSLATED")
	comXML = Replace(thisPath,"sniff.asp",comXML)
	
	Set FSO = Server.CreateObject("Scripting.FileSystemObject") 
	
	
	action  = Request.Form("action")
	If action = "ADD" Then
		Call addCOM()
	ElseIF action = "UPDATE" Then
		Call deleteCOM()
	End If
	
	Sub addCOM
		'=== add a component to the XML file
		company = TRIM(Request.Form("company"))
		name = TRIM(Request.Form("name"))
		id = TRIM(Request.Form("id"))
		
		Set oXML = Server.CreateObject("Microsoft.XMLDOM")
		If NOT FSO.FileExists(comXML) Then 
			'=== create new XML file
			Set oRoot = oXML.CreateElement("comlist")
			oXML.appendChild oRoot	
			Set oPI = oXML.createProcessingInstruction("xml","version='1.0'")
			oXML.insertBefore oPI, oXML.childNodes(0)
			oXML.Save comXML
		Else
			'=== load exisiting XML file
			oXML.async = FALSE
			oXML.load(comXML)
		End If
		
		Set oCom = oXML.createElement("com")
			
		Set oAttribCompany = oXML.createAttribute("company")
		oAttribCompany.text = company
		oCom.Attributes.setNamedItem oAttribCompany
		
		Set oAttribName = oXML.createAttribute("name")
		oAttribName.text = name
		oCom.Attributes.setNamedItem oAttribName
		
		Set oAttribID = oXML.createAttribute("id")
		oAttribID.text = id
		oCom.Attributes.setNamedItem oAttribID
		
		Set oRoot = oXML.documentElement
		oRoot.appendChild oCom
		oXML.Save comXML
		Set oXML = nothing
	End Sub
	
	Sub deleteCOM
		'=== remove a component from the XML file
		Set oXML = Server.CreateObject("Microsoft.XMLDOM")
		oXML.async = FALSE
		oXML.load(comXML)
		com = Request.Form("com")
		Set node = oXML.documentElement.selectSingleNode("com[@id='" & com & "']")
		node.parentNode.removeChild(node)
		oXML.Save comXML
		Set node = nothing
		Set oXML = nothing
		
	End Sub
	
	Sub getCOMList
		'=== open XML file and retrive Component list to test		
		If FSO.FileExists(comXML) Then 
			Set oXML = Server.CreateObject("Microsoft.XMLDOM")
			oXML.async = FALSE
			oXML.load(comXML)	
			
			Set nodeList = oXML.documentElement.selectNodes("com") 
			nodeNum = 0
			For Each node In nodeList 
				thisCompany = node.getAttribute("company")
				thisName = node.getAttribute("name")
				thisID = node.getAttribute("id")
				If testCom(thisID) Then
					Response.Write "<tr class=""yes"">"
					thisTest = "YES"
				Else
					Response.Write "<tr class=""no"">"
					thisTest = "NO"
				End If
				Response.Write "<td>" & thisCompany & "</td><td>" & thisName & "</td><td>" & thisID & "</td><td>" & thisTest & "</td>"
				'Response.Write "<td><a href=""#"" onclick=""whackCOM('" & thisID & "')""><font color=""#0000FF"">DELETE</font></a></td>"
				Response.Write "<td></td>"
				Response.Write "</tr>"
				nodeNum = nodeNum + 1	
			Next
			
			Set oXML = nothing
		Else
			Response.Write "<tr><td colspan=""5"">-- no components on list --</td></tr>"
		End If
		
	End Sub
	
	
	Function testCOM(comName)
		On error resume next
		Set xxx = Server.CreateObject(comName)
		If err.number <> 0 Then
			testCOM = FALSE
		Else
			Set xxx = nothing
			testCOM = TRUE
		End If
		
	End Function
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
	<title>COM Sniffer</title>
	<META name="description" content="Check for components installed on IIS web server.  Screen allows for user to add to component list.">
	<META name="keywords" content="component, COM, sniffer, detect, 3rd party, XML">
	<META name="author" content="mas">
	<META HTTP-EQUIV="Content-Language" CONTENT="EN">
	<META NAME="robots" CONTENT="FOLLOW,INDEX">
	<link rel="STYLESHEET" type="text/css" href="com.css">
	<script language="JavaScript" type="text/javascript">
	<!-- 
		function whackCOM(com){
			if(confirm("Are you sure you want to delete the "+com+" from list?")){
				formUpdate.com.value = com;
				formUpdate.submit();
			}
		}
	// -->
	</script>
</head>

<body>

<div align="right"><a href="sniff.asp"><strong>[Refresh]</strong></a></div>
<h2 align="center">COM Informant - Detect Installed Components</h2>

<form action="sniff.asp" method="post" name="formAdd" id="formAdd">
<input type="hidden" name="action" value="ADD">
<table cellspacing="0" cellpadding="5" align="center" bgcolor="#FFFFFF">
<tr><th colspan="4">Add Component To List</th></tr>
<tr>
	<td>Company: <input type="text" name="company" size="20" maxlength="30"></td>
	<td>Name: <input type="text" name="name" size="20" maxlength="30"></td>
	<td>COM: <input type="text" name="id" size="20" maxlength="30"></td>
	<td><input type="submit" name="add" value="Add" disabled></td>
</tr>
</table>
</form>
<h2 align="center"><%= Request.Servervariables("SERVER_NAME") %> - <%= Now() %></h2>
<form action="sniff.asp" method="post" name="formUpdate" id="formUpdate">
<input type="hidden" name="action" value="UPDATE">
<input type="hidden" name="com" value="">
<table cellspacing="0" cellpadding="3" align="center" bgcolor="#FFFFFF">
<tr><th>Company</th><th>Name</th><th>COM</th><th>Installed?</th><th>Update</th></tr>
<%  
	Call getCOMList()
	Set FSO = nothing
%>
</table>
</form>
<p>&nbsp;</p>
</body>
</html>
