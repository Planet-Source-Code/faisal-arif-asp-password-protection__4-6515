<div align="center">

## asp password protection


</div>

### Description

Asp Password Protection for website
 
### More Info
 
name the files accordingly as mentioned in the source code


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Faisal  Arif](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/faisal-arif.md)
**Level**          |Intermediate
**User Rating**    |3.0 (15 globes from 5 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Algorithims](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/algorithims__4-29.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/faisal-arif-asp-password-protection__4-6515/archive/master.zip)





### Source Code

```
password.asp
<%@ LANGUAGE="VBSCRIPT" %>
<!--- This example is a simple login system --->
<HTML>
<HEAD>
   <TITLE>Password.asp</TITLE>
</HEAD>
<BODY>
   <FORM ACTION="engine.asp" NAME="form1" METHOD="post">
     User Name: <INPUT TYPE="text" SIZE=30 NAME="username">
     <P>
     Password: <INPUT TYPE="password" SIZE=30 NAME="password">
     <P>
     <INPUT TYPE="submit" VALUE="Login">
   </FORM>
</BODY>
</HTML>
--------------------------------------------------------------------------------
engine.asp
-------------
<%@ LANGUAGE="VBSCRIPT" %>
<%
   ' Connects and opens the text file
   ' DATA FORMAT IN TEXT FILE= "username<SPACE>password"
   Set MyFileObject=Server.CreateObject("Scripting.FileSystemObject")
   Set MyTextFile=MyFileObject.OpenTextFile(Server.MapPath("\path\path\path\path") & "\passwords.txt")
   ' Scan the text file to determine if the user is legal
   WHILE NOT MyTextFile.AtEndOfStream
      ' If username and password found
      IF MyTextFile.ReadLine = Request.form("username") & " " & Request.form("password") THEN
        ' Close the text file
        MyTextFile.Close
        ' Go to login success page
        Session("GoBack")=Request.ServerVariables("SCRIPT_NAME")
        Response.Redirect "inyougo.asp"
        Response.end
      END IF
   WEND
   ' Close the text file
   MyTextFile.Close
   ' Go to error page if login unsuccessful
   Session("GoBack")=Request.ServerVariables("SCRIPT_NAME")
   Response.Redirect "invalid.asp"
   Response.end
%>
------------
invalid.asp
-----
<%@ LANGUAGE="VBSCRIPT" %>
<!--- This example is a simple login system --->
<HTML>
<HEAD>
<TITLE>invalid.asp</TITLE>
</HEAD>
<BODY>
You have entered an invalid username or password.
<P>
<A HREF="password.asp">Try to log in again</A>
</BODY>
</HTML>
------
inyougo.asp
------
<%@ LANGUAGE="VBSCRIPT" %>
<!--- This example is a simple login system --->
<HTML>
<HEAD>
   <TITLE>In You Go</TITLE>
</HEAD>
<BODY>
   You have now entered the password protected page.
</BODY>
</HTML>
-------
passwords.txt
----
gates bill
burns joe
```

