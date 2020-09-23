<div align="center">

## A few cool ASP Tricks\!


</div>

### Description

This will teach the users a little of the cool ASP tricks. Nothing to do with a database really, but They are still kinda cool just to test with!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Travis Howle](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/travis-howle.md)
**Level**          |Intermediate
**User Rating**    |3.2 (54 globes from 17 users)
**Compatibility**  |ASP \(Active Server Pages\), VbScript \(browser/client side\)

**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__4-33.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/travis-howle-a-few-cool-asp-tricks__4-8573/archive/master.zip)





### Source Code

</font><font color="#5F5F5F" face="Verdana" size="2">Forms are great for collecting and handling data on your website. But how can you use ASP to make the most of your forms? Most webhosts provide standard scripts for form handling, but to make your site truly dynamic, it will not be long before you want forms to do something more than the average generic script offers.<br><br>
In this tutorial, we will look at how we can collect and begin to process data from submitted forms on your website. Let`s start with a basic form (say call the file form.asp) as below with two fields, name and email address:<br><br>
< form method="post" action="formhandler.asp" < br>name="whateveryouwant">< br>
Name: < input type="text" name="name">< br><br>
Email address: < input type="text" name="email"><BR>
< input type="submit" name="Submit" value="Submit"><br>
< /form><br><br>
Now let`s begin to create formhandler.asp. The input from the form field called "name" is retrieved by calling request.form("name") and for the email form field, request.form("email"). Simple?<br><br>
The first thing to consider is, what happens if someone leaves the fields blank. I imagine you don`t want to let them do that, so let`s set up a basic statement to handle that. We will be using If - Then - Else which is talked about in a later tutorial:<br><br>
< %<br>
If request.form("name")="" or request.form<br>("email")="" then<br>
response.redirect "form.asp"<br>
response.end<br>
End if<br>
% ><br><br>
OK, so if either the name or email field were left blank, nothing is processed and the user is redirected back to the form to try again. In a later tutorial we may show you a nice way to highlight the fields that are required.<br><br>
Right lets assume the user had the wit to actually input their name and email address. So what can you do with it? There are various possibilities: you could store the information in a database (more later on that), you could email it to yourself (again more later on this) or you could simply output the data on the webpage....ok that`s not very useful, but for the purposes of this basic tutorial it is what we will do.<br><br>
So if you want to test this, set up your form.asp using the form code in the first codebox and set up formhandler.asp as follows:<br><br>
< %<br>
If request.form("name")="" or request.form<br>("email")="" then<br>
response.redirect "form.asp"<br>
response.end<br>
End if<br><br>
response.write "Name: " & request.form("name") <br>& "<BR>Email: " & request.form("email")<br>
response.end<br>
% ><br><br>
So simply, if both fields are filled in, it will output the Name and Email address. Note how you can combine text and ASP variables in your response.write statement as above.<br><br>
----------------------------------------------------------<br><br>
There are several different ways of displaying dates.<br>
You can also add and subtract from them as well.
I used one (1) in the examples below but, you can use any number you like.<br><br>
Date 10/15/2003<br>
< %=Date%> <br><br>
Long Date Wednesday, October 15, 2003<br>
< %=FormatDateTime(Now(),vbLongDate)%><br><br>
Adding to Long Date Thursday, October 16, 2003<br>
< %=FormatDateTime(Now()+1,vbLongDate)%> <br><br>
Subtracting From Long Date Tuesday, October 14, 2003<br>
< %=FormatDateTime(Now()-1,vbLongDate)%> <br><br>
Short Date 10/15/2003<br>
< %=FormatDateTime(Now(),vbShortDate)%> <br><br>
Adding To Short Date 10/16/2003<br>
< %=FormatDateTime(Now()+1,vbShortDate)%> <br><br>
Subtracting From Short Date Tuesday, October 14, 2003<br>
< %=FormatDateTime(Now()-1,vbShortDate)%><br><br>
General Date 10/15/2003 4:32:27 PM<br>
< %=FormatDateTime(Now(),vbGeneralDate)%><br><br>
Adding To General Date 10/16/2003 4:32:27 PM
< %=FormatDateTime(Now()+1,vbGeneralDate)%> <br><br>
Subtracting From General Date 10/14/2003 4:32:27 PM<br>
< %=FormatDateTime(Now()-1,vbGeneralDate)%><br>
<br>
-----------------------------------------------------<br><br>
To see if a specific file exists on the web server, then show the results to the user use the following code in an ASP page:<br><br>
< %<br>
Set fs=Server.CreateObject<br>("Scripting.FileSystemObject")<br><br>
If (fs.FileExists("c:\inetpub\www\default.asp"))=true Then<br>
 Response.Write("File <br>c:\winnt\cursors\The DEFUALT.asp File exists.")
Else<br>
 Response.Write("File c:\winnt\cursors\The File does not exist.")<br>
End If<br><br>
set fs=nothing<br>
% >
<br><br>
-----------------------------------------------------<br><br>
To check to see if a drive is ready. Such as drive D for a CD copy or something type:
<br><br>
< %<br>
dim fs,d,n<br>
set fs=Server.CreateObject<br>("Scripting.FileSystemObject")<br>
set d=fs.GetDrive("c:")<br>
n = "The " & d.DriveLetter<br>
if d.IsReady=true then <br>
 n = n & " drive is ready."<br>
else<br>
 n = n & " drive is not ready."<br>
end if <br>
Response.Write(n)<br>
set d=nothing<br>
set fs=nothing<br>
% ><br><br><br>
I hope you liked it! And I hope it helped you :D Pleaes remember to vote! lol ;)
</font>

