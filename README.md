<div align="center">

## MS Access 97 : Undocumented Startup option that repairs and optimizes


</div>

### Description

While searching for a way to repair a corrupted database file, that couldn't be fixed with just Repair and Compact, i found this undocumented option that prooved to repair and optimize my database file. Read the whole article. Any comments or feedback are welcomed.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[John Galanopoulos](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/john-galanopoulos.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VBA MS Access
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/john-galanopoulos-ms-access-97-undocumented-startup-option-that-repairs-and-optimizes__1-31076/archive/master.zip)





### Source Code

```
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>New Page 3</title>
</head>
<body>
<p class=MsoTitle><span lang=EN-US><b><font size="4" color="#0000FF">The undocumented /decompile option</font></b></span></p>
<p class=MsoSubtitle style='text-align:justify;text-indent:36.0pt'><span
lang=EN-US style='font-size:11.0pt;mso-bidi-font-size:12.0pt;color:maroon;
font-weight:normal;text-decoration:none;text-underline:none'>I don’t know how
many of you, are familiar to this, but there is an undocumented option on MS
Access 97 that can be issued not only as a command line option to a database,
that has a small chance of recovery, but also to a database that has many
objects (Tables, forms, reports, source etc)<span style="mso-spacerun: yes">
</span>and works very slowly or has a strange attitude or even memory leaks.<o:p></o:p></span></p>
<p class=MsoSubtitle style='text-align:justify;text-indent:36.0pt'><span
lang=EN-US style='font-size:11.0pt;mso-bidi-font-size:12.0pt;color:maroon;
font-weight:normal;text-decoration:none;text-underline:none'>What this option
seems to do, is a “cold boot” to the database file that we try to fix or
optimize. It reorganizes every object respectively in a way, that results to a
more efficient and </span><span lang=EN-US
style='font-size:11.0pt;mso-bidi-font-size:12.0pt;color:maroon;font-weight:
normal;text-decoration:none;text-underline:none'>Well structured database file.
<o:p></o:p></span></p>
<p class=MsoSubtitle style='text-align:justify'><span lang=EN-US
style='font-size:11.0pt;mso-bidi-font-size:12.0pt;color:maroon;font-weight:
normal;text-decoration:none;text-underline:none'>Take a closer look to the steps we must take so that
this hidden option works at it’s best :<o:p></o:p></span></p>
<p class=MsoSubtitle style='text-align:justify'><span lang=EN-US
style='font-size:11.0pt;mso-bidi-font-size:12.0pt;color:maroon;font-weight:
normal;text-decoration:none;text-underline:none'><![if !supportEmptyParas]> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span><![if !supportLists]><span
lang=EN-US style='font-size:11.0pt;mso-bidi-font-size:12.0pt;color:maroon;
font-weight:normal;text-decoration:none;text-underline:none'>&nbsp;&nbsp;&nbsp;&nbsp;
1.<span
style='font:7.0pt "Times New Roman"'>    </span></span><![endif]><span
lang=EN-US style='font-size:11.0pt;mso-bidi-font-size:12.0pt;color:maroon;
font-weight:normal;text-decoration:none;text-underline:none'>Create a backup of
the original database file.<o:p></o:p></span></p>
<p class=MsoSubtitle style='margin-left:36.0pt;text-align:justify;text-indent:
-18.0pt;mso-list:l0 level1 lfo1;tab-stops:list 36.0pt'><![if !supportLists]><span
lang=EN-US style='font-size:11.0pt;mso-bidi-font-size:12.0pt;color:maroon;
font-weight:normal;text-decoration:none;text-underline:none'>&nbsp;&nbsp;&nbsp;&nbsp;
2.<span
style='font:7.0pt "Times New Roman"'>    </span></span><![endif]><span
lang=EN-US style='font-size:11.0pt;mso-bidi-font-size:12.0pt;color:maroon;
font-weight:normal;text-decoration:none;text-underline:none'>Open the database
file like this <o:p></o:p></span></p>
<p class=MsoSubtitle style='margin-left:36.0pt;text-align:justify'><span lang=EN-US
style='font-size:11.0pt;mso-bidi-font-size:12.0pt;color:maroon;text-decoration:
none;text-underline:none'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
Msaccess.exe /decompile mydb.mdb</span></p>
<p class=MsoSubtitle style='margin-left:36.0pt;text-align:justify'><![if !supportLists]><span
lang=EN-US style='font-size:11.0pt;mso-bidi-font-size:12.0pt;color:maroon;
font-weight:normal;text-decoration:none;text-underline:none'>3.<span
style='font:7.0pt "Times New Roman"'>    </span></span><![endif]><span
lang=EN-US style='font-size:11.0pt;mso-bidi-font-size:12.0pt;color:maroon;
font-weight:normal;text-decoration:none;text-underline:none'>open the database
file, select a module and go to design view<o:p></o:p></span></p>
<p class=MsoSubtitle style='margin-left:36.0pt;text-align:justify;text-indent:
-18.0pt;mso-list:l0 level1 lfo1;tab-stops:list 36.0pt'><![if !supportLists]><i><span
lang=EN-US style='font-size:11.0pt;mso-bidi-font-size:12.0pt;color:maroon;
font-weight:normal;text-decoration:none;text-underline:none'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
4.<span
style='font:7.0pt "Times New Roman"'>   </span></span></i><![endif]><span
lang=EN-US style='font-size:11.0pt;mso-bidi-font-size:12.0pt;color:maroon;
font-weight:normal;text-decoration:none;text-underline:none'>from the menu,
select <i>Compile and save all modules<o:p></o:p></i></span></p>
<p class=MsoSubtitle style='margin-left:36.0pt;text-align:justify;text-indent:
-18.0pt;mso-list:l0 level1 lfo1;tab-stops:list 36.0pt'><![if !supportLists]><i><span
lang=EN-US style='font-size:11.0pt;mso-bidi-font-size:12.0pt;color:maroon;
font-weight:normal;text-decoration:none;text-underline:none'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
5.<span
style='font:7.0pt "Times New Roman"'>   </span></span></i><![endif]><span
lang=EN-US style='font-size:11.0pt;mso-bidi-font-size:12.0pt;color:maroon;
font-weight:normal;text-decoration:none;text-underline:none'>Then compact your
database file..<i><o:p></o:p></i></span></p>
<p class=MsoSubtitle style='margin-left:18.0pt;text-align:justify'><span
lang=EN-US style='font-size:11.0pt;mso-bidi-font-size:12.0pt;color:maroon;
font-weight:normal;text-decoration:none;text-underline:none'><![if !supportEmptyParas]>
.. and that’s it. <o:p></o:p></span></p>
<p class=MsoSubtitle style='text-align:justify;text-indent:18.0pt'><span
lang=EN-US style='font-size:11.0pt;mso-bidi-font-size:12.0pt;color:maroon;
font-weight:normal;text-decoration:none;text-underline:none'><![if !supportEmptyParas]>
In most cases, a
corrupted database file should be fixed. In every case though, mostly in a
slow, heavy-duty database file you ‘ll notice that the file size of the mdb
will be reduced, the memory leaks will also be reduced and finally no strange
behaviour will be noticed.<span style="mso-spacerun: yes"> </span><o:p></o:p></span></p>
<p class=MsoSubtitle style='margin-left:18.0pt;text-align:justify'><span
lang=EN-US style='font-size:11.0pt;mso-bidi-font-size:12.0pt;color:maroon;
font-weight:normal;text-decoration:none;text-underline:none'>More information about how to repair an
access database file, can be found here :<o:p></o:p></span></p>
<h3 style='text-align:justify'><span style='font-size:12.0pt;mso-bidi-font-size:
13.5pt'><![if !supportEmptyParas]> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ACC97: How to Repair a Damaged Jet 3.5 Database (Q279334)</span></h3>
<h3 style='text-align:justify'><span style="font-size: 12.0pt; mso-bidi-font-size: 13.5pt">&nbsp;&nbsp;&nbsp;&nbsp;
</span><span lang=EN-US
style='font-size:11.0pt;mso-bidi-font-size:12.0pt;color:maroon'><a
href="http://support.microsoft.com/default.aspx?scid=kb;EN-US;q279334">http://support.microsoft.com/default.aspx?scid=kb;EN-US;q279334</a></span><span
lang=EN-US style='font-size:11.0pt;mso-bidi-font-size:12.0pt;color:maroon;
font-weight:normal;text-decoration:none;text-underline:none'><o:p></o:p></span></h3>
<p class=MsoSubtitle style='text-align:justify'><span lang=EN-US
style='font-size:11.0pt;mso-bidi-font-size:12.0pt;color:maroon;font-weight:
normal;text-decoration:none;text-underline:none'><![if !supportEmptyParas]> </span></p>
</body>
</html>
```

