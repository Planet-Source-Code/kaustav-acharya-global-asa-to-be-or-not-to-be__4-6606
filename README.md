<div align="center">

## Global\.asa: To be or not to be?


</div>

### Description

To educate all coders when or when not to use the 'Global.asa' file and the conditions you must fullfill when you do use it. I knew I had a lot of trouble finding documentation like this. If you like it, please feel free to vote for it. Let me know what I can do to improve myself too, cuz this is my very first tutorial that I wrote. I've been told I'm good at explaining, but let me know ok? :) I'll see if I can dig up any more info on the 'Global.asa' file and try to continue this lesson. The most important reason why I created this tutorial is because most advanced scripts(i.e like my chat script, some database scripts, etc) use it. Very important piece of knowledge to be able to have at your fingertips. Now you can even impress your date with this stuff! ^_^ Not!! Hehe.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Kaustav Acharya](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/kaustav-acharya.md)
**Level**          |Beginner
**User Rating**    |4.2 (63 globes from 15 users)
**Compatibility**  |ASP \(Active Server Pages\), HTML
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__4-33.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/kaustav-acharya-global-asa-to-be-or-not-to-be__4-6606/archive/master.zip)





### Source Code

```
<P>
<FONT size="2"><FONT face="Verdana">Did you use web based
email accounts, shopping bags or online auctions before? If you did , didn't you
ask your self how that web application recognize that this is me when I go from
one page to another inside the site, or how it recognizes what I'm doing and its
me who did that event and not another visitor to the site? How could the site
remember visitors and know if they are still connected or not, or what are they
doing?<BR><BR>If they are using IIS the answer probably lies within the
Global.asa, Global.asa is an optional file that can be used to handle
application and session events. For example with the Global.asa help you can do
some function when a new user come to your site and another function when that
user leaves the site. With the session object you can track users and their
events and much more.<BR><BR><FONT color="#3333cc"><B>Some Facts about Global.asa
:</B></FONT></FONT></FONT><FONT style="LINE-HEIGHT: 1.5em"><BR><BR><FONT size="2"><FONT face="Verdana"><FONT color="#ff0000">¨</FONT>  The file must be
named (Global.asa).<BR><FONT color="#ff0000">¨</FONT>  Global.asa must be
stored in the root directory of the web application.<BR><FONT color="#ff0000">¨</FONT>  Global.asa is processed automatically by the server
when<BR>* The IIS starts and stops.<BR>* Users start and stop sessions that use
the application's web pages.<BR><FONT color="#ff0000">¨</FONT>  The scripts
in the Global.asa are used to:<BR>*Initialize application and session
variables.<BR>*Connect to databases.<BR>*Send cookies.<BR><FONT color="#ff0000">¨</FONT>  In the Global.asa your script must be enclosed with
the <script> tag.<BR><FONT color="#ff0000">¨</FONT>  You can use any
supported scripting language to write your scripts inside Global.asa<BR><FONT color="#ff0000">¨</FONT>  The IIS/PWS have a build in timelimit (defaulted to
20 minutes), so if the user don't request for any thing for that time limit the
session will be ended automatically.<BR><BR>We use global.asa to make use of the
application and session in the ASP, it can contain four scripts to handle
application and session events.<BR><FONT color="#ff0000">¨</FONT> 
Application_OnStart :run when IIS/PWS is started<BR><FONT color="#ff0000">¨</FONT>  Application_OnEnd :when IIS/PWS is
ended<BR>(Usually the application two scripts are run when the server is
rebooted)<BR><BR><FONT color="#ff0000">¨</FONT>  Session_OnStart :run when a
user starts his/hers session<BR><FONT color="#ff0000">¨</FONT>  Session_OnEnd
:when the session runs out<BR>(two scripts that are run when a user starts
his/hers session )<BR><BR><BR></FONT></FONT></FONT><FONT face="Verdana" color="#3333cc" size="2"><B>Things you must know about Global.asa:</B></FONT><FONT style="LINE-HEIGHT: 1.5em"><BR><BR><FONT size="2"><FONT face="Verdana"><FONT color="#ff0000">¨</FONT>  Sessions will only work with visitors that have
cookies enabled.<BR><FONT color="#ff0000">¨</FONT>  If you make a change to
the script. Usually you have to reboot the server to get it refreshed and work
correctly.<BR><BR></FONT></FONT></FONT><FONT face="Verdana" color="#3333cc" size="2"><B>how Global.asa look from inside :</B></FONT><FONT style="LINE-HEIGHT: 1.5em"><BR><BR><FONT face="Verdana" size="2"><SCRIPT
LANGUAGE=VBScript RUNAT=Server><BR>Sub Application_OnStart<BR>'Your
Application_OnStart code will be here<BR>End
Sub<BR></SCRIPT><BR><BR><SCRIPT LANGUAGE=VBScript
RUNAT=Server><BR>Sub Application_OnEnd<BR>'Your Application_OnEnd code will
be here<BR>End Sub<BR></SCRIPT><BR><BR><BR><SCRIPT LANGUAGE=VBScript
RUNAT=Server><BR>Sub Session_OnStart<BR>'Your Session_OnStart code will be
here<BR>End Sub<BR></SCRIPT><BR><BR><BR><SCRIPT LANGUAGE=VBScript
RUNAT=Server><BR>Sub Session_OnEnd<BR>'Your Session_OnEnd code will be
here<BR>End Sub<BR></SCRIPT><BR>
End of Tutorial
```

