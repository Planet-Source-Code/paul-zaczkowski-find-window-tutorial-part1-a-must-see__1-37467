<div align="center">

## Find Window Tutorial Part1 \<A must see\>


</div>

### Description

This tutorial will walk you through the steps of basic FindWindow functions: FindWindow, FindWindowEx, ShowWindow, DestroyWindow, and GetClassName. Please vote, I spent 2 days on this, and for a 14 year old, that is a lot! Thank you!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2002-07-31 13:52:58
**By**             |[Paul Zaczkowski](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/paul-zaczkowski.md)
**Level**          |Intermediate
**User Rating**    |4.6 (32 globes from 7 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Find\_Windo1130317312002\.zip](https://github.com/Planet-Source-Code/paul-zaczkowski-find-window-tutorial-part1-a-must-see__1-37467/archive/master.zip)





### Source Code

<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=windows-1252">
<META NAME="Generator" CONTENT="Internet Assistant for Word Version 3.0">
</HEAD>
<BODY>
<B><U><FONT FACE="Arial Black" SIZE=5><P ALIGN="CENTER">FINDWINDOW TUTORIAL PART 1</P>
</U></FONT><FONT FACE="Courier New" SIZE=4><P ALIGN="CENTER"></P>
</B></FONT><FONT FACE="Arial"><P>&#9;If you' re serious about Windows Programming, and you' re ready to move to the next step, advanced API, this tutorial is for you. Whether you want to control another program, or you want to see if a program is open, FindWindow (and friends) will do the job for you. This tutorial will walk you through the simple process of finding any window (child or parent) within Windows. So read on!</P>
<P>&#9;First off, we will use the Following API' s, throughout the tutorial:</P>
</FONT><FONT FACE="Courier New">
</FONT><B><FONT FACE="Courier New" SIZE=2><P>1. </B></FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">Declare Function</FONT><FONT FACE="Courier New" SIZE=2> FindWindow </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">Lib</FONT><FONT FACE="Courier New" SIZE=2> "user32" </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">Alias</FONT><FONT FACE="Courier New" SIZE=2> "FindWindowA" (</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">ByVal</FONT><FONT FACE="Courier New" SIZE=2> lpClassName </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">As String</FONT><FONT FACE="Courier New" SIZE=2>, </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">ByVal </FONT><FONT FACE="Courier New" SIZE=2>lpWindowName </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">As String</FONT><FONT FACE="Courier New" SIZE=2>) </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">As Long</P>
</FONT><FONT FACE="Courier New" SIZE=2><P>&nbsp;</P>
<B><P>2. </B></FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">Declare Function </FONT><FONT FACE="Courier New" SIZE=2>FindWindowEx </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">Lib </FONT><FONT FACE="Courier New" SIZE=2>"user32" </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">Alias      </FONT><FONT FACE="Courier New" SIZE=2>"FindWindowExA" (</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">ByVal </FONT><FONT FACE="Courier New" SIZE=2>hWnd1 </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">As Long</FONT><FONT FACE="Courier New" SIZE=2>, </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">ByVal </FONT><FONT FACE="Courier New" SIZE=2>hWnd2 </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">As Long</FONT><FONT FACE="Courier New" SIZE=2>, </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">ByVal </FONT><FONT FACE="Courier New" SIZE=2>lpsz1   </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">As String</FONT><FONT FACE="Courier New" SIZE=2>, </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">ByVal </FONT><FONT FACE="Courier New" SIZE=2>lpsz2 </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">As String</FONT><FONT FACE="Courier New" SIZE=2>) </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">As Long</P>
</FONT><B><FONT FACE="Courier New" SIZE=2><P>3. </B></FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">Declare Function </FONT><FONT FACE="Courier New" SIZE=2>GetClassName </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">Lib </FONT><FONT FACE="Courier New" SIZE=2>"user32" </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">Alias </FONT><FONT FACE="Courier New" SIZE=2>"GetClassNameA" (</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">ByVal </FONT><FONT FACE="Courier New" SIZE=2>hwnd </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">As Long</FONT><FONT FACE="Courier New" SIZE=2>, </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">ByVal </FONT><FONT FACE="Courier New" SIZE=2>lpClassName </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">As String</FONT><FONT FACE="Courier New" SIZE=2>, </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">ByVal </FONT><FONT FACE="Courier New" SIZE=2>nMaxCount </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">As Long</FONT><FONT FACE="Courier New" SIZE=2>)</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"> As Long</P>
</FONT><FONT FACE="Courier New" SIZE=2>
<B><P>4. </B></FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">Declare Function </FONT><FONT FACE="Courier New" SIZE=2>ShowWindow </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">Lib </FONT><FONT FACE="Courier New" SIZE=2>"user32" </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">Alias </FONT><FONT FACE="Courier New" SIZE=2>"ShowWindow" (</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">ByVal </FONT><FONT FACE="Courier New" SIZE=2>hwnd </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">As Long</FONT><FONT FACE="Courier New" SIZE=2>, </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">ByVal </FONT><FONT FACE="Courier New" SIZE=2>nCmdShow </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">As Long</FONT><FONT FACE="Courier New" SIZE=2>) </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">As Long</P>
</FONT><B><FONT FACE="Courier New" SIZE=2><P>5. </B></FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">Declare Function </FONT><FONT FACE="Courier New" SIZE=2>DestroyWindow </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">Lib </FONT><FONT FACE="Courier New" SIZE=2>"user32" </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">Alias </FONT><FONT FACE="Courier New" SIZE=2>"DestroyWindow" (</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">ByVal </FONT><FONT FACE="Courier New" SIZE=2>hwnd </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">As Long</FONT><FONT FACE="Courier New" SIZE=2>) </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">As Long</P>
</FONT><FONT FACE="Arial" COLOR="#00007f">
<P>&#9;</FONT><FONT FACE="Arial">You may be asking yourself, “what in the world do these mean?”, and, if you are, I have the answer. Let' s look at Number 1 first. FindWindow is used to find the handle of any parent window using its classname or its WindowName. The context is below:</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00">
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>Declare Function </FONT><FONT FACE="Courier New" SIZE=2>FindWindow </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">Lib </FONT><FONT FACE="Courier New" SIZE=2>“user32” </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">Alias </FONT><FONT FACE="Courier New" SIZE=2>“FindWindowA</FONT><FONT FACE="Arial">”</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"><P>' Makes a simple function to call the FindWindow routine within </P>
<P>' User32.dll.</P>
</FONT><FONT FACE="Courier New" SIZE=2><P>(</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">ByVal</FONT><FONT FACE="Courier New" SIZE=2> lpClassName </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">As String,</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"><P>' The classname of a window is a STRING value that holds the Windows</P>
<P>' Internal name of your application. Set to vbNullString if you don' t </P>
<P>' know it.</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">
<P>ByVal </FONT><FONT FACE="Courier New" SIZE=2>lpWindowName </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">As String</FONT><FONT FACE="Courier New" SIZE=2>)</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"><P>' The windowname of a window is a STRING value that holds the CAPTION of ' your application. (i.e. If you set the caption of a form to “hello” ' then the WindowName would be: “hello”. Set to vbNullString if you </P>
<P>' dont' know it.</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>As Long</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"><P>' The return value of this function is a long value which holds the</P>
<P>' WINDOWS HANDLE, only if the window exists. If the window doesn' t </P>
<P>' exist, it holds 0. (The hWnd property of a form, is the WINDOWS </P>
<P>' HANDLE.</P>
</FONT><B><FONT FACE="Arial" SIZE=2 COLOR="#ff0000"><P>NOTE: You MUST fill in either ClassName or WindowName in order to find what you are looking for.</P>
</B></FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00">
<P>' Below is an example of FindWindow being used in a real program:</P>
<P> </P>
<P>' Declare the FindWindow API.</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>Private Declare Function </FONT><FONT FACE="Courier New" SIZE=2>FindWindow </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">Lib </FONT><FONT FACE="Courier New" SIZE=2>"user32" </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">Alias </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#003300">_</P>
</FONT><FONT FACE="Courier New" SIZE=2><P>"FindWindowA" </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#003300">(</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">ByVal </FONT><FONT FACE="Courier New" SIZE=2>lpClassName</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#003300"> </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">As String</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#003300">, </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">ByVal </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#003300">_</P>
</FONT><FONT FACE="Courier New" SIZE=2><P>lpWindowName</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#003300"> </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">As String</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#003300">) </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">As Long</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#003300">
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>Private Sub </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#003300">cmdFind_Click()</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"><P>' Run notepad</P>
</FONT><FONT FACE="Courier New" SIZE=2><P>Shell "notepad.exe", vbNormalNoFocus</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"><P>' Search for the window.</P>
</FONT><FONT FACE="Courier New" SIZE=2><P>mWnd = FindWindow("notepad", vbNullString)</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#003300">
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"><P>' NOTE: Also could have used:</P>
<P>' mWnd = FindWindow(vbNullString, "Untitled - Notepad")</P>
<P>' But, since we know the classname, it is better</P>
<P>' programming habit.</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#003300">
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"><P>' If we go the handle display: YAY!</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>If </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00">mWnd &lt;&gt; 0 </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">Then</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#003300"><P>  MsgBox "YAY!"</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"><P>' If we didn' t get the handle Display: darn</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>Else</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#003300"><P>  </FONT><FONT FACE="Courier New" SIZE=2>MsgBox "DARN!"</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>End If</P>
<P>End Sub</P>
</FONT><B><FONT FACE="Arial" SIZE=2 COLOR="#ff0000"><P>WARNING: This program accesses the Windows API. This likely will not crash your system, but may under extreme circumstances.</P>
</B></FONT><FONT FACE="Arial"><P>&#9;You see, it isn' t so hard is it? Now that you have a grasp on FindWindow, we can move on to FindWindowEx. Although the names of the functions may be similar, FindWindow and FindWindowEx do completely different things. You will likely use FindWindowEx in conjunction with FindWindow, because FindWindowEx is used to find a child window of a parent window.  First let' s look at the context of FindWindowEx. It is below:</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>Declare Function </FONT><FONT FACE="Courier New" SIZE=2>FindWindowEx </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">Lib </FONT><FONT FACE="Courier New" SIZE=2>"user32" </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">Alias      </FONT><FONT FACE="Courier New" SIZE=2>"FindWindowExA" </P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"><P>' Makes a simple function to call the FindWindow routine within </P>
<P>' User32.dll.</P>
</FONT><FONT FACE="Arial">
</FONT><FONT FACE="Courier New" SIZE=2><P>(</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">ByVal </FONT><FONT FACE="Courier New" SIZE=2>hWnd1 </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">As Long</FONT><FONT FACE="Courier New" SIZE=2>, </P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"><P>' This calls for the HANDLE of the parent window. Enter the parent </P>
<P>' window' s handle to search it' s child windows.</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>ByVal </FONT><FONT FACE="Courier New" SIZE=2>hWnd2 </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">As Long</FONT><FONT FACE="Courier New" SIZE=2>,</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"><P>' This identifies what child you will start searching from. This takes ' the classname of the child window. Set to 0 to search all of the </P>
<P>' child windows.</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>ByVal </FONT><FONT FACE="Courier New" SIZE=2>lpsz1 </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">As String</FONT><FONT FACE="Courier New" SIZE=2>,</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"><P>' This is your search term. Enter the classname of the window you want ' to search. NOTE: If this window appears BEFORE hWnd2, then you will </P>
<P>' not be able to find this window.</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>ByVal </FONT><FONT FACE="Courier New" SIZE=2>lpsz2 </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">As String</FONT><FONT FACE="Courier New" SIZE=2>)</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"><P>' I know that this is a FindWindow tutorial, and I should know what</P>
<P>' this does. If you know what it does, please comment. Sorry for the</P>
<P>' inconvenience. SET THIS TO: vbNullString!</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>As Long</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"><P>' The function returns a long value, which hold the handle of the child ' window specified.</P>
</FONT><B><FONT FACE="Arial" SIZE=2 COLOR="#0000ff"><P>If you set both hWnd1 and hWnd2 to vbNullString, you will begin searching for parent windows, instead of child windows.</P>
</B></FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"><P>' Below is an example of FindWindowEx being used in a real program:</P>
<P>' Declare the FindWindowEx API.</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>Private Declare Function </FONT><FONT FACE="Courier New" SIZE=2>FindWindowEx </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">Lib </FONT><FONT FACE="Courier New" SIZE=2>"user32"</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"> </FONT><FONT FACE="Courier New" SIZE=2>_</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>Alias </FONT><FONT FACE="Courier New" SIZE=2>"FindWindowExA" </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00">(</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">ByVal </FONT><FONT FACE="Courier New" SIZE=2>hWnd1 </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">As Long</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00">, </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">ByVal </FONT><FONT FACE="Courier New" SIZE=2>hWnd2 _</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>As Long</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00">, </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">ByVal </FONT><FONT FACE="Courier New" SIZE=2>lpsz1 </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">As String</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00">, </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">ByVal </FONT><FONT FACE="Courier New" SIZE=2>lpsz2 </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">As String</FONT><FONT FACE="Courier New" SIZE=2>) _</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>As Long</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00">
<P>' Declare the FindWindow API.</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>Private Declare Function </FONT><FONT FACE="Courier New" SIZE=2>FindWindow </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">Lib </FONT><FONT FACE="Courier New" SIZE=2>"user32" </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">Alias </FONT><FONT FACE="Courier New" SIZE=2>_</P>
<P>"FindWindowA" </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00">(</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">ByVal </FONT><FONT FACE="Courier New" SIZE=2>lpClassName </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">As String</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00">, </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">ByVal </FONT><FONT FACE="Courier New" SIZE=2>_</P>
<P>lpWindowName</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"> </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">As String</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00">) </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">As Long</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00">
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>Private Sub </FONT><FONT FACE="Courier New" SIZE=2>cmdFindWindows_Click()</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"><P>' Search for the TaskBar window.</P>
</FONT><FONT FACE="Courier New" SIZE=2><P>pWnd = FindWindow("shell_traywnd", vbNullString)</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00">
<P>' If the window exists then...</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>If </FONT><FONT FACE="Courier New" SIZE=2>pWnd &lt;&gt; 0 </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">Then</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"><P>  ' Search for the START button.</P>
<P>  </FONT><FONT FACE="Courier New" SIZE=2>cWnd = FindWindowEx(pWnd, 0, "BUTTON", vbNullString)</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"><P>  ' If the window exits then...</P>
<P>  </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">If </FONT><FONT FACE="Courier New" SIZE=2>cWnd &lt;&gt; 0</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"> </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">Then</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"><P>    ' Display The StartButton was found!</P>
<P>    </FONT><FONT FACE="Courier New" SIZE=2>MsgBox "The StartButton was found!"</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"><P>  </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">Else</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"><P>    ' Display The StartButton wasn' t found!</P>
<P>    </FONT><FONT FACE="Courier New" SIZE=2>MsgBox "The StartButton wasn' t found!"</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>  End If</P>
<P>End If</P>
</FONT><B><FONT FACE="Arial" SIZE=2 COLOR="#ff0000"><P>WARNING: This program accesses the Windows API. This likely will not crash your system, but may under extreme circumstances.</P>
</B></FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00">
</FONT><FONT FACE="Arial"><P>&#9;The FindWindowEx is a bit more complicated than FindWindow, but is still pretty simple. Now we are ready to move on to GetClassName. GetClassName is a very simple function, and is also very helpful. If all you know about a window is its WindowName (caption), you can use GetClassName to get the ClassName of the window by sending its HANDLE. The context for GetClassName is below:</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>Declare Function</FONT><FONT FACE="Courier New" SIZE=2> GetClassName </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">Lib </FONT><FONT FACE="Courier New" SIZE=2>"user32" </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">Alias </FONT><FONT FACE="Courier New" SIZE=2>"GetClassNameA"</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"><P>' Makes a simple function to call the GetClassName routine within </P>
<P>' User32.dll.</P>
</FONT><FONT FACE="Courier New" SIZE=2><P>(</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">ByVal </FONT><FONT FACE="Courier New" SIZE=2>hwnd </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">As Long</FONT><FONT FACE="Courier New" SIZE=2>,</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"><P>' This calls for the handle of the window you want to get the ClassName ' of.</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>ByVal </FONT><FONT FACE="Courier New" SIZE=2>lpClassName </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">As String</FONT><FONT FACE="Courier New" SIZE=2>,</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"><P>' Points to the buffer that is to receive the ClassName string. You </P>
<P>' will see how this is used in the example.</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>ByVal </FONT><FONT FACE="Courier New" SIZE=2>nMaxCount </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">As Long</FONT><FONT FACE="Courier New" SIZE=2>)</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"><P>' This is the maximum number of characters that you will allow for the </P>
<P>' ClassName to be. If the ClassName is longer, it will be truncated.</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>As Long</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"><P>' This is a bit misleading. The return value can be a long value if </P>
<P>' called incorrectly, but if called correctly, the return value will be ' a string.</P>
<P>&nbsp;</P>
<P>' Below is an example of GetClassName being used in a real program:</P>
<P>' Declare the FindWindow API.</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>Private Declare Function </FONT><FONT FACE="Courier New" SIZE=2>FindWindow </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">Lib </FONT><FONT FACE="Courier New" SIZE=2>"user32" </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">Alias </FONT><FONT FACE="Courier New" SIZE=2>_</P>
<P>"FindWindowA" (</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">ByVal </FONT><FONT FACE="Courier New" SIZE=2>lpClassName </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">As String</FONT><FONT FACE="Courier New" SIZE=2>, </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">ByVal </FONT><FONT FACE="Courier New" SIZE=2>_</P>
<P>lpWindowName </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">As String</FONT><FONT FACE="Courier New" SIZE=2>) </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">As Long</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00">
<P>' Declare the GetClassName API.</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>Private Declare Function</FONT><FONT FACE="Courier New" SIZE=2> GetClassName </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">Lib </FONT><FONT FACE="Courier New" SIZE=2>"user32" </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">Alias </FONT><FONT FACE="Courier New" SIZE=2>_</P>
<P>"GetClassNameA" (</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">ByVal </FONT><FONT FACE="Courier New" SIZE=2>hwnd </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">As Long</FONT><FONT FACE="Courier New" SIZE=2>, </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">ByVal </FONT><FONT FACE="Courier New" SIZE=2>lpClassName _</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>As String</FONT><FONT FACE="Courier New" SIZE=2>, </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">ByVal </FONT><FONT FACE="Courier New" SIZE=2>nMaxCount </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">As Long</FONT><FONT FACE="Courier New" SIZE=2>) </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">As Long</P>
</FONT><FONT FACE="Courier New" SIZE=2>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"><P>' Declare an empty buffer.</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>Dim </FONT><FONT FACE="Courier New" SIZE=2>ClassNameBuffer </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">As String</P>
</FONT><FONT FACE="Courier New" SIZE=2>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>Private Sub </FONT><FONT FACE="Courier New" SIZE=2>cmdGetClassName_Click()</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"><P>' Make the buffer an empty string with 255 spaces.</P>
</FONT><FONT FACE="Courier New" SIZE=2><P>ClassNameBuffer = String(255, " ")</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"><P>' Run notepad</P>
</FONT><FONT FACE="Courier New" SIZE=2><P>Shell "notepad.exe"</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"><P>' Search for the window.</P>
</FONT><FONT FACE="Courier New" SIZE=2><P>pWnd = FindWindow(vbNullString, "Untitled - Notepad")</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"><P>' If the window exists then...</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>If </FONT><FONT FACE="Courier New" SIZE=2>pWnd &lt;&gt; 0 </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">Then</P>
</FONT><FONT FACE="Courier New" SIZE=2><P>  </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00">' Get the classname</P>
</FONT><FONT FACE="Courier New" SIZE=2><P>  GetClassName pWnd, ClassNameBuffer, 255</P>
<P>  </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00">' Display the ClassName</P>
</FONT><FONT FACE="Courier New" SIZE=2><P>  MsgBox Trim(ClassNameBuffer)</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>End If</P>
<P>End Sub</P>
</FONT><B><FONT FACE="Arial" SIZE=2 COLOR="#ff0000"><P>WARNING: This program accesses the Windows API. This likely will not crash your system, but may under extreme circumstances.</P>
</B></FONT><FONT FACE="Arial"><P>&#9;As you can see, GetClassName can be very helpful, but it isn' t at all complicated. Just like the next function, ShowWindow. This has to be the second easiest of the 5. The context to ShowWindow is shown below:</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>Declare Function </FONT><FONT FACE="Courier New" SIZE=2>ShowWindow </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">Lib </FONT><FONT FACE="Courier New" SIZE=2>"user32" </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">Alias </FONT><FONT FACE="Courier New" SIZE=2>"ShowWindow"</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"><P>' Makes a simple function to call the ShowWindow routine within </P>
<P>' User32.dll.</P>
</FONT><FONT FACE="Courier New" SIZE=2><P>(</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">ByVal </FONT><FONT FACE="Courier New" SIZE=2>hwnd </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">As Long,</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"><P>' Calls for the HANDLE of the window in which you want to edit.</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>ByVal </FONT><FONT FACE="Courier New" SIZE=2>nCmdShow </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">As Long</FONT><FONT FACE="Courier New" SIZE=2>)</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"><P>' Calls for he command that you want to send to the window. The </P>
<P>' commands are listed after the context.</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">
<P>As Long</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"><P>' Although the function doesn' t return anything, it is declared as long.</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#7f0000"><P>COMMANDS:</P>
<P>0 - Hide the Window</P>
<P>1 - Show the Window normally</P>
<P>3 - Maximize</P>
<P>4 - Shows the window without giving it the focus</P>
<P>5 - Shows the window in Current size and posistion</P>
<P>6 - Minimize</P>
<P>7 - Shows the window minimized without giving it the focus</P>
<P>9 - Restore</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"><P>' Below is an example of ShowWindow being used in a real program:</P>
<P>' Declare the FindWindow API.</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>Private Declare Function </FONT><FONT FACE="Courier New" SIZE=2>FindWindow </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">Lib </FONT><FONT FACE="Courier New" SIZE=2>"user32" </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">Alias </FONT><FONT FACE="Courier New" SIZE=2>_</P>
<P>"FindWindowA" (</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">ByVal </FONT><FONT FACE="Courier New" SIZE=2>lpClassName </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">As String</FONT><FONT FACE="Courier New" SIZE=2>, </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">ByVal </FONT><FONT FACE="Courier New" SIZE=2>_</P>
<P>lpWindowName </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">As String</FONT><FONT FACE="Courier New" SIZE=2>) </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">As Long</P>
</FONT><FONT FACE="Courier New" SIZE=2>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"><P>' Declare the ShowWindow API.</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>Private Declare Function </FONT><FONT FACE="Courier New" SIZE=2>ShowWindow </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">Lib </FONT><FONT FACE="Courier New" SIZE=2>"user32" (</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">ByVal </FONT><FONT FACE="Courier New" SIZE=2>_</P>
<P>hwnd </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">As Long</FONT><FONT FACE="Courier New" SIZE=2>, </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">ByVal </FONT><FONT FACE="Courier New" SIZE=2>nCmdShow </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">As Long</FONT><FONT FACE="Courier New" SIZE=2>) </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">As Long</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"><P>' Declare the ShowWindow Constants.</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>Private Const </FONT><FONT FACE="Courier New" SIZE=2>SW_MAXIMIZE = 3</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>Private Const </FONT><FONT FACE="Courier New" SIZE=2>SW_HIDE = 0</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>Private Const </FONT><FONT FACE="Courier New" SIZE=2>SW_NORMAL = 1</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>Private Const </FONT><FONT FACE="Courier New" SIZE=2>SW_SHOWNOACTIVATE = 4</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>Private Const </FONT><FONT FACE="Courier New" SIZE=2>SW_SHOW = 5</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>Private Const </FONT><FONT FACE="Courier New" SIZE=2>SW_MINIMIZE = 6</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>Private Const </FONT><FONT FACE="Courier New" SIZE=2>SW_SHOWMINNOACTIVE = 7</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>Private Const </FONT><FONT FACE="Courier New" SIZE=2>SW_RESTORE = 9</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>Dim </FONT><FONT FACE="Courier New" SIZE=2>FW </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">As Long</P>
</FONT><FONT FACE="Courier New" SIZE=2>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>Private Sub </FONT><FONT FACE="Courier New" SIZE=2>cmdHide_Click()</P>
<P>  FindTheNotePadWindow ("Hide")</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>End Sub</P>
</FONT><FONT FACE="Courier New" SIZE=2>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>Private Sub </FONT><FONT FACE="Courier New" SIZE=2>cmdMax_Click()</P>
<P>  FindTheNotePadWindow ("Max")</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>End Sub</P>
<P>Private Sub </FONT><FONT FACE="Courier New" SIZE=2>cmdMin_Click()</P>
<P>  FindTheNotePadWindow ("Min")</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>End Sub</P>
</FONT><FONT FACE="Courier New" SIZE=2>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>Private Sub </FONT><FONT FACE="Courier New" SIZE=2>cmdMinNF_Click()</P>
<P>  FindTheNotePadWindow ("MinNF")</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>End Sub</P>
</FONT><FONT FACE="Courier New" SIZE=2>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>Private Sub </FONT><FONT FACE="Courier New" SIZE=2>cmdRes_Click()</P>
<P>  FindTheNotePadWindow ("Res")</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>End Sub</P>
</FONT><FONT FACE="Courier New" SIZE=2>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>Private Sub </FONT><FONT FACE="Courier New" SIZE=2>cmdShow_Click()</P>
<P>  FindTheNotePadWindow ("Show")</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>End Sub</P>
</FONT><FONT FACE="Courier New" SIZE=2>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>Private Sub </FONT><FONT FACE="Courier New" SIZE=2>cmdShowCur_Click()</P>
<P>  FindTheNotePadWindow ("ShowCur")</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>End Sub</P>
</FONT><FONT FACE="Courier New" SIZE=2>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>Private Sub </FONT><FONT FACE="Courier New" SIZE=2>cmdSNF_Click()</P>
<P>  FindTheNotePadWindow ("ShowNF")</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>End Sub</P>
</FONT><FONT FACE="Courier New" SIZE=2>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>Private Sub </FONT><FONT FACE="Courier New" SIZE=2>cmdRunNotepad_Click()</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"><P>' Run notepad.</P>
</FONT><FONT FACE="Courier New" SIZE=2><P>Shell "notepad.exe", vbNormalFocus</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>End Sub</P>
</FONT><FONT FACE="Courier New" SIZE=2>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>Private Sub </FONT><FONT FACE="Courier New" SIZE=2>tmrSeeIfNPIsOpen_Timer()</P>
<P>lWnd = FindWindow("notepad", vbNullString)</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>If</FONT><FONT FACE="Courier New" SIZE=2> lWnd &lt;&gt; 0 </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">Then</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"><P>' If another instance of notepad is running, don' t allow</P>
<P>' the user to open one with the button.</P>
</FONT><FONT FACE="Courier New" SIZE=2><P>  cmdRunNotepad.Enabled = </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">False</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"><P>' If notepad is running, the buttons should be enabled.</P>
</FONT><FONT FACE="Courier New" SIZE=2><P>  cmdHide.Enabled = </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">True</P>
</FONT><FONT FACE="Courier New" SIZE=2><P>  cmdShow.Enabled = </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">True</P>
</FONT><FONT FACE="Courier New" SIZE=2><P>  cmdMax.Enabled = </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">True</P>
</FONT><FONT FACE="Courier New" SIZE=2><P>  cmdMin.Enabled = </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">True</P>
</FONT><FONT FACE="Courier New" SIZE=2><P>  cmdRes.Enabled = </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">True</P>
</FONT><FONT FACE="Courier New" SIZE=2><P>  cmdShowCur.Enabled = </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">True</P>
</FONT><FONT FACE="Courier New" SIZE=2><P>  cmdSNF.Enabled = </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">True</P>
</FONT><FONT FACE="Courier New" SIZE=2><P>  cmdMinNF.Enabled = </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">True</P>
<P>Else</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"><P>' If another instance of notepad is running, don' t allow</P>
<P>' the user to open one with the button.</P>
</FONT><FONT FACE="Courier New" SIZE=2><P>   cmdRunNotepad.Enabled = </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">True</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"><P>' If notepad isn' t running, the buttons shouldn' t be enabled.</P>
</FONT><FONT FACE="Courier New" SIZE=2><P>  cmdHide.Enabled = </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">False</P>
</FONT><FONT FACE="Courier New" SIZE=2><P>  cmdShow.Enabled = </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">False</P>
</FONT><FONT FACE="Courier New" SIZE=2><P>  cmdMax.Enabled = </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">False</P>
</FONT><FONT FACE="Courier New" SIZE=2><P>  cmdMin.Enabled = </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">False</P>
</FONT><FONT FACE="Courier New" SIZE=2><P>  cmdRes.Enabled = </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">False</P>
</FONT><FONT FACE="Courier New" SIZE=2><P>  cmdShowCur.Enabled = </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">False</P>
</FONT><FONT FACE="Courier New" SIZE=2><P>  cmdSNF.Enabled = </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">False</P>
</FONT><FONT FACE="Courier New" SIZE=2><P>  cmdMinNF.Enabled = </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">False</P>
<P>End If</P>
<P>End Sub</P>
</FONT><FONT FACE="Courier New" SIZE=2>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>Private Function </FONT><FONT FACE="Courier New" SIZE=2>FindTheNotePadWindow(Action </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">As String</FONT><FONT FACE="Courier New" SIZE=2>)</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"><P>' Find the Window.</P>
</FONT><FONT FACE="Courier New" SIZE=2><P>FW = FindWindow("notepad", vbNullString)</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"><P>' If the window exists then ...</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>If </FONT><FONT FACE="Courier New" SIZE=2>FW &lt;&gt; 0 </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">Then</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"><P>' Begin actions.</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>Select Case </FONT><FONT FACE="Courier New" SIZE=2>Action</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>Case </FONT><FONT FACE="Courier New" SIZE=2>"Hide"</P>
<P>  ShowWindow FW, SW_HIDE</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>Case </FONT><FONT FACE="Courier New" SIZE=2>"Show"</P>
<P>  ShowWindow FW, SW_NORMAL</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>Case </FONT><FONT FACE="Courier New" SIZE=2>"Min"</P>
<P>  ShowWindow FW, SW_MINIMIZE</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>Case </FONT><FONT FACE="Courier New" SIZE=2>"Max"</P>
<P>  ShowWindow FW, SW_MAXIMIZE</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>Case </FONT><FONT FACE="Courier New" SIZE=2>"MinNF"</P>
<P>  ShowWindow FW, SW_SHOWMINNOACTIVE</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>Case </FONT><FONT FACE="Courier New" SIZE=2>"Res"</P>
<P>  ShowWindow FW, SW_RESTORE</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>Case </FONT><FONT FACE="Courier New" SIZE=2>"ShowCur"</P>
<P>  ShowWindow FW, SW_SHOW</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>Case </FONT><FONT FACE="Courier New" SIZE=2>"ShowNF"</P>
<P>  ShowWindow FW, SW_SHOWNOACTIVATE</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>End Select</P>
<P>End If</P>
<P>End Function</P>
</FONT><B><FONT FACE="Arial" SIZE=2 COLOR="#ff0000"><P>WARNING: This program accesses the Windows API. This likely will not crash your system, but may under extreme circumstances.</P>
</B></FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">
</FONT><FONT FACE="Arial"><P>&#9;Now that we are finally done with ShowWindow, we can now move on to DestroyWindow. This one is the EASIEST one of the 5. It only take one parameter! The context for DestroyWindowis below:</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>Declare Function </FONT><FONT FACE="Courier New" SIZE=2>DestroyWindow </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">Lib </FONT><FONT FACE="Courier New" SIZE=2>"user32" </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">Alias </FONT><FONT FACE="Courier New" SIZE=2>"DestroyWindow" </P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"><P>' Makes a simple function to call the DestroyWindow routine within </P>
<P>' User32.dll.</P>
</FONT><FONT FACE="Courier New" SIZE=2>
<P>(</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">ByVal </FONT><FONT FACE="Courier New" SIZE=2>hwnd </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">As Long</FONT><FONT FACE="Courier New" SIZE=2>) </P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"><P>' This calls for the HANDLE of the window in which you want to close. </P>
<P>' This includes both child and parent windows!</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">
<P>As Long</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"><P>' This function doesn' t return any value if called properly!</P>
<P>' Below is an example of DestroyWindow being used in a real program:</P>
<P>' Declare the DestroyWindow API.</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>Private Declare Function</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"> </FONT><FONT FACE="Courier New" SIZE=2>DestroyWindow </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">Lib </FONT><FONT FACE="Courier New" SIZE=2>"user32" _</P>
<P>(</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">ByVal </FONT><FONT FACE="Courier New" SIZE=2>hwnd </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">As Long</FONT><FONT FACE="Courier New" SIZE=2>) </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">As Long</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00">
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>Private Sub </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00">cmdCloseWindow_Click()</P>
<P>' Destroy the window.</P>
</FONT><FONT FACE="Courier New" SIZE=2><P>DestroyWindow Me.hwnd</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#007f00"><P>' Note: This doesn' t end the project, it just closes the</P>
<P>' window. Uncomment below to close the project!</P>
<P>  ' End</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f"><P>End Sub</P>
</FONT><B><FONT FACE="Arial" SIZE=2 COLOR="#ff0000"><P>WARNING: This program accesses the Windows API. This likely will not crash your system, but may under extreme circumstances.</P>
</B></FONT><FONT FACE="Courier New" SIZE=2 COLOR="#00007f">
</FONT><FONT FACE="Arial"><P>&#9;Well, that about raps is up for the FindWindow Tutorial Part 1. All of the projects presented here today are available for download. Just get the accomanoing files to see how they work. Good luck, and be sure to keep your eyes open for part 2 of this tutorial!</P></FONT></BODY>
</HTML>

