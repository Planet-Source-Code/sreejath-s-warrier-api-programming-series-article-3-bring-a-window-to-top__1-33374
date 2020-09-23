<div align="center">

## API Programming Series Article \#3 Bring a Window to Top


</div>

### Description

In the third article we see how to bring a window to top using the Win32 API.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Sreejath S\. Warrier](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/sreejath-s-warrier.md)
**Level**          |Intermediate
**User Rating**    |4.0 (20 globes from 5 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/sreejath-s-warrier-api-programming-series-article-3-bring-a-window-to-top__1-33374/archive/master.zip)





### Source Code

<p style='margin-bottom:12.0pt'><span style='font-size:10.0pt;
font-family:Verdana;color:black'>In the previous article we saw how to declare
and invoke API functions from Visual Basic. In this article we see a small
example how to bring a window to top.<br>
As usual we start with the API declaration.<br>
<br>
Create a new VB Standard EXE project.<br>
When you created the project, a form Form1 should have been added to the
project by default. Add another form to the project. Since this is only an
example to illustrate the above API call, we'll not change the properties of
the forms. So now we have two forms named Form1 and Form2 in the project. Add a
command button each to both the forms. Leave their names as Command1 itself.<br>
In the General | Declarations section of both the forms, type in the following
code</span></p>
<pre><span style='color:black'><br>
Private Declare Function BringWindowToTop Lib "user32" Alias "BringWindowToTop" (ByVal hwnd As Long) As Long</span></pre>
<p style='margin-bottom:12.0pt'><span style='font-size:10.0pt;
font-family:Verdana;color:black'><br>
<br>
In the Click event procedure of the button on Form1 add the following code:</span></p>
<pre><span style='color:black'><br>
Private Sub Command1_Click ()<br>
BringWindowToTop Form2.hwnd<br>
End Sub</span></pre>
<p style='margin-bottom:12.0pt'><span style='font-size:10.0pt;
font-family:Verdana;color:black'><br>
<br>
<br>
In the Click event procedure of the button on Form2 add the following code:</span></p>
<pre><span style='color:black'><br>
Private Sub Command1_Click ()<br>
BringWindowToTop Form1.hwnd<br>
End Sub</span></pre>
<p style='margin-bottom:12.0pt'><span style='font-size:10.0pt;
font-family:Verdana;color:black'><br>
<br>
<br>
Now in the load event of Form1 (which should be the default form of the project
add the following code.</span></p>
<pre><span style='color:black'><br>
Private Sub Form_Load ()<br>
Form2.Show<br>
End Sub</span></pre>
<p><span style='font-size:10.0pt;font-family:Verdana;
color:black'><br>
<br>
<br>
Now if we press the command button on Form1, Form2 will be brought to top and
vice versa.<br>
<b>Analysis</b></span></p>
<p><span style='font-size:10.0pt;font-family:Verdana;
color:black'>Let us see how this works.<br>
First we declared the API function that we’re going to use, which in this case
is the BringWindowToTop encapsulated in the user32.dll. If you are not familiar
with the mechanics of declaring and invoking API functions, please go through
the Parts 1 and 2 of this series which describe the basics of API programming
in considerable detail.</span></p>
<p><span style='font-size:10.0pt;font-family:Verdana;
color:black'> </span></p>
<p><span style='font-size:10.0pt;font-family:Verdana;
color:black'>The API function BringWindowToTop accepts the hwnd (Handle to the
Window, a unique id that all windows have) of the window that is to be brought
on top. It brings the specified window to the top of the Z order. If the window
is a top-level window, it is activated. If the window is a child window, the
top-level parent window associated with the child window is activated.<br>
While adequate for the purpose of explaining the function, the above example is
rather trivial in nature. I.e. it doesn't achieve anything practical. So what
would be a practical application for this? Hmmm, say, you've got a long process
running in a window. Naturally you can expect your users to switch to other
windows during this period. However, once the process is complete you may want
to put this window on top. In such a case this code can be put to use. </span></p>
<p><span style='font-size:10.0pt;font-family:Verdana;
color:black'> </span></p>
<p><b><span style='font-size:10.0pt;font-family:Verdana;
color:black'>Summary</span></b></p>
<p><span style='font-size:10.0pt;font-family:Verdana;
color:black'> </span></p>
<p><span style='font-size:10.0pt;font-family:Verdana;
color:black'>In this article we saw how to make a Window come on top of all
other windows using the API. If you have any questions or comments, please feel
free to contact me.</span></p>

