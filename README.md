<div align="center">

## Introduction to Error Handling


</div>

### Description

This article introduces you to the basic concepts behind Visual Basic's run-time error handling methodology. You'll learn what causes run-time errors, how to deal with them, and how to generate them yourself. A good understanding of run-time errors is critical to becoming a seasoned VB programmer. Arm yourself now.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[James Vincent Carnicelli](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/james-vincent-carnicelli.md)
**Level**          |Beginner
**User Rating**    |4.7 (28 globes from 6 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[Debugging and Error Handling](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/debugging-and-error-handling__1-26.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/james-vincent-carnicelli-introduction-to-error-handling__1-9049/archive/master.zip)





### Source Code

All VB programmers feel the kiss of death when they see a familiar run-time error message box that looks a little like this:
<P><CENTER>
<TABLE BGCOLOR="#CCCCCC" CELLSPACING="0" CELLPADDING="4" BORDER="4">
<TR><TD BGCOLOR="000066"><FONT COLOR="#FFFFFF"><B> Microsoft Visual Basic </B></FONT></TD></TR>
<TR><TD>
<BR>Run-time error '381':
<P>Invalid property array index
<BR><BR><BR>
<TABLE CELLSPACING="10"><TR>
<TD><TABLE CELLSPACING="0" CELLPADDING="2" BORDER="2"><TR><TD>&nbsp;&nbsp;&nbsp;&nbsp;
Continue
&nbsp;&nbsp;&nbsp;&nbsp;</TD></TR></TABLE></TD>
<TD><TABLE CELLSPACING="0" CELLPADDING="2" BORDER="2"><TR><TD>&nbsp;&nbsp;&nbsp;&nbsp;
End
&nbsp;&nbsp;&nbsp;&nbsp;</TD></TR></TABLE></TD>
<TD><TABLE CELLSPACING="0" CELLPADDING="2" BORDER="2"><TR><TD>&nbsp;&nbsp;&nbsp;&nbsp;
Debug
&nbsp;&nbsp;&nbsp;&nbsp;</TD></TR></TABLE></TD>
<TD><TABLE CELLSPACING="0" CELLPADDING="2" BORDER="2"><TR><TD>&nbsp;&nbsp;&nbsp;&nbsp;
Help
&nbsp;&nbsp;&nbsp;&nbsp;</TD></TR></TABLE></TD>
</TR></TABLE>
</TD></TR>
</TABLE>
</CENTER>
<P>If you've compiled a program to an executable (.EXE) and this sort of error pops up, you know by now that you don't get to debug the program. It just crashes. Is that what you want to happen? Probably not. But then, you probably wouldn't want a program to start acting unpredictably or worse because of an unexpected state of corruption. That's what critical run-time errors are supposed to prevent.
<P>But what if you actually do expect certain kinds of errors and want your program to continue running despite them? You can "trap" and handle these errors. To "trap" an error simply means to allow an error to occur on the assumption that your code will deal with it. There are two basic ways to trap and handle an error: "resume" and "go-to". They can be illustrated by the following examples:
<UL><PRE>
<FONT COLOR="#009900">'"Resume" approach</FONT>
<FONT COLOR="#000099">Sub</FONT> Demo1
&nbsp;&nbsp;&nbsp;&nbsp;<FONT COLOR="#000099">On Error Resume Next</FONT>
&nbsp;&nbsp;&nbsp;&nbsp;X = 1 / 0 <FONT COLOR="#009900">'Division by zero</FONT>
&nbsp;&nbsp;&nbsp;&nbsp;<FONT COLOR="#000099">MsgBox</FONT> Err.Description
&nbsp;&nbsp;&nbsp;&nbsp;<FONT COLOR="#000099">On Error GoTo 0</FONT>
<FONT COLOR="#000099">End Sub</FONT>
</PRE></UL>
<UL><PRE>
<FONT COLOR="#009900">'"Go-To" approach</FONT>
<FONT COLOR="#009900">'This is not currently applicable to VBScript</FONT>
<FONT COLOR="#000099">Sub</FONT> Demo2
&nbsp;&nbsp;&nbsp;&nbsp;<FONT COLOR="#000099">On Error GoTo</FONT> Oopsie
&nbsp;&nbsp;&nbsp;&nbsp;X = 1 / 0 <FONT COLOR="#009900">'Division by zero</FONT>
<BR>&nbsp;&nbsp;&nbsp;&nbsp;Exit Sub
Oopsie:
&nbsp;&nbsp;&nbsp;&nbsp;<FONT COLOR="#000099">MsgBox</FONT> Err.Description
<FONT COLOR="#000099">End Sub</FONT>
</PRE></UL>
<P>The key difference between these two approaches to error handling is that <TT>On Error Resume Next</TT> tells VB you want your code to keep executing as if nothing had happened, whereas <TT>On Error GoTo <I>Some_Label</I></TT> tells VB you want execution to jump to some specific location in your routine at any time a run-time error occurs.
<P>Notice the use of <TT>On Error GoTo 0</TT> in <TT>Demo1</TT> above? Although it looks like a contorted version of <TT>On Error GoTo <I>Label</I></TT>, it's actually a special way to tell VB that you want to stop trapping errors and let VB perform its own built-in handling.
<P>Recovering gracefully from a run-time error, once you've trapped it, really requires you to make use of the Err object. Err is an object VB uses to give your program access to information about the error. Here are the most important public members Err exposes:
<P><CENTER><TABLE WIDTH="90%" CELLSPACING="0" CELLPADDING="2" BORDER="2">
<TR><TD><TT> Err.Number </TT></TD><TD>
Long integer indicating the error code number. This is pretty much useless except where the vendor of the product that generated this error was too lazy to provide a useful description.
</TD></TR>
<TR><TD><TT> Err.Source </TT></TD><TD>
Generally used to tell your handler what component or code element is responsible for generating the error. With custom errors, you might want to set this to <TT><NOBR>"ModuleName.MethodName()"</NOBR></TT>.
</TD></TR>
<TR><TD><TT> Err.Description </TT></TD><TD>
The all-important, human-readable description. The point of this is so you're not left scratching your head wondering "what the heck does '<NOBR>-10021627</NOBR>' mean?"
</TD></TR>
<TR><TD><TT> Err.Clear() </TT></TD><TD>
Allows you to sweep the error under the rug, so to speak.
</TD></TR>
<TR><TD><TT> Err.Raise(Number, [Source], [Description], [HelpFile], [HelpContext]) </TT></TD><TD>
Allows you to "raise", or invoke, your own run-time error. Number can be <TT>vbObjectError + CustomErrorCode</TT> if you're not raising one of the standard ones. Be sure to provide a source and description.
</TD></TR>
</TABLE></CENTER>
<P>The <TT>.HelpFile</TT> and <TT>.HelpContext</TT> properties, not listed above, can be used by your program to refer users to a relevant passage in some help file. Few programs bother.
<P>The nice thing about go-to error trapping is that it allows you to easily enwrap a large chunk of code with your error handler with one single line of code (<TT>On Error GoTo <I>Label</I></TT>). The resume approach really requires you to either include error handling code after every line or to take a blind leap of faith that a given line will either never encounter an error or that it won't matter. As a general rule, use On Error Resume Next only for short blocks of code.
<P>One of the interesting nuances of the VB run-time error mechanism is that it propagates errors "backwards". To illustrate what this means, consider the following code:
<UL><PRE>
<FONT COLOR="#000099">Sub</FONT> A
&nbsp;&nbsp;&nbsp;&nbsp;<FONT COLOR="#000099">On Error Resume Next</FONT>
&nbsp;&nbsp;&nbsp;&nbsp;<FONT COLOR="#000099">Call</FONT> B
&nbsp;&nbsp;&nbsp;&nbsp;<FONT COLOR="#000099">MsgBox</FONT> Err.Description
<FONT COLOR="#000099">End Sub</FONT>
</PRE></UL>
<UL><PRE>
<FONT COLOR="#000099">Sub</FONT> B
&nbsp;&nbsp;&nbsp;&nbsp;<FONT COLOR="#000099">Call</FONT> C
<FONT COLOR="#000099">End Sub</FONT>
</PRE></UL>
<UL><PRE>
<FONT COLOR="#000099">Sub</FONT> C
&nbsp;&nbsp;&nbsp;&nbsp;X = 1 / 0 <FONT COLOR="#009900">'Division by zero</FONT>
<FONT COLOR="#000099">End Sub</FONT>
</PRE></UL>
<P><TT>A</TT> calls <TT>B</TT>, which in turn calls <TT>C</TT>. Since <TT>C</TT> will cause a division-by-zero run-time error and itself has no error handler, VB will effectively leave <TT>C</TT> and go back to <TT>B</TT>. But <TT>B</TT> doesn't have an error handler, either, so VB leaves <TT>B</TT> to go back to <TT>A</TT>. Fortunately, <TT>A</TT> does have an error handler. If it didn't, <TT>A</TT> would also immediately exit and control would go back to whatever called it. If there's nothing left up this "calling stack", your program will courteously commit suicide.
<P>You can use this "backward propagation" property of VB's error mechanism to your advantage in many ways. First, you can enwrap a block of code by putting it in its own subroutine and putting your error handler in the code that calls that subroutine. In this case, any run-time error in that subroutine will propagate back to your calling code. Second, you can add value to an error message by adding more context information. You might use code like the following, for instance:
<UL><PRE>
<FONT COLOR="#000099">Sub</FONT> A
&nbsp;&nbsp;&nbsp;&nbsp;<FONT COLOR="#000099">On Error GoTo</FONT> AwShoot
&nbsp;&nbsp;&nbsp;&nbsp;<FONT COLOR="#000099">Call</FONT> B
&nbsp;&nbsp;&nbsp;&nbsp;<FONT COLOR="#000099">Exit Sub</FONT>
AwShoot:
&nbsp;&nbsp;&nbsp;&nbsp;Err.Raise vbObjectError, "MyModule.A(): " & Err.Source, _
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"Unexpected failure in A: " & Err.Description
<FONT COLOR="#000099">End Sub</FONT>
</PRE></UL>
<UL><PRE>
<FONT COLOR="#000099">Sub</FONT> B
&nbsp;&nbsp;&nbsp;&nbsp;<FONT COLOR="#000099">On Error GoTo</FONT> AwShoot
&nbsp;&nbsp;&nbsp;&nbsp;Err.Raise vbObjectError, "My left nostril", "Stabbing pain"
&nbsp;&nbsp;&nbsp;&nbsp;<FONT COLOR="#000099">Exit Sub</FONT>
AwShoot:
&nbsp;&nbsp;&nbsp;&nbsp;Err.Raise vbObjectError, "MyModule.B(): " & Err.Source, _
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"Couldn't complete B: " & Err.Description
<FONT COLOR="#000099">End Sub</FONT>
</PRE></UL>
<P>Calling <TT>A</TT> will result in an error whose source is <TT>"MyModule.A(): MyModule.B(): My left nostril"</TT> and whose description is <TT>"Unexpected failure in A: Couldn't complete B: Stabbing pain"</TT>. Having the extra "source" information probably won't help your end-users. But then, your end users probably won't care about the source of the problem, any way. But as the person who gets to fix it, this will be invaluable to you. The extra description information might actually help your end users, but it too will be invaluable to you. Note, incidentally, that calling <TT>Err.Raise()</TT> in your error handler will not cause the error to be thrown back to itself, again. With the go-to method of error handling, as soon as the error is raised and before control is passed to your error handler (right after the <TT>AwShoot:</TT> line label), the error handler for your routine is automatically switched off. If you want to trap errors in your error handler code, you'll have to reset the error handler with another <TT>On Error Resume Next</TT> or <TT>On Error GoTo <I>Some_Other_Label</I></TT> line in your handler.
<P>For those times you use the resume approach, be aware that calling <TT>On Error GoTo 0</TT> not only disables error handling in the current routine, it also clears the current error properties, including the description. If you want to add your own custom error message before propagating the error back up the call stack in a fashion like that above, you'll need to grab the properties from Err, first. Here's a simple way to do it:
<UL><PRE>
<FONT COLOR="#000099">Sub</FONT> Doodad
&nbsp;&nbsp;&nbsp;&nbsp;<FONT COLOR="#000099">On Error Resume Next</FONT>
&nbsp;&nbsp;&nbsp;&nbsp;X = 1 / 0
<BR>&nbsp;&nbsp;&nbsp;&nbsp;<FONT COLOR="#000099">If</FONT> Err.Number <> 0 <FONT COLOR="#000099">Then</FONT>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<FONT COLOR="#009900">'Dump Err properties into an array</FONT>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;EP = Array(Err.Number, Err.Source, _
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Err.Description, Err.HelpFile, Err.HelpContext)
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<FONT COLOR="#009900">'Re-enable VB's own error handler</FONT>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<FONT COLOR="#000099">On Error GoTo 0</FONT>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<FONT COLOR="#009900">'Propagate error back up the call stack with my two cents added</FONT>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Err.Raise EP(0), "MyModule.Doodad(): ", EP(1), _
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"Something bad happened: " & EP(2), EP(3), EP(4)
&nbsp;&nbsp;&nbsp;&nbsp;<FONT COLOR="#000099">End If</FONT>
&nbsp;&nbsp;&nbsp;&nbsp;<FONT COLOR="#000099">On Error GoTo 0</FONT>
<BR>&nbsp;&nbsp;&nbsp;&nbsp;<FONT COLOR="#000099">Exit Sub</FONT>
<FONT COLOR="#000099">End Sub</FONT>
</PRE></UL>
<P>Finally, let me strongly urge you to have your programs raise errors as a natural matter of course. Functions often return special values like 0, "", Null and so on to indicate that an error has occurred. Instead of doing this and requiring your users (other programmers) to figure out your special error representations and to make non-standard error handlers for them, try calling <TT>Err.Raise()</TT>. If your users don't realize that an invocation of your code may cause an error, the first case may leave them with a difficult mystery to solve, whereas the second case will leave little doubt about the real cause. Plus, they'll be able to make their code more readable and consistent with best-practice standards.
<P>In summary, VB's run-time error trapping and handling mechanism allows your code to take control of how errors are managed. This can be used to allow your programs to more gracefully end, to let your programs continue running despite certain kinds of problems, to give developers better clues about the causes of bugs in their code, and more. There are two basic approaches: "resume" and "go-to". VB's built-in Err object holds the information you need to find out where the error occurred and what its nature is and allows you to clear or raise errors of your own.

