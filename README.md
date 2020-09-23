<div align="center">

## Relative Path Function \<BR\>\<SMALL\>Gets the relative path of a file w\.r\.t\. a directory\</SMALL\>


</div>

### Description

A reusable function that computes the relative path of a file with respect to a given directory

<BR>

Examples will make the point clear, so here goes:

<PRE style="font-family: Courier New, monospaced; font-size: 10pt;">

<OL>

<LI><P>GetRelativePath ("C:\VB\", "C:\VB\File.ext")

returns "File.ext"</P>

<LI><P>GetRelativePath ("C:\VB\", "C:\VB\Program\File.ext")

returns "Program\File.ext"</P>

<LI><P>GetRelativePath ("C:\VB\", "C:\File.ext")

returns "..\File.ext"</P>

</OL>

</PRE>

<P>It is useful to insert images and hyperlinks into webpages, given the filenames of the images and the HTML file.</P>
 
### More Info
 
<OL>

<LI><P><B>sBase</B><BR>Fully Qualified Path of the Base Directory</P>

<LI><P><B>sFile</B><BR>Fully Qualified Path of the File of which the relative path is to be computed.</P>

</OL>

Just remember to pass *complete* qualified paths to it. The CommonDialog.Filename property can be directly assigned in the call. e.g.

<PRE>GetRelativePath ("C:\VB\", CommonDialog1.Filename)</PRE>

<P>Relative Path of sFile with respect to sBase.</P>

<P>None.</P>


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Manas Tungare](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/manas-tungare.md)
**Level**          |Intermediate
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, ASP \(Active Server Pages\) 
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/manas-tungare-relative-path-function-br-small-gets-the-relative-path-of-a-file-w-r-t-a-dir__1-12589/archive/master.zip)





### Source Code

```
Public Function GetRelativePath(sBase As String, sFile As String)
'------------------------------------------------------------
' Accepts : sBase= Fully Qualified Path of the Base Directory
'  sFile= Fully Qualified Path of the File of which
'   the relative path is to be computed.
' Returns : Relative Path of sFile with respect to sBase.
' Modifies: Nothing.
'------------------------------------------------------------
' Author : Manas Tungare (www.manastungare.com)
'------------------------------------------------------------
Dim Base() As String, File() As String
Dim I As Integer, NewTreeStart As Long, sRel As String
 If Left(sBase, 3) <> Left(sFile, 3) Then
 'Since the files lie on different drives, the relative
 'filename is same as the Absolute Filename
 GetRelativePath = sFile
 Exit Function
 End If
 Base = Split(sBase, "\")
 File = Split(sFile, "\")
 While Base(I) = File(I)
 I = I + 1
 Wend
 If I = UBound(Base) Then
 'Then the Base Path is over, and the file lies
 'in a subdirectory of the base directory.
 'So simply append the rest of the path.
 While I <= UBound(File)
  sRel = sRel + File(I) + "\"
  I = I + 1
 Wend
 'Now remove the extra trailing "\" we put earlier.
 GetRelativePath = Left(sRel, Len(sRel) - 1)
 Exit Function
 End If
 NewTreeStart = I
 'The base path is not yet over, and we need to step
 'back using the "..\"
 While I < UBound(Base)
 sRel = sRel & "..\"
 I = I + 1
 Wend
 While NewTreeStart <= UBound(File)
 sRel = sRel & File(NewTreeStart) + "\"
 NewTreeStart = NewTreeStart + 1
 Wend
 'Now remove the extra trailing "\" we put earlier.
 GetRelativePath = Left(sRel, Len(sRel) - 1)
End Function
```

