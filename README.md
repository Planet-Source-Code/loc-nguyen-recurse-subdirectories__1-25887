<div align="center">

## Recurse Subdirectories


</div>

### Description

Does anything you want to files in a directory and its subdirectories. For example, if you would like to add MP3's in a directory to a playlist in your MP3 player program, this would be handy for adding all of the files including those residing in subdirectories of that directory.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Loc Nguyen](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/loc-nguyen.md)
**Level**          |Beginner
**User Rating**    |4.2 (21 globes from 5 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/loc-nguyen-recurse-subdirectories__1-25887/archive/master.zip)





### Source Code

```
'Create a DirListBox named "Dir1" and a FileListBox named "File1" and set their Visible properties to False
'Call:
'  DoDirs "C:\MyDir", "*.mp3"
Public Sub DoDirs(DirPath As String, DirFilters As String)
  File1.Pattern = DirFilters
  Dir1.Path = DirPath
  DoFiles DirPath
  If Dir1.ListCount = 0 Then Exit Sub
  For k = 0 To Dir1.ListCount - 1
    Dir1.Path = DirPath
    DoDirs Dir1.List(k), DirFilters
    'DoEvents
  Next k
  Dir1.Path = DirPath
End Sub
Private Sub DoFiles(DirPath As String)
  File1.Path = DirPath
  If File1.ListCount = 0 Then Exit Sub
  For k = 0 To File1.ListCount - 1
    FileName = File1.Path & String(1 - Abs(CInt(Right(File1.Path, 1) = "\")), "\") & File1.List(k)
    'Place what you would like to do to the files here
  Next k
End Sub
```

