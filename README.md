<div align="center">

## Print MSHFlexGrid


</div>

### Description

This function retrieves data from MSHFlex Grid and prints it directly to the printer. It determines whether the information should be printed landscape or portrait.
 
### More Info
 
you must supply MSHFlex Grid

e.g. X = PrintMSHGrid(MSHFlexGrid1)

The function is currently limited to 50 columns, but it can be increased.

the function does not return anything


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Irene V\. Yuzbasheva](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/irene-v-yuzbasheva.md)
**Level**          |Intermediate
**User Rating**    |3.8 (15 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/irene-v-yuzbasheva-print-mshflexgrid__1-6504/archive/master.zip)





### Source Code

```
Private Sub Command1_Click()
  X = PrintMSHGrid(MSHFlexGrid1)
End Sub
Public Function PrintMSHGrid(ByVal GridToPrint As MSHFlexGrid) As Long
'This function retrieves data from MSHFlexGrid and prints it directly to the
'printer. It uses MyArray to store the distance between columns. The max number
'of columns is 50, but it can be increased if there is a need.
'Print information from mshflexgrid
  Dim MyRows, MyCols As Integer  'for-loop counters
  Dim MyText As String      'text to be printed
  Dim Titles As String      'column titles
  Dim Header As String      'page headers
  Dim MyLines As Integer     'number of lines for portrait/landscape
  Dim LLCount As Integer     'temporary line counter
  Dim MyArray(50) As Integer
  Screen.MousePointer = vbHourglass
  Titles = ""
  LLCount = 0
  Header = " - Page: "          'setup page header
  'get column headers
  For MyCols = 0 To GridToPrint.Cols - 1
    MyArray(MyCols) = Len(GridToPrint.ColHeaderCaption(0, MyCols)) + 15
    Titles = Titles & Space(15) & GridToPrint.ColHeaderCaption(0, MyCols)
  Next MyCols
  'setup printer
  Printer.Font.Size = 8          '8pts font size
  Printer.Font.Bold = True        'titles to be bold
  Printer.Font.Name = "Courier New"    'courier new font
  'determine whether to print landscape or portrait
  If (Len(MyText) > 120) Then       'landscape
    Printer.Orientation = vbPRORLandscape
    MyLines = 60
  Else                  'portrait
    Printer.Orientation = vbPRORPortrait
    MyLines = 85
  End If
  Printer.Print Header; Printer.Page
  Printer.Print Titles
  Printer.Font.Bold = False
  'get column/row values
  For MyRows = 1 To GridToPrint.Rows - 1
    MyText = ""
    GridToPrint.Row = MyRows
    For MyCols = 0 To GridToPrint.Cols - 1
      GridToPrint.Col = MyCols
        MyText = MyText & GridToPrint.Text & Space(MyArray(MyCols) - Len(GridToPrint.Text))
    Next MyCols
    LLCount = LLCount + 1
    If LLCount <= MyLines Then
      Printer.Print MyText
    Else
      Printer.NewPage
      Printer.Print Header; Printer.Page
      Printer.Print Titles
      Printer.Print MyText
      LLCount = 0
    End If
  Next MyRows
  Printer.EndDoc
  Screen.MousePointer = vbNormal
End Function
```

