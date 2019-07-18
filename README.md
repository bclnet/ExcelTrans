# ExcelTrans

**ExcelTrans** Transforms Csv to Excel spreadsheets using .NET and EPPlus.

It has the following services:

* ExcelTrans a command based stream csv transformer
* CsvReader for quickly reading csv documents
* CsvWriter for quickly writing csv documents
* ExcelReader for quickly reading excel documents

# Source Code

This repository holds the implementation of ExcelTrans in C#.


## Simple Example

```csharp
static Tuple<Stream, string, string> MakeInvoiceFile(IEnumerable<MyData> myData)
{
    var transform = ExcelService.Encode(new List<IExcelCommand>
    {
        new WorksheetsAdd("Invoice"),
        new CellsStyle(Address.Range, 0, 1, 2, 1, "lc:Yellow"),
    });
    
    var s = new MemoryStream();
    var w = new StreamWriter(s);
    // add transform to output
    w.WriteLine(transform);
    // add csv file to output
    CsvWriter.Write(w, myData);
    w.Flush(); s.Position = 0;
    var result = new Tuple<Stream, string, string>(s, "text/csv", "invoice.csv");
    // optionally transform
    result = ExcelService.Transform(result);
    return result;
}

static void TransferFile(string path, Stream stream, string file)
{
    path = Path.Combine(path, file);
    if (!Directory.Exists(Path.GetDirectoryName(path)))
        Directory.CreateDirectory(Path.GetDirectoryName(path));
    using (var fileStream = File.Create(path))
    {
        stream.CopyTo(fileStream);
        stream.Seek(0, SeekOrigin.Begin);
    }
}

var path = ...some path...;
var myData = ...some data...;
var file = MakeInvoiceFile(myData);
TransferFile(path, file.Item1, file.Item3);
```


# Reference

## Commands
*List of commands available*

Command     | Description | See Also
---         | --- | ---:
CellsStyle  | Applies `.Styles` to the `.Cells` in range | [Styles](#styles)
CellsValue  | Applies `.Value` of `.ValueKind` to the `.Cells` in range | [CellValueKind](#cellvaluekind)
ColumnValue | Applies `.Value` of `.ValueKind` to the `.Col` column | [ColumnValueKind](#columnvaluekind)
Command     | Executes `action()`
CommandCol  | Executes `func()` per Column
CommandRow  | Executes `func()` per Row
ConditionalFormatting | Applies `.Value` with conditional-formatting of `.FormattingKind` to `.Address` | [ConditionalFormattingKind](#conditionalformattingkind)
Flush       | Flushes all pending commands
PopFrame    | Pops a Frame off the context stack
PopSet      | Pops a Set off the context stack
PushFrame   | Pushes a new Frame with `cmds` onto the context stack
PushSet     | Pushes a new Set with `group` and `cmds` onto the context stack
RowValue    | Applies `.Value` of `.ValueKind` to the `.Row` row | [RowValueKind](#rowvaluekind)
ViewAction  | Applies `.Value` of `.ActionKind` to the active spreadsheet | [ViewActionKind](#viewactionkind)
WorkbookOpen    | Opens a Workbook at `.Path`
WorksheetsAdd   | Adds a Worksheet with `.Name` to current Workbook
WorksheetsCopy  | Copies the Worksheet with `.Name` to a new Worksheet with `.NewName` to the current Workbook
WorksheetsDelete| Deletes a Worksheet with `.Name` from the current Workbook
WorksheetsOpen  | Opens a Worksheet with `.Name` from the current Workbook


## Styles
*Values for the CellsStyle command*

`n*`| The numberformat  | Description
--- | ---:              | ---
n:* | *Format*          | Set the range to a specific value
n$* | *NumberformatPrec*| Set the range to a specific value
n%* | *NumberformatPrec*| Set the range to a specific value
n,* | *NumberformatPrec*| Set the range to a specific value
nd  | ShortDatePattern  | Set the range to a specific value

`f*`| Font styling      | Description
--- | ---:              | ---
f:* | *Font*            | The name of the font
fx* | *Size*            | The Size of the font
ff* | *Family*          | Font family
fc:*| *Color*           | Cell color
fs:*| *Scheme*          | Scheme
fB  | true              | Font-bold
fb  | false             | Font-bold
fI  | true              | Font-italic
fi  | false             | Font-italic
fS  | true              | Font-Strikeout
fs  | false             | Font-Strikeout
f_  | true              | Font-Underline
f!_ | false             | Font-Underline
fv* | *ExcelVerticalAlignmentFont*| Font-Vertical Align

`l*`| Fill styling      | Description
--- | ---:              | ---
lc:*| *Color*           | The background color
lf* | *ExcelFillStyle*  | The pattern for solid fills.

`b*`| Border            | Description
--- | ---:              | ---
bl* | *ExcelBorderStyle*| Left border style
br* | *ExcelBorderStyle*| Right border style
bt* | *ExcelBorderStyle*| Top border style
bb* | *ExcelBorderStyle*| Bottom border style
bdU | true              | A diagonal from the bottom left to top right of the cell
bdu | false             | A diagonal from the bottom left to top right of the cell
bdD | true              | A diagonal from the top left to bottom right of the cell
bdd | false             | A diagonal from the top left to bottom right of the cell
bd* | *ExcelBorderStyle*| Diagonal border style
ba* | *ExcelBorderStyle*| Set the border style around the range.

`ha*`| Horizontal alignment| Description
--- | ---:              | ---
ha* | *ExcelHorizontalAlignment*| The horizontal alignment in the cell

`va*`| Vertical alignment| Description
--- | ---:              | ---
va* | *ExcelVerticalAlignment*| The vertical alignment in the cell

`*` | Style             | Description
--- | ---:              | ---
W   | true              | Wrap the text
w   | false             | Wrap the text


## CellValueKind
*Values for the CellValue command*

Enum            | Description
---             | ---
Value           | Set the range to a specific value
Text            | Returns the formatted value.
AutoFilter      | Set an autofilter for the range
AutoFitColumns  | Set the column width from the content of the range. The minimum width is the value of the ExcelWorksheet.defaultColumnWidth property. Note: Cells containing formulas must be calculated before autofit is called. Wrapped and merged cells are also ignored.
Comment         | The comment text
CommentMore     | n/a
ConditionalFormattingMore | n/a
Copy            | Copies the range of cells to an other range
Formula         | Gets or sets a formula for a range.
FormulaR1C1     | Gets or Set a formula in R1C1 format.
Hyperlink       | Set the hyperlink property for a range of cells
Merge           | If the cells in the range are merged.
RichText        | Add a rich text string
RichTextClear   | Clear the collection
StyleName       | The named style


## ColumnValueKind
*Values for the ColumnValue command*

Enum            | Description
---             | ---
AutoFit         | Set the column width from the content of the range. The minimum width is the value of the ExcelWorksheet.defaultColumnWidth property. Note: Cells containing formulas are ignored since EPPlus don't have a calculation engine. Wrapped and merged cells are also ignored.
BestFit         | If set to true a column automaticlly resize(grow wider) when a user inputs numbers in a cell.
Merged          | none
Width           | Sets the width of the column in the worksheet
TrueWidth^      | Set width to a scaled-value that should result in the nearest possible value to the true desired setting.


## ConditionalFormattingKind
*Values for the ConditionalFormatting command*

Enum            | Description
---             | ---
AboveAverage    | Add AboveAverage Rule
AboveOrEqualAverage | Add AboveOrEqualAverage Rule
AboveStdDev     | Add AboveStdDev Rule
BeginsWith      | Add BeginsWith Rule
BelowAverage    | Add BelowAverage Rule
BelowOrEqualAverage | Add BelowOrEqualAverage Rule
BelowStdDev     | Add BelowStdDev Rule
Between         | Add Between Rule
Bottom          | Add Bottom Rule
BottomPercent   | Add BottomPercent Rule
ContainsBlanks  | Add ContainsBlanks Rule
ContainsErrors  | Add ContainsErrors Rule
ContainsText    | Add ContainsText Rule
Databar         | Adds a databar rule
DuplicateValues | Add DuplicateValues Rule
EndsWith        | Add EndsWith Rule
Equal           | Add Equal Rule
Expression      | Add Expression Rule
FiveIconSet     | Adds a FiveIconSet rule
FourIconSet     | Adds a FourIconSet rule
GreaterThan     | Add GreaterThan Rule
GreaterThanOrEqual | Add GreaterThanOrEqual Rule
Last7Days       | Add Last7Days Rule
LastMonth       | Add LastMonth Rule
LastWeek        | Add LastWeek Rule
LessThan        | Add LessThan Rule
LessThanOrEqual | Add LessThanOrEqual Rule
NextMonth       | Add NextMonth Rule
NextWeek        | Add NextWeek Rule
NotBetween      | Add NotBetween Rule
NotContainsBlanks | Add NotContainsBlanks Rule
NotContainsErrors | Add NotContainsErrors Rule
NotContainsText | Add NotContainsText Rule
NotEqual        | Add NotEqual Rule
ThisMonth       | Add ThisMonth Rule
ThisWeek        | Add ThisWeek Rule
ThreeColorScale | Add ThreeColorScale Rule
ThreeIconSet    | Add ThreeIconSet Rule
Today           | Add Today Rule
Tomorrow        | Add Tomorrow Rule
Top             | Add Top Rule
TopPercent      | Add TopPercent Rule
TwoColorScale   | Add TwoColorScale Rule
UniqueValues    | Add Unique Rule
Yesterday       | Add Yesterday Rule


## RowValueKind
*Values for the RowValue command*

Enum            | Description
---             | ---
Value           | Set the range to a specific value
Text            | Returns the formatted value.
AutoFilter      | Set an autofilter for the range
AutoFitColumns  | Set the column width from the content of the range. The minimum width is the value of the ExcelWorksheet.defaultColumnWidth property. Note: Cells containing formulas must be calculated before autofit is called. Wrapped and merged cells are also ignored.
Comment         | The comment text
CommentMore     | n/a
ConditionalFormattingMore | n/a
Copy            | Copies the range of cells to an other range
Formula         | Gets or sets a formula for a range.
FormulaR1C1     | Gets or Set a formula in R1C1 format.
Hyperlink       | Set the hyperlink property for a range of cells
Merge           | If the cells in the range are merged.
RichText        | Add a rich text string
RichTextClear   | Clear the collection
StyleName       | The named style


## ViewActionKind
*Values for the ViewAction command*

Enum            | Description
---             | ---
FreezePane      | Freeze the columns/rows to left and above the cell
SetTabSelected  | Sets whether the worksheet is selected within the workbook.
UnfreezePane    |  Unlock all rows and columns to scroll freely


# Author

The author of this library is [Sky Morey](https://www.linkedin.com/in/sky-morey/).
