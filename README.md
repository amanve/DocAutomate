Errors -
1. CS0012	The type 'Object' is defined in an assembly that is not referenced. You must add a reference to assembly 'netstandard, Version=2.0.0.0, Culture=neutral, PublicKeyToken=cc7b13ffcd2ddd51'.	DocAutomate	E:\AutomateDoc\DocAutomate\DocAutomate\Form1.cs	97	Active

<b>Sln</b> : I added the missing netstandard assembly reference to DocAutomate.csproj, inside the <ItemGroup> containing the other references:
<Reference Include="netstandard, Version=2.0.0.0, Culture=neutral, PublicKeyToken=cc7b13ffcd2ddd51" />
ClosedXML uses .NET Standard 2.0, and your project needed this reference to compile calls such as new XLWorkbook(...).
After saving, reload the project if prompted and select Build → Rebuild Solution in Visual Studio.
I made this change earlier, but couldn’t confirm a successful full build because the local older MSBuild toolset had separate dependency-resolution errors.
Supported Excel files: .xlsx, .xlsm, and .xls. Generated documents keep the source extension. Without a PowerPoint selection, they are exact copies of the source workbook. Existing destination files are never overwritten.


PowerPoint software updates:
- Select an Excel workbook and optionally a `.pptx` presentation before generating.
- Slides titled `Software` or `Software Changes` must contain four-column tables: group, setting, As Is, To Be. The sample uses separate text boxes for the As Is/To Be headings; an in-table heading row is also accepted.
- Excel must contain adjacent group, setting, and value columns. Vertically merged group cells are supported. Group and setting text must match exactly, and each setting must resolve to one workbook cell.
- The current Excel value must exactly match As Is. The generated copy receives To Be. Missing/duplicate matches and formula targets stop generation before an output is created.
- Updates support `.xlsx` and `.xlsm`; legacy `.xls` supports copying only. Existing cell formatting and unrelated package parts are retained; source files are untouched.
- In `test/test.pptx`, Fan Duty / High changes from 121 to 108, and Medium from 108 to 100. In `test/hello.xlsx`, these are Sheet1!C6 and C7; the other three values remain unchanged.

Changed cells in generated Excel workbooks receive a yellow background and red text. Other cell formatting is preserved; unchanged cells keep their existing appearance.
