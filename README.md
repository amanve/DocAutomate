Errors -
1. CS0012	The type 'Object' is defined in an assembly that is not referenced. You must add a reference to assembly 'netstandard, Version=2.0.0.0, Culture=neutral, PublicKeyToken=cc7b13ffcd2ddd51'.	DocAutomate	E:\AutomateDoc\DocAutomate\DocAutomate\Form1.cs	97	Active

<b>Sln</b> : I added the missing netstandard assembly reference to DocAutomate.csproj, inside the <ItemGroup> containing the other references:
<Reference Include="netstandard, Version=2.0.0.0, Culture=neutral, PublicKeyToken=cc7b13ffcd2ddd51" />
ClosedXML uses .NET Standard 2.0, and your project needed this reference to compile calls such as new XLWorkbook(...).
After saving, reload the project if prompted and select Build → Rebuild Solution in Visual Studio.
I made this change earlier, but couldn’t confirm a successful full build because the local older MSBuild toolset had separate dependency-resolution errors.
Supported Excel files: .xlsx, .xlsm, and .xls. Generated documents keep the source extension. Without a PowerPoint selection, they are exact copies of the source workbook. Existing destination files are never overwritten.


