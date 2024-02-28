SimplePOI (Excel import and export)
===========================
 This project is modified from  [AutoPOI](https://github.com/jeecgboot/autopoi) , retaining frequently used and simplified functions including Excel import and export, no templates.  I've extensively modified the original code to create a refactored version with a more robust structure, enhancing extensibility for future changes and reducing susceptibility to bugs. While the usage API remains similar to the original, the internal implementation differs. I've cut many seemingly-useless or redundant functions/parameters from the usage interface, resulting in corresponding changes in the implementation code. Along the original path, the project extend to some new features, support one-to-many structure export and import, by a recursive algorithm. This code also represents a higher-level encapsulation of [POI](https://github.com/apache/poi), offering a simplified API for exporting and importing data between lists and Excel sheets. In addition, it is better to directly run/debug the source/test(prg.simplepoi.test) code to understand how to use it rather than reading this documentation.

---------------------------
Refactor principles
--------------------------
In this section, I want to explain principles I go along with to modify the code. They are:
1. Ensure that each class/method serves a single purpose. If a class encompasses diverse functions, relocate unclassified code to a common place, such as a CommonUtility class. This enables other classes to make clear and straightforward calls when necessary.
2. Consolidate functionally similar classes, avoiding arbitrary or scattered placement. This practice reduces unnecessary coupling and enhances overall code organization.
3. Consider the use of inner classes judiciously. For straightforward classes with limited functionality, placing them as inner classes under another class can contribute to a more concise and organized structure.
4. Reduce unnecessary nested if-else.
5. Brevity. 

Overall, the project consists of two parts: export and import. Both parts need to read Excel annotations present on fields of the class to be exported or imported before operating on the Excel sheet. For the common part that will be used by both export and import, the corresponding specific changes are:
1. The org.simplepoi.excel.ReflectionUtil class is created to consolidate all reflection-related code, such as reading setMethod or getMethod, and creating entity objects.
2. The 'entity' package has been deprecated, and its classes (ExcelExportEntity, ExcelCollectionParams, ExcelImportEntity) have been relocated to the 'import' or 'export' packages, resulting in a more streamlined structure. This adjustment was made to eliminate ambiguity in functionality caused by the 'entity' package encompassing diverse purposes. The rationale is to place classes where they are most frequently used, avoiding scattering them in other locations solely based on formal similarities.
3. Package 'constant' and 'exception' are created in root package directory , making the package structure flat and easy to understand. 


In the export section, the primary focus is on converting list objects to a sheet. A challenging aspect involves addressing the merging of cells within the one-to-many structure. This functionality is  isolated into a class called ExcelManipulator. This class is responsible for considering how to recursively construct a sheet, line by line, while also handling the merging of regions. Different with original one, this project additionally support one-to-many structures with more than three levels. The corresponding specific changes are:
1. ExcelManipulator is created specifically to handle the sheet construction process, encompassing the merging of regions.
2. ExcelExportServer. Abandoning the original extend-father class structure, a single class proves sufficient. This class is dedicated to reading list objects and @ExcelEntity annotations, delegating the sheet construction task to the ExcelManipulator. 

In the import section, the primary focus is on reading Excel sheets and generating list objects. A challenging aspect involves addressing recursive reading in a one-to-many structure, supporting structures with more than three levels (distinct from the original). This functionality is independently implemented in the ExcelImportServer through recursive calls. The corresponding specific changes are: 
1. ExcelImportServer. Abandoning the original extend-father class structure, a single class proves sufficient. This class is dedicated to reading list objects and @ExcelEntity annotations, delegating the value conversion task to the CellValueServer.
2. CellValueServer is designed for a separate function, converting data from Excel into the corresponding values for target class fields. 
 


---------------------------
Future changes
--------------------------

1. I have also noticed an alternative library [EasyExcel](https://github.com/alibaba/easyexcel), which can handle Out-Of-Memory problem for batch export or import. In the future, it may be needed to incorporate its API into this project.
2. The code is still not perfectly clean (even though addition of new functions become easier),  some code are still redundant and should be cut.
3. The code support image export and import in Excel (from original one) that can be run but incomplete, which need to be improved including its API design.


---------------------------
Demo
--------------------------

A test is created for Excel import and export. In class org.simplepoi.functest.ImportExport1Test, test methods testExport and testImport provide a demo to import or export data. The result Excel is as below:
![avatar](/demo.PNG)
The method testExport export Teacher list initialized using three class Teacher, Student and Grade, to Excel file in desktop, the file is just the same as the one in resource folder of this project. The method  testImport just import the exported Excel file, to convert it to Teacher list data, then based on the converted data create an Excel file in desktop to verify the converted data is same as original.