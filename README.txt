1/19/2016
git log
--get the hash code
git fetch
git checkout -m hash code path\file

1/12/2016

work
oxApp.Workbooks.OpenText(Filename:=CurProj.Directory & "appvel.out", StartRow:=3, DataType:=Excel.XlTextParsingType.xlDelimited, TextQualifier:=Excel.XlTextQualifier.xlTextQualifierNone, Comma:=True)
NOT work
oxApp.Workbooks.OpenText(Filename:=CurProj.Directory & "appvel.out", StartRow:=3, DataType:=Excel.XlTextParsingType.xlDelimited, TextQualifier:=Excel.Constants.xlDoubleClosed, Comma:=True)

 