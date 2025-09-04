// table 7001239 "Excel Buffer 2"
// {
//     Caption = 'Excel Buffer';
//     ReplicateData = false;

//     fields
//     {
//         field(1; "Row No."; Integer)
//         {
//             Caption = 'Row No.';
//             DataClassification = SystemMetadata;

//             trigger OnValidate()
//             begin
//                 xlRowID := '';
//                 if "Row No." <> 0 then
//                     xlRowID := Format("Row No.");
//             end;
//         }
//         field(2; xlRowID; Text[10])
//         {
//             Caption = 'xlRowID';
//             DataClassification = SystemMetadata;
//         }
//         field(3; "Column No."; Integer)
//         {
//             Caption = 'Column No.';
//             DataClassification = SystemMetadata;

//             trigger OnValidate()
//             var
//                 x: Integer;
//                 i: Integer;
//                 y: Integer;
//                 c: Char;
//                 t: Text[30];
//             begin
//                 xlColID := '';
//                 x := "Column No.";
//                 while x > 26 do begin
//                     y := x mod 26;
//                     if y = 0 then
//                         y := 26;
//                     c := 64 + y;
//                     i := i + 1;
//                     t[i] := c;
//                     x := (x - y) div 26;
//                 end;
//                 if x > 0 then begin
//                     c := 64 + x;
//                     i := i + 1;
//                     t[i] := c;
//                 end;
//                 for x := 1 to i do
//                     xlColID[x] := t[1 + i - x];
//             end;
//         }
//         field(4; xlColID; Text[10])
//         {
//             Caption = 'xlColID';
//             DataClassification = SystemMetadata;
//         }
//         field(5; "Cell Value as Text"; Text[250])
//         {
//             Caption = 'Cell Value as Text';
//             DataClassification = SystemMetadata;
//         }
//         field(6; Comment; Text[250])
//         {
//             Caption = 'Comment';
//             DataClassification = SystemMetadata;
//         }
//         field(7; Formula; Text[250])
//         {
//             Caption = 'Formula';
//             DataClassification = SystemMetadata;
//         }
//         field(8; Bold; Boolean)
//         {
//             Caption = 'Bold';
//             DataClassification = SystemMetadata;
//         }
//         field(9; Italic; Boolean)
//         {
//             Caption = 'Italic';
//             DataClassification = SystemMetadata;
//         }
//         field(10; Underline; Boolean)
//         {
//             Caption = 'Underline';
//             DataClassification = SystemMetadata;
//         }
//         field(11; NumberFormat; Text[250])
//         {
//             Caption = 'NumberFormat';
//             DataClassification = SystemMetadata;
//         }
//         field(12; Formula2; Text[250])
//         {
//             Caption = 'Formula2';
//             DataClassification = SystemMetadata;
//         }
//         field(13; Formula3; Text[250])
//         {
//             Caption = 'Formula3';
//             DataClassification = SystemMetadata;
//         }
//         field(14; Formula4; Text[250])
//         {
//             Caption = 'Formula4';
//             DataClassification = SystemMetadata;
//         }
//         field(15; "Cell Type"; Option)
//         {
//             Caption = 'Cell Type';
//             DataClassification = SystemMetadata;
//             OptionCaption = 'Number,Text,Date,Time';
//             OptionMembers = Number,Text,Date,Time;
//         }
//         field(16; "Double Underline"; Boolean)
//         {
//             Caption = 'Double Underline';
//             DataClassification = SystemMetadata;
//         }
//         field(17; "Cell Value as Blob"; Blob)
//         {
//             Caption = 'Cell Value as Blob';
//             DataClassification = SystemMetadata;
//         }
//         field(18; "Formato Columna"; Text[250])
//         {
//             caption = 'Formato Columna';
//             DataClassification = ToBeClassified;
//         }
//         field(19; "Font Name"; Text[250])
//         {
//             caption = 'Fuente';
//             DataClassification = ToBeClassified;
//         }
//         field(20; "Font Size"; Integer)
//         {
//             caption = 'Tamaño';
//             DataClassification = ToBeClassified;
//         }
//         field(21; "Font Color"; Text[30])
//         {
//             caption = 'Color';
//             DataClassification = ToBeClassified;
//         }
//         field(22; "Background Color"; Text[30])
//         {
//             caption = 'Color Fondo';
//             DataClassification = ToBeClassified;
//         }
//         field(23; Vinculo; Text[1024])
//         {

//             DataClassification = ToBeClassified;
//         }
//         field(24; "Shet Name"; Text[250])
//         {
//             DataClassification = ToBeClassified;
//         }
//         field(25; "Sheet Name"; Text[20])
//         {
//             DataClassification = ToBeClassified;
//         }
//     }

//     keys
//     {
//         key(Key1; "Shet Name", "Row No.", "Column No.")
//         {
//             Clustered = true;
//         }
//     }

//     var
//         WrkShtHelper: DotNet WorksheetHelper;
//         FileManagement: Codeunit "File Management";
//         OpenXMLManagement: Codeunit "OpenXML Management";
//         XlWrkBkWriter: DotNet WorkbookWriter;
//         XlWrkBkReader: DotNet WorkbookReader;
//         XlWrkShtWriter: DotNet WorksheetWriter;
//         XlWrkShtReader: DotNet WorksheetReader;
//         TempInfoExcelBuf: Record "Excel Buffer 2" temporary;
//         //         RangeStartXlRow: Text[30];
//         RangeStartXlCol: Text[30];
//         RangeEndXlRow: Text[30];
//         RangeEndXlCol: Text[30];
//         FileNameServer: Text;
//         FriendlyName: Text;
//         CurrentRow: Integer;
//         CurrentCol: Integer;
//         UseInfoSheet: Boolean;
//         ErrorMessage: Text;
//         ReadDateTimeInUtcDate: Boolean;

//         Text001: Label 'Debe indicar un nombre de fichero.';
//         Text002: Label 'Debe indicar un nombre para la hoja Excel.', Comment = '{Locked="Excel"}';
//         Text003: Label 'El dichero %1 no existe.';
//         Text004: Label 'La hoja de excel %1 no existe.', Comment = '{Locked="Excel"}';
//         Text005: Label 'Creando hoja Excel...\\', Comment = '{Locked="Excel"}';
//         PageTxt: Label 'Página';
//         Text007: Label 'Leyendo Hoja de Excel...\\', Comment = '{Locked="Excel"}';
//         Text013: Label '&B';
//         Text014: Label '&D';
//         Text015: Label '&P';
//         Text016: Label 'A1';
//         Text017: Label 'SUMIF';
//         Text018: Label '#N/A';
//         Text019: Label 'GLAcc', Comment = 'Used to define an Excel range name. You must refer to Excel rules to change this term.', Locked = true;
//         Text020: Label 'Period', Comment = 'Used to define an Excel range name. You must refer to Excel rules to change this term.', Locked = true;
//         Text021: Label 'Budget';
//         Text022: Label 'CostAcc', Locked = true, Comment = 'Used to define an Excel range name. You must refer to Excel rules to change this term.';
//         Text023: Label 'Information';
//         Text034: Label 'Ficheros Excel (*.xls*)|*.xls*|All Files (*.*)|*.*', Comment = '{Split=r''\|\*\..{1,4}\|?''}{Locked="Excel"}';
//         Text035: Label 'La operación se ha canceledo.';
//         Text037: Label 'No se ha podido crear el libro.', Comment = '{Locked="Excel"}';
//         Text038: Label 'Global variable %1 is not included for test.';
//         Text039: Label 'Cell type has not been set.';
//         SavingDocumentMsg: Label 'Saving the following document: %1.';
//         ExcelFileExtensionTok: Label '.xlsx', Locked = true;
//         VmlDrawingXmlTxt: Label '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202"  path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>', Locked = true;
//         EndXmlTokenTxt: Label '</xml>', Locked = true;
//         CellNotFoundErr: Label 'Cell %1 not found.', Comment = '%1 - cell name';

//     procedure SetColumnWidth(ColName: Text[10]; NewColWidth: Decimal)
//     begin
//         if not IsNull(XlWrkShtWriter) then
//             XlWrkShtWriter.SetColumnWidth(ColName, NewColWidth);
//     end;

//     procedure UpdateBookStream(var Base64: Text; SheetName: Text; PreserveDataOnUpdate: Boolean)
//     var
//         Base64Convert: codeunit "Base64 Convert";
//         TempBlob: Codeunit "Temp Blob";
//         Filename: Text;
//         ExcelStream: OutStream;
//     begin
//         FileName := CopyStr(FileManagement.ServerTempFileName('xlsx'), 1, 250);
//         TempBlob.CreateOutStream(ExcelStream);
//         Base64Convert.FromBase64(Base64, ExcelStream);
//         FileManagement.BLOBExportToServerFile(TempBlob, Filename);
//         FileNameServer := Filename;//FileManagement.InstreamExportToServerFile(ExcelStream, 'xlsx');

//         UpdateBookExcel(FileNameServer, SheetName, PreserveDataOnUpdate);
//     end;

//     procedure CreateNewBook(SheetName: Text[250])
//     begin
//         CreateBook('', SheetName);
//     end;

//     procedure CloseBook()
//     begin
//         if not IsNull(XlWrkBkWriter) then begin
//             XlWrkBkWriter.ClearFormulaCalculations();
//             //XlWrkBkWriter.ValidateDocument();
//             XlWrkBkWriter.Close();
//             Clear(XlWrkShtWriter);
//             Clear(XlWrkBkWriter);
//         end;

//         if not IsNull(XlWrkBkReader) then begin
//             Clear(XlWrkShtReader);
//             Clear(XlWrkBkReader);
//         end;
//     end;

//     procedure SaveToStream(var ResultStream: OutStream; EraseFileAfterCompletion: Boolean)
//     var
//         TempBlob: Codeunit "Temp Blob";
//         BlobStream: InStream;
//     begin
//         FileManagement.BLOBImportFromServerFile(TempBlob, FileNameServer);
//         TempBlob.CreateInStream(BlobStream);
//         CopyStream(ResultStream, BlobStream);
//         if EraseFileAfterCompletion then
//             FILE.Erase(FileNameServer);
//     end;

//     procedure SetFriendlyFilename(Name: Text)
//     begin
//         FriendlyName := Name;
//     end;

//     procedure WriteSheet(ReportHeader: Text; CompanyName2: Text; UserID2: Text; Orientacion: Enum "Orientacion")
//     var
//         TypeHelper: Codeunit "Type Helper";
//         OrientationValues: DotNet OrientationValues;
//         XmlTextWriter: DotNet XmlTextWriter;
//         //IoFileMode: DotNet FileMode;
//         Encoding: DotNet Encoding;
//         StreamWriter: DotNet IOStreamWriter;
//         VmlDrawingPart: DotNet VmlDrawingPart;
//         StringBld: DotNet TextStringBuilder;
//         IsHandled: Boolean;
//     begin
//         if Orientacion = Orientacion::Vertical then
//             XlWrkShtWriter.AddPageSetup(OrientationValues.Portrait, 9) // 9 - default value for Paper Size - A4
//         else
//             XlWrkShtWriter.AddPageSetup(OrientationValues.Landscape, 9); // 9 - default value for Paper Size - A4
//         if ReportHeader <> '' then
//             XlWrkShtWriter.AddHeader(
//               true,
//               StrSubstNo('%1%2%1%3%4', GetExcelReference(1), ReportHeader, TypeHelper.LFSeparator(), CompanyName2));

//         XlWrkShtWriter.AddHeader(
//           false,
//           StrSubstNo('%1%3%4%3%5 %2', GetExcelReference(2), GetExcelReference(3), TypeHelper.LFSeparator(), UserID2, PageTxt));

//         IsHandled := false;
//         //OnWriteSheetOnBeforeAddAndInitializeCommentsPart(Rec, IsHandled);
//         if not IsHandled then
//             OpenXMLManagement.AddAndInitializeCommentsPart(XlWrkShtWriter, VmlDrawingPart);

//         StringBld := StringBld.StringBuilder();
//         StringBld.Append(VmlDrawingXmlTxt);

//         WriteAllToCurrentSheet(Rec);

//         StringBld.Append(EndXmlTokenTxt);

//         IsHandled := false;
//         //OnWriteSheetOnBeforeUseXmlTextWriter(Rec, IsHandled);
//         if not IsHandled then begin
//             XmlTextWriter := XmlTextWriter.XmlTextWriter(VmlDrawingPart.GetStream());
//             XmlTextWriter.WriteRaw(StringBld.ToString());
//             XmlTextWriter.Flush();
//             XmlTextWriter.Close();
//         end;

//         if UseInfoSheet then
//             if not TempInfoExcelBuf.IsEmpty() then begin
//                 SelectOrAddSheet(Text023);
//                 WriteAllToCurrentSheet(TempInfoExcelBuf);
//             end;
//     end;


//     procedure WriteAllToCurrentSheet(var ExcelBuffer: Record "Excel Buffer 2")
//     var
//         ExcelBufferDialogMgt: Codeunit "Excel Buffer Dialog Management";
//         RecNo: Integer;
//         TotalRecNo: Integer;
//         LastUpdate: DateTime;
//     begin
//         if ExcelBuffer.IsEmpty() then
//             exit;
//         ExcelBufferDialogMgt.Open(Text005);
//         LastUpdate := CurrentDateTime;
//         TotalRecNo := ExcelBuffer.Count();
//         if ExcelBuffer.FindSet() then
//             repeat
//                 RecNo := RecNo + 1;
//                 if not UpdateProgressDialog(ExcelBufferDialogMgt, LastUpdate, RecNo, TotalRecNo) then begin
//                     CloseBook();
//                     Error(Text035)
//                 end;
//                 if (ExcelBuffer.Formula = '') then
//                     WriteCellValueInternal(ExcelBuffer)
//                 else
//                     WriteCellFormula(ExcelBuffer)
//             until ExcelBuffer.Next() = 0;
//         ExcelBufferDialogMgt.Close();
//     end;

//     procedure WriteCellFormula(ExcelBuffer: Record "Excel Buffer 2")
//     var
//         Decorator: DotNet CellDecorator;
//         IsHandled: Boolean;

//     begin
//         IsHandled := false;
//         if IsHandled then
//             exit;

//         with ExcelBuffer do begin
//             GetCellDecorator(Bold, Italic, Underline, "Double Underline", "Font Name", "Font Size", "Font Color", "Background Color", Decorator);
//             if (Vinculo <> '') And (Formula <> '') then
//                 XlWrkShtWriter.AddHyperlink("Row No.", xlColID, Vinculo)
//             else begin
//                 If Vinculo <> '' then begin
//                     If StrPos(Vinculo, 'http://NAV-MALLA01:48900') <> 0 then
//                         Vinculo := 'https://bc220.malla.es/' + CopyStr(Vinculo, 26);
//                     XlWrkShtWriter.AddHyperlink("Row No.", xlColID, Vinculo);
//                     //http://NAV-MALLA01:48900 https://bc220.malla.es/
//                     // If StrPos(Vinculo, 'http://NAV-MALLA01:48900') <> 0 then
//                     //     Vinculo := 'https://bc220.malla.es/' + CopyStr(Vinculo, 26);
//                     // Case "Cell Type" of
//                     //     "Cell Type"::Number:
//                     //         Formula := StrSubstNo('HYPERLINK("%1",%2)', Vinculo, "Cell Value as Text");
//                     //     "Cell Type"::Text:
//                     //         Formula := StrSubstNo('HYPERLINK("%1","%2")', Vinculo, "Cell Value as Text");
//                     //     "Cell Type"::Date:
//                     //         Formula := StrSubstNo('HYPERLINK("%1",%2)', Vinculo, "Cell Value as Text");
//                     //     "Cell Type"::Time:
//                     //         Formula := StrSubstNo('HYPERLINK("%1",%2)', Vinculo, "Cell Value as Text");
//                     // end;

//                 end;
//             end;
//             XlWrkShtWriter.SetCellFormula("Row No.", xlColID, GetFormula(), NumberFormat, Decorator);
//         end;
//     end;

//     local procedure WriteCellValueInternal(var ExcelBuffer: Record "Excel Buffer 2")
//     var
//         Decorator: DotNet CellDecorator;
//         RecInStream: Instream;
//         StringBld: DotNet TextStringBuilder;
//         CellTextValue: Text;
//     begin
//         with ExcelBuffer do begin
//             GetCellDecorator(Bold, Italic, Underline, "Double Underline", "Font Name", "Font Size", "Font Color", "Background Color", Decorator);
//             CellTextValue := "Cell Value as Text";
//             if Vinculo <> '' then begin
//                 If StrPos(Vinculo, 'http://NAV-MALLA01:48900') <> 0 then
//                     Vinculo := 'https://bc220.malla.es/' + CopyStr(Vinculo, 26);
//                 XlWrkShtWriter.AddHyperlink("Row No.", xlColID, Vinculo);
//             end;

//             if "Cell Value as Blob".HasValue() then begin
//                 CalcFields("Cell Value as Blob");
//                 "Cell Value as Blob".CreateInStream(RecInStream, TextEncoding::Windows);
//                 RecInStream.ReadText(CellTextValue);
//             end;

//             case "Cell Type" of
//                 "Cell Type"::Number:
//                     XlWrkShtWriter.SetCellValueNumber("Row No.", xlColID, CellTextValue, NumberFormat, Decorator);
//                 "Cell Type"::Text:
//                     XlWrkShtWriter.SetCellValueText("Row No.", xlColID, CellTextValue, Decorator);
//                 "Cell Type"::Date:
//                     XlWrkShtWriter.SetCellValueDate("Row No.", xlColID, CellTextValue, NumberFormat, Decorator);
//                 "Cell Type"::Time:
//                     XlWrkShtWriter.SetCellValueTime("Row No.", xlColID, CellTextValue, NumberFormat, Decorator);
//                 else
//                     Error(Text039)
//             end;

//             if Comment <> '' then begin
//                 OpenXMLManagement.SetCellComment(XlWrkShtWriter, StrSubstNo('%1%2', xlColID, "Row No."), Comment);
//                 StringBld.Append(OpenXMLManagement.CreateCommentVmlShapeXml("Column No.", "Row No."));
//             end;
//         end;
//     end;

//     local procedure UpdateProgressDialog(var ExcelBufferDialogManagement: Codeunit "Excel Buffer Dialog Management"; var LastUpdate: DateTime; CurrentCount: Integer; TotalCount: Integer): Boolean
//     var
//         CurrentTime: DateTime;
//     begin
//         // Refresh at 100%, and every second in between 0% to 100%
//         // Duration is measured in miliseconds -> 1 sec = 1000 ms
//         CurrentTime := CurrentDateTime;
//         if (CurrentCount = TotalCount) or (CurrentTime - LastUpdate >= 1000) then begin
//             LastUpdate := CurrentTime;
//             if not ExcelBufferDialogManagement.SetProgress(Round(CurrentCount / TotalCount * 10000, 1)) then
//                 exit(false);
//         end;

//         exit(true)
//     end;

//     procedure SelectOrAddSheet(NewSheetName: Text)
//     begin
//         if NewSheetName = '' then
//             exit;
//         if IsNull(XlWrkBkWriter) then
//             Error(Text037);
//         if XlWrkBkWriter.HasWorksheet(NewSheetName) then
//             XlWrkShtWriter := XlWrkBkWriter.GetWorksheetByName(NewSheetName)
//         else
//             XlWrkShtWriter := XlWrkBkWriter.AddWorksheet(NewSheetName);
//     end;

//     procedure GetExcelReference(Which: Integer): Text[250]
//     begin
//         case Which of
//             1:
//                 exit(Text013);
//             // DO NOT TRANSLATE: &B is the Excel code to turn bold printing on or off for customized Header/Footer.
//             2:
//                 exit(Text014);
//             // DO NOT TRANSLATE: &D is the Excel code to print the current date in customized Header/Footer.
//             3:
//                 exit(Text015);
//             // DO NOT TRANSLATE: &P is the Excel code to print the page number in customized Header/Footer.
//             4:
//                 exit('$');
//             // DO NOT TRANSLATE: $ is the Excel code for absolute reference to cells.
//             5:
//                 exit(Text016);
//             // DO NOT TRANSLATE: A1 is the Excel reference of the first cell.
//             6:
//                 exit(Text017);
//             // DO NOT TRANSLATE: SUMIF is the name of the Excel function used to summarize values according to some conditions.
//             7:
//                 exit(Text018);
//             // DO NOT TRANSLATE: The #N/A Excel error value occurs when a value is not available to a function or formula.
//             8:
//                 exit(Text019);
//             // DO NOT TRANSLATE: GLAcc is used to define an Excel range name. You must refer to Excel rules to change this term.
//             9:
//                 exit(Text020);
//             // DO NOT TRANSLATE: Period is used to define an Excel range name. You must refer to Excel rules to change this term.
//             10:
//                 exit(Text021);
//             // DO NOT TRANSLATE: Budget is used to define an Excel worksheet name. You must refer to Excel rules to change this term.
//             11:
//                 exit(Text022);
//         // DO NOT TRANSLATE: CostAcc is used to define an Excel range name. You must refer to Excel rules to change this term.
//         end;
//     end;

//     [Scope('OnPrem')]
//     procedure CreateBook(FileName: Text; SheetName: Text)
//     begin
//         if SheetName = '' then
//             Error(Text002);

//         if FileName = '' then
//             FileNameServer := FileManagement.ServerTempFileName('xlsx')
//         else begin
//             if Exists(FileName) then
//                 Erase(FileName);
//             FileNameServer := FileName;
//         end;

//         FileManagement.IsAllowedPath(FileNameServer, false);
//         XlWrkBkWriter := XlWrkBkWriter.Create(FileNameServer);
//         if IsNull(XlWrkBkWriter) then
//             Error(Text037);

//         XlWrkShtWriter := XlWrkBkWriter.FirstWorksheet;
//         if SheetName <> '' then
//             XlWrkShtWriter.Name := SheetName;

//         OpenXMLManagement.SetupWorksheetHelper(XlWrkBkWriter);
//     end;

//     procedure GetValueByCellName(CellName: Text): Text
//     var
//         CellPosition: DotNet CellPosition;
//         RowInt: Integer;
//         ColumnInt: Integer;
//     begin
//         CellPosition := CellPosition.CellPosition(CellName);
//         RowInt := CellPosition.Row;
//         ColumnInt := CellPosition.Column;
//         if Get(RowInt, ColumnInt) then
//             exit("Cell Value as Text");
//     end;

//     [Scope('OnPrem')]
//     procedure UpdateBookExcel(FileName: Text; SheetName: Text; PreserveDataOnUpdate: Boolean)
//     begin
//         if FileName = '' then
//             Error(Text001);

//         if SheetName = '' then
//             Error(Text002);

//         FileNameServer := FileName;
//         FileManagement.IsAllowedPath(FileName, false);
//         XlWrkBkWriter := XlWrkBkWriter.Open(FileNameServer);
//         if XlWrkBkWriter.HasWorksheet(SheetName) then begin
//             XlWrkShtWriter := XlWrkBkWriter.GetWorksheetByName(SheetName);
//             // Set PreserverDataOnUpdate to false if the sheet writer should clear all empty cells
//             // in which NAV does not have new data. Notice that the sheet writer will only clear Excel
//             // cells that are addressed by the writer. All other cells will be left unmodified.
//             XlWrkShtWriter.PreserveDataOnUpdate := PreserveDataOnUpdate;

//             OpenXMLManagement.SetupWorksheetHelper(XlWrkBkWriter);
//         end else begin
//             CloseBook();
//             Error(Text004, SheetName);
//         end;
//     end;

//     local procedure GetCellDecorator(IsBold: Boolean; IsItalic: Boolean; IsUnderlined: Boolean; IsDoubleUnderlined: Boolean;
//     FontName: Text; FontSize: Integer; FontColor: Text; BackgroundColor: Text; var Decorator: DotNet CellDecorator)
//     var
//         Fill: Dotnet "Fill";
//         PatterFill: Dotnet PatternFill;
//         DotnetBackgroundColor: Dotnet BackgroundColor;
//         DotnetForegroundColor: Dotnet ForegroundColor;
//         HexBinaryValue: Dotnet HexBinaryValue;
//         PatternValues: Dotnet PatternValues;
//         Map: DotNet Map;
//         //Enum: Dotnet Enum;
//         UInt32Value: DotNet UInt32Value;
//         Double: Dotnet DoubleValue;
//         DotnetFontSize: DotNet FontSize;
//         DotnetFont: DotNet Font;
//         DotNetFontName: DotNet FontName;
//         DotnetFontColor: DotNet FontColor;
//         StringValue: DotNet StringValue;
//         ownerXML: DotNet SystemString;
//     begin

//         if IsBold and IsItalic then begin
//             if IsDoubleUnderlined then begin
//                 Decorator := XlWrkShtWriter.DefaultBoldItalicDoubleUnderlinedCellDecorator;
//                 exit;
//             end;
//             if IsUnderlined then begin
//                 Decorator := XlWrkShtWriter.DefaultBoldItalicUnderlinedCellDecorator;
//                 exit;
//             end;
//         end;

//         if IsBold and IsItalic then begin
//             Decorator := XlWrkShtWriter.DefaultBoldItalicCellDecorator;
//             exit;
//         end;
//         if IsBold then begin
//             if IsDoubleUnderlined then begin
//                 Decorator := XlWrkShtWriter.DefaultBoldDoubleUnderlinedCellDecorator;
//                 exit;
//             end;
//             if IsUnderlined then begin
//                 Decorator := XlWrkShtWriter.DefaultBoldUnderlinedCellDecorator;
//                 exit;
//             end;
//         end;

//         if IsBold then begin
//             Decorator := XlWrkShtWriter.DefaultBoldCellDecorator;
//             exit;
//         end;

//         if IsItalic then begin
//             if IsDoubleUnderlined then begin
//                 Decorator := XlWrkShtWriter.DefaultItalicDoubleUnderlinedCellDecorator;
//                 exit;
//             end;
//             if IsUnderlined then begin
//                 Decorator := XlWrkShtWriter.DefaultItalicUnderlinedCellDecorator;
//                 exit;
//             end;
//         end;

//         if IsItalic then begin
//             Decorator := XlWrkShtWriter.DefaultItalicCellDecorator;
//             exit;
//         end;

//         if IsDoubleUnderlined then
//             Decorator := XlWrkShtWriter.DefaultDoubleUnderlinedCellDecorator
//         else
//             if IsUnderlined then
//                 Decorator := XlWrkShtWriter.DefaultUnderlinedCellDecorator
//             else
//                 Decorator := XlWrkShtWriter.DefaultCellDecorator;
//         //error(Decorator.Fill.PatternFill.BackgroundColor.Rgb.HexBinaryValue().ToString());
//         If (FontColor <> '') Or (FontName <> '') or (FontSize <> 0) then begin
//             DotnetFont := DotnetFont.Font();
//             If FontColor <> '' then begin
//                 HexBinaryValue := HexBinaryValue.HexBinaryValue();
//                 HexBinaryValue.Value := FontColor;
//                 DotnetFontColor := DotNetFontColor.Color();
//                 DotnetFontColor.Rgb := HexBinaryValue;
//                 DotnetFont.Color := DotnetFontColor;
//             end;
//             If FontName <> '' then begin
//                 DotNetFontName := DotNetFontName.FontName();
//                 StringValue := StringValue.StringValue();
//                 StringValue.Value := FontName;
//                 DotNetFontName.Val := StringValue;
//                 DotnetFont.FontName := DotNetFontName;
//             end;
//             If FontSize <> 0 then begin
//                 DotnetFontSize := DotnetFontSize.FontSize();
//                 Double := Double.DoubleValue();
//                 Double.Value := FontSize;
//                 DotnetFontSize.Val := Double;
//                 DotnetFont.FontSize := DotnetFontSize;
//             end;
//             Decorator.Font := DotNetFont;
//         end;

//         If BackgroundColor <> '' then begin
//             Fill := Decorator.Fill.CloneNode(true);
//             ownerXML := '<x:patternFill xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" patternType="solid">' + '<x:fgColor rgb="' + BackgroundColor + '" /></x:patternFill>';
//             PatterFill := PatterFill.PatternFill(ownerXML);
//             Fill.PatternFill := PatterFill;
//             Decorator.Fill := Fill;
//         end;


//     end;

//     procedure GetFormula(): Text[1000]
//     begin
//         exit(Formula + Formula2 + Formula3 + Formula4);
//     end;

//     procedure SetFormula(LongFormula: Text[1000])
//     begin
//         ClearFormula();
//         if LongFormula = '' then
//             exit;

//         Formula := CopyStr(LongFormula, 1, MaxStrLen(Formula));
//         if StrLen(LongFormula) > MaxStrLen(Formula) then
//             Formula2 := CopyStr(LongFormula, MaxStrLen(Formula) + 1, MaxStrLen(Formula2));
//         if StrLen(LongFormula) > MaxStrLen(Formula) + MaxStrLen(Formula2) then
//             Formula3 := CopyStr(LongFormula, MaxStrLen(Formula) + MaxStrLen(Formula2) + 1, MaxStrLen(Formula3));
//         if StrLen(LongFormula) > MaxStrLen(Formula) + MaxStrLen(Formula2) + MaxStrLen(Formula3) then
//             Formula4 := CopyStr(LongFormula, MaxStrLen(Formula) + MaxStrLen(Formula2) + MaxStrLen(Formula3) + 1, MaxStrLen(Formula4));
//     end;

//     procedure ClearFormula()
//     begin
//         Formula := '';
//         Formula2 := '';
//         Formula3 := '';
//         Formula4 := '';
//     end;
// }