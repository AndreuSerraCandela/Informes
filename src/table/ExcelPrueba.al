table 7001239 "Excel Buffer 2"
{
    Caption = 'Excel Buffer';
    ReplicateData = false;

    fields
    {
        field(1; "Row No."; Integer)
        {
            Caption = 'Row No.';
            DataClassification = SystemMetadata;

            trigger OnValidate()
            begin
                xlRowID := '';
                if "Row No." <> 0 then
                    xlRowID := Format("Row No.");
            end;
        }
        field(2; xlRowID; Text[10])
        {
            Caption = 'xlRowID';
            DataClassification = SystemMetadata;
        }
        field(3; "Column No."; Integer)
        {
            Caption = 'Column No.';
            DataClassification = SystemMetadata;

            trigger OnValidate()
            var
                x: Integer;
                i: Integer;
                y: Integer;
                c: Char;
                t: Text[30];
            begin
                xlColID := '';
                x := "Column No.";
                while x > 26 do begin
                    y := x mod 26;
                    if y = 0 then
                        y := 26;
                    c := 64 + y;
                    i := i + 1;
                    t[i] := c;
                    x := (x - y) div 26;
                end;
                if x > 0 then begin
                    c := 64 + x;
                    i := i + 1;
                    t[i] := c;
                end;
                for x := 1 to i do
                    xlColID[x] := t[1 + i - x];
            end;
        }
        field(4; xlColID; Text[10])
        {
            Caption = 'xlColID';
            DataClassification = SystemMetadata;
        }
        field(5; "Cell Value as Text"; Text[250])
        {
            Caption = 'Cell Value as Text';
            DataClassification = SystemMetadata;
        }
        field(6; Comment; Text[250])
        {
            Caption = 'Comment';
            DataClassification = SystemMetadata;
        }
        field(7; Formula; Text[250])
        {
            Caption = 'Formula';
            DataClassification = SystemMetadata;
        }
        field(8; Bold; Boolean)
        {
            Caption = 'Bold';
            DataClassification = SystemMetadata;
        }
        field(9; Italic; Boolean)
        {
            Caption = 'Italic';
            DataClassification = SystemMetadata;
        }
        field(10; Underline; Boolean)
        {
            Caption = 'Underline';
            DataClassification = SystemMetadata;
        }
        field(11; NumberFormat; Text[250])
        {
            Caption = 'NumberFormat';
            DataClassification = SystemMetadata;
        }
        field(12; Formula2; Text[250])
        {
            Caption = 'Formula2';
            DataClassification = SystemMetadata;
        }
        field(13; Formula3; Text[250])
        {
            Caption = 'Formula3';
            DataClassification = SystemMetadata;
        }
        field(14; Formula4; Text[250])
        {
            Caption = 'Formula4';
            DataClassification = SystemMetadata;
        }
        field(15; "Cell Type"; Option)
        {
            Caption = 'Cell Type';
            DataClassification = SystemMetadata;
            OptionCaption = 'Number,Text,Date,Time';
            OptionMembers = Number,Text,Date,Time;
        }
        field(16; "Double Underline"; Boolean)
        {
            Caption = 'Double Underline';
            DataClassification = SystemMetadata;
        }
        field(17; "Cell Value as Blob"; Blob)
        {
            Caption = 'Cell Value as Blob';
            DataClassification = SystemMetadata;
        }
        field(18; "Formato Columna"; Text[250])
        {
            caption = 'Formato Columna';
            DataClassification = ToBeClassified;
        }
        field(19; "Font Name"; Text[250])
        {
            caption = 'Fuente';
            DataClassification = ToBeClassified;
        }
        field(20; "Font Size"; Integer)
        {
            caption = 'Tamaño';
            DataClassification = ToBeClassified;
        }
        field(21; "Font Color"; Text[30])
        {
            caption = 'Color';
            DataClassification = ToBeClassified;
        }
        field(22; "Background Color"; Text[30])
        {
            caption = 'Color Fondo';
            DataClassification = ToBeClassified;
        }
        field(23; Vinculo; Text[1024])
        {

            DataClassification = ToBeClassified;
        }
        field(24; "Shet Name"; Text[250])
        {
            DataClassification = ToBeClassified;
        }
        field(25; "Sheet Name"; Text[20])
        {
            DataClassification = ToBeClassified;
        }
    }

    keys
    {
        key(Key1; "Shet Name", "Row No.", "Column No.")
        {
            Clustered = true;
        }
    }

    fieldgroups
    {
    }

    var
        WrkShtHelper: DotNet WorksheetHelper;
        TempInfoExcelBuf: Record "Excel Buffer 2" temporary;
        FileManagement: Codeunit "File Management";
        OpenXMLManagement: Codeunit "OpenXML Management";
        XlWrkBkWriter: DotNet WorkbookWriter;
        XlWrkBkReader: DotNet WorkbookReader;
        XlWrkShtWriter: DotNet WorksheetWriter;
        XlWrkShtReader: DotNet WorksheetReader;
        StringBld: DotNet StringBuilder;
        RangeStartXlRow: Text[30];
        RangeStartXlCol: Text[30];
        RangeEndXlRow: Text[30];
        RangeEndXlCol: Text[30];
        FileNameServer: Text;
        FriendlyName: Text;
        CurrentRow: Integer;
        CurrentCol: Integer;
        UseInfoSheet: Boolean;
        ErrorMessage: Text;
        ReadDateTimeInUtcDate: Boolean;

        Text001: Label 'Debe indicar un nombre de fichero.';
        Text002: Label 'Debe indicar un nombre para la hoja Excel.', Comment = '{Locked="Excel"}';
        Text003: Label 'El dichero %1 no existe.';
        Text004: Label 'La hoja de excel %1 no existe.', Comment = '{Locked="Excel"}';
        Text005: Label 'Creando hoja Excel...\\', Comment = '{Locked="Excel"}';
        PageTxt: Label 'Pájina';
        Text007: Label 'Leyendo Hoja de Excel...\\', Comment = '{Locked="Excel"}';
        Text013: Label '&B';
        Text014: Label '&D';
        Text015: Label '&P';
        Text016: Label 'A1';
        Text017: Label 'SUMIF';
        Text018: Label '#N/A';
        Text019: Label 'GLAcc', Comment = 'Used to define an Excel range name. You must refer to Excel rules to change this term.', Locked = true;
        Text020: Label 'Period', Comment = 'Used to define an Excel range name. You must refer to Excel rules to change this term.', Locked = true;
        Text021: Label 'Budget';
        Text022: Label 'CostAcc', Locked = true, Comment = 'Used to define an Excel range name. You must refer to Excel rules to change this term.';
        Text023: Label 'Information';
        Text034: Label 'Ficheros Excel (*.xls*)|*.xls*|All Files (*.*)|*.*', Comment = '{Split=r''\|\*\..{1,4}\|?''}{Locked="Excel"}';
        Text035: Label 'La operación se ha canceledo.';
        Text037: Label 'No se ha podido crear el libro.', Comment = '{Locked="Excel"}';
        Text038: Label 'Global variable %1 is not included for test.';
        Text039: Label 'Cell type has not been set.';
        SavingDocumentMsg: Label 'Saving the following document: %1.';
        ExcelFileExtensionTok: Label '.xlsx', Locked = true;
        VmlDrawingXmlTxt: Label '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202"  path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>', Locked = true;
        EndXmlTokenTxt: Label '</xml>', Locked = true;
        CellNotFoundErr: Label 'Cell %1 not found.', Comment = '%1 - cell name';

    procedure SetReadDateTimeInUtcDate(NewValue: Boolean)
    begin
        ReadDateTimeInUtcDate := NewValue;
    end;



    procedure CreateNewBook(SheetName: Text[250])
    begin
        CreateBook('', SheetName);
    end;

    [Scope('OnPrem')]
    procedure CreateBook(FileName: Text; SheetName: Text)
    begin
        if SheetName = '' then
            Error(Text002);

        if FileName = '' then
            FileNameServer := FileManagement.ServerTempFileName('xlsx')
        else begin
            if Exists(FileName) then
                Erase(FileName);
            FileNameServer := FileName;
        end;

        FileManagement.IsAllowedPath(FileNameServer, false);
        XlWrkBkWriter := XlWrkBkWriter.Create(FileNameServer);
        if IsNull(XlWrkBkWriter) then
            Error(Text037);

        XlWrkShtWriter := XlWrkBkWriter.FirstWorksheet;
        if SheetName <> '' then
            XlWrkShtWriter.Name := SheetName;

        OpenXMLManagement.SetupWorksheetHelper(XlWrkBkWriter);
    end;

    procedure GetValueByCellName(CellName: Text): Text
    var
        CellPosition: DotNet CellPosition;
        RowInt: Integer;
        ColumnInt: Integer;
    begin
        CellPosition := CellPosition.CellPosition(CellName);
        RowInt := CellPosition.Row;
        ColumnInt := CellPosition.Column;
        if Get(RowInt, ColumnInt) then
            exit("Cell Value as Text");
    end;

    [Scope('OnPrem')]
    procedure GetNextColumnCellName(CellName: Text[30]): Text
    begin
        if not IsNull(XlWrkShtReader) then
            exit(XlWrkShtReader.GetNextColumnCellName(CellName));

        exit(XlWrkShtWriter.GetNextColumnCellName(CellName));
    end;

    [Scope('OnPrem')]
    procedure GetNextRowCellName(CellName: Text[30]): Text
    begin
        if not IsNull(XlWrkShtReader) then
            exit(XlWrkShtReader.GetNextRowCellName(CellName));

        exit(XlWrkShtWriter.GetNextRowCellName(CellName));
    end;

    [Scope('OnPrem')]
    procedure OpenBook(FileName: Text; SheetName: Text)
    begin
        if FileName = '' then
            Error(Text001);

        if SheetName = '' then
            Error(Text002);

        if SheetName = 'G/L Account' then
            SheetName := 'GL Account';

        FileManagement.IsAllowedPath(FileName, false);
        XlWrkBkReader := XlWrkBkReader.Open(FileName);
        if XlWrkBkReader.HasWorksheet(SheetName) then
            XlWrkShtReader := XlWrkBkReader.GetWorksheetByName(SheetName)
        else begin
            CloseBook();
            Error(Text004, SheetName);
        end;
    end;

    procedure OpenBookStream(FileStream: InStream; SheetName: Text): Text
    begin
        if SheetName = '' then
            exit(Text002);

        if SheetName = 'G/L Account' then
            SheetName := 'GL Account';

        XlWrkBkReader := XlWrkBkReader.Open(FileStream);
        if XlWrkBkReader.HasWorksheet(SheetName) then
            XlWrkShtReader := XlWrkBkReader.GetWorksheetByName(SheetName)
        else begin
            CloseBook();
            ErrorMessage := StrSubstNo(Text004, SheetName);
            exit(ErrorMessage);
        end;
    end;

    [Scope('OnPrem')]
    procedure OpenBookForUpdate(FileName: Text)
    begin
        FileNameServer := FileName;
        FileManagement.IsAllowedPath(FileName, false);
        XlWrkBkWriter := XlWrkBkWriter.Open(FileNameServer);
    end;

    [Scope('OnPrem')]
    procedure UpdateBook(FileName: Text; SheetName: Text)
    begin
        UpdateBookExcel(FileName, SheetName, true);
    end;

    [Scope('OnPrem')]
    procedure UpdateBookExcel(FileName: Text; SheetName: Text; PreserveDataOnUpdate: Boolean)
    begin
        if FileName = '' then
            Error(Text001);

        if SheetName = '' then
            Error(Text002);

        FileNameServer := FileName;
        FileManagement.IsAllowedPath(FileName, false);
        XlWrkBkWriter := XlWrkBkWriter.Open(FileNameServer);
        if XlWrkBkWriter.HasWorksheet(SheetName) then begin
            XlWrkShtWriter := XlWrkBkWriter.GetWorksheetByName(SheetName);
            // Set PreserverDataOnUpdate to false if the sheet writer should clear all empty cells
            // in which NAV does not have new data. Notice that the sheet writer will only clear Excel
            // cells that are addressed by the writer. All other cells will be left unmodified.
            XlWrkShtWriter.PreserveDataOnUpdate := PreserveDataOnUpdate;

            OpenXMLManagement.SetupWorksheetHelper(XlWrkBkWriter);
        end else begin
            CloseBook();
            Error(Text004, SheetName);
        end;
    end;

    procedure UpdateBookStream(var ExcelStream: InStream; SheetName: Text; PreserveDataOnUpdate: Boolean)
    begin
        FileNameServer := FileManagement.InstreamExportToServerFile(ExcelStream, 'xlsx');

        UpdateBookExcel(FileNameServer, SheetName, PreserveDataOnUpdate);
    end;

    procedure CloseBook()
    begin
        if not IsNull(XlWrkBkWriter) then begin
            XlWrkBkWriter.ClearFormulaCalculations();
            XlWrkBkWriter.ValidateDocument();
            XlWrkBkWriter.Close();
            Clear(XlWrkShtWriter);
            Clear(XlWrkBkWriter);
        end;

        if not IsNull(XlWrkBkReader) then begin
            Clear(XlWrkShtReader);
            Clear(XlWrkBkReader);
        end;
    end;

    procedure SelectOrAddSheet(NewSheetName: Text)
    begin
        if NewSheetName = '' then
            exit;
        if IsNull(XlWrkBkWriter) then
            Error(Text037);
        if XlWrkBkWriter.HasWorksheet(NewSheetName) then
            XlWrkShtWriter := XlWrkBkWriter.GetWorksheetByName(NewSheetName)
        else
            XlWrkShtWriter := XlWrkBkWriter.AddWorksheet(NewSheetName);
    end;

    procedure SetActiveReaderSheet(NewSheetName: Text)
    begin
        if NewSheetName = '' then
            exit;

        if XlWrkBkReader.HasWorksheet(NewSheetName) then
            XlWrkShtReader := XlWrkBkReader.GetWorksheetByName(NewSheetName)
        else begin
            CloseBook();
            Error(Text004, NewSheetName);
        end;
    end;

    procedure WriteSheet(ReportHeader: Text; CompanyName2: Text; UserID2: Text; Orientacion: Enum "Orientacion")
    var
        TypeHelper: Codeunit "Type Helper";
        OrientationValues: DotNet OrientationValues;
        XmlTextWriter: DotNet XmlTextWriter;
        FileMode: DotNet FileMode;
        Encoding: DotNet Encoding;
        VmlDrawingPart: DotNet VmlDrawingPart;
        IsHandled: Boolean;
    begin
        if Orientacion = Orientacion::Vertical then
            XlWrkShtWriter.AddPageSetup(OrientationValues.Portrait, 9) // 9 - default value for Paper Size - A4
        else
            XlWrkShtWriter.AddPageSetup(OrientationValues.Landscape, 9); // 9 - default value for Paper Size - A4
        if ReportHeader <> '' then
            XlWrkShtWriter.AddHeader(
              true,
              StrSubstNo('%1%2%1%3%4', GetExcelReference(1), ReportHeader, TypeHelper.LFSeparator(), CompanyName2));

        XlWrkShtWriter.AddHeader(
          false,
          StrSubstNo('%1%3%4%3%5 %2', GetExcelReference(2), GetExcelReference(3), TypeHelper.LFSeparator(), UserID2, PageTxt));

        IsHandled := false;
        OnWriteSheetOnBeforeAddAndInitializeCommentsPart(Rec, IsHandled);
        if not IsHandled then
            OpenXMLManagement.AddAndInitializeCommentsPart(XlWrkShtWriter, VmlDrawingPart);

        StringBld := StringBld.StringBuilder();
        StringBld.Append(VmlDrawingXmlTxt);

        WriteAllToCurrentSheet(Rec);

        StringBld.Append(EndXmlTokenTxt);

        IsHandled := false;
        OnWriteSheetOnBeforeUseXmlTextWriter(Rec, IsHandled);
        if not IsHandled then begin
            XmlTextWriter := XmlTextWriter.XmlTextWriter(VmlDrawingPart.GetStream(FileMode.Create), Encoding.UTF8);
            XmlTextWriter.WriteRaw(StringBld.ToString());
            XmlTextWriter.Flush();
            XmlTextWriter.Close();
        end;

        if UseInfoSheet then
            if not TempInfoExcelBuf.IsEmpty() then begin
                SelectOrAddSheet(Text023);
                WriteAllToCurrentSheet(TempInfoExcelBuf);
            end;
    end;

    procedure WriteAllToCurrentSheet(var ExcelBuffer: Record "Excel Buffer 2")
    var
        ExcelBufferDialogMgt: Codeunit "Excel Buffer Dialog Management";
        RecNo: Integer;
        TotalRecNo: Integer;
        LastUpdate: DateTime;
    begin
        if ExcelBuffer.IsEmpty() then
            exit;
        ExcelBufferDialogMgt.Open(Text005);
        LastUpdate := CurrentDateTime;
        TotalRecNo := ExcelBuffer.Count();
        if ExcelBuffer.FindSet() then
            repeat
                RecNo := RecNo + 1;
                if not UpdateProgressDialog(ExcelBufferDialogMgt, LastUpdate, RecNo, TotalRecNo) then begin
                    CloseBook();
                    Error(Text035)
                end;
                if (ExcelBuffer.Formula = '') then
                    WriteCellValueInternal(ExcelBuffer)
                else
                    WriteCellFormula(ExcelBuffer)
            until ExcelBuffer.Next() = 0;
        ExcelBufferDialogMgt.Close();
    end;

    procedure WriteCellValue(ExcelBuffer: Record "Excel Buffer 2")
    begin
        WriteCellValueInternal(ExcelBuffer);
    end;

    local procedure WriteCellValueInternal(var ExcelBuffer: Record "Excel Buffer 2")
    var
        Decorator: DotNet CellDecorator;
        RecInStream: Instream;
        CellTextValue: Text;
    begin
        with ExcelBuffer do begin
            GetCellDecorator(Bold, Italic, Underline, "Double Underline", "Font Name", "Font Size", "Font Color", "Background Color", Decorator);
            CellTextValue := "Cell Value as Text";
            if Vinculo <> '' then begin
                If StrPos(Vinculo, 'http://NAV-MALLA01:48900') <> 0 then
                    Vinculo := 'https://bc220.malla.es/' + CopyStr(Vinculo, 26);
                XlWrkShtWriter.AddHyperlink("Row No.", xlColID, Vinculo);
            end;

            if "Cell Value as Blob".HasValue() then begin
                CalcFields("Cell Value as Blob");
                "Cell Value as Blob".CreateInStream(RecInStream, TextEncoding::Windows);
                RecInStream.ReadText(CellTextValue);
            end;

            OnWriteCellValueOnBeforeSetCellValue(Rec, CellTextValue);
            case "Cell Type" of
                "Cell Type"::Number:
                    XlWrkShtWriter.SetCellValueNumber("Row No.", xlColID, CellTextValue, NumberFormat, Decorator);
                "Cell Type"::Text:
                    XlWrkShtWriter.SetCellValueText("Row No.", xlColID, CellTextValue, Decorator);
                "Cell Type"::Date:
                    XlWrkShtWriter.SetCellValueDate("Row No.", xlColID, CellTextValue, NumberFormat, Decorator);
                "Cell Type"::Time:
                    XlWrkShtWriter.SetCellValueTime("Row No.", xlColID, CellTextValue, NumberFormat, Decorator);
                else
                    Error(Text039)
            end;

            if Comment <> '' then begin
                OpenXMLManagement.SetCellComment(XlWrkShtWriter, StrSubstNo('%1%2', xlColID, "Row No."), Comment);
                StringBld.Append(OpenXMLManagement.CreateCommentVmlShapeXml("Column No.", "Row No."));
            end;
        end;
    end;

    procedure WriteCellFormula(ExcelBuffer: Record "Excel Buffer 2")
    var
        Decorator: DotNet CellDecorator;
        IsHandled: Boolean;

    begin
        IsHandled := false;
        OnBeforeWriteCellFormula(Rec, ExcelBuffer, IsHandled);
        if IsHandled then
            exit;

        with ExcelBuffer do begin
            GetCellDecorator(Bold, Italic, Underline, "Double Underline", "Font Name", "Font Size", "Font Color", "Background Color", Decorator);
            if (Vinculo <> '') And (Formula <> '') then
                XlWrkShtWriter.AddHyperlink("Row No.", xlColID, Vinculo)
            else begin
                If Vinculo <> '' then begin
                    If StrPos(Vinculo, 'http://NAV-MALLA01:48900') <> 0 then
                        Vinculo := 'https://bc220.malla.es/' + CopyStr(Vinculo, 26);
                    XlWrkShtWriter.AddHyperlink("Row No.", xlColID, Vinculo);
                    //http://NAV-MALLA01:48900 https://bc220.malla.es/
                    // If StrPos(Vinculo, 'http://NAV-MALLA01:48900') <> 0 then
                    //     Vinculo := 'https://bc220.malla.es/' + CopyStr(Vinculo, 26);
                    // Case "Cell Type" of
                    //     "Cell Type"::Number:
                    //         Formula := StrSubstNo('HYPERLINK("%1",%2)', Vinculo, "Cell Value as Text");
                    //     "Cell Type"::Text:
                    //         Formula := StrSubstNo('HYPERLINK("%1","%2")', Vinculo, "Cell Value as Text");
                    //     "Cell Type"::Date:
                    //         Formula := StrSubstNo('HYPERLINK("%1",%2)', Vinculo, "Cell Value as Text");
                    //     "Cell Type"::Time:
                    //         Formula := StrSubstNo('HYPERLINK("%1",%2)', Vinculo, "Cell Value as Text");
                    // end;

                end;
            end;
            XlWrkShtWriter.SetCellFormula("Row No.", xlColID, GetFormula(), NumberFormat, Decorator);
        end;
    end;

    local procedure GetCellDecorator(IsBold: Boolean; IsItalic: Boolean; IsUnderlined: Boolean; IsDoubleUnderlined: Boolean;
    FontName: Text; FontSize: Integer; FontColor: Text; BackgroundColor: Text; var Decorator: DotNet CellDecorator)
    var
        Fill: Dotnet "Fill";
        PatterFill: Dotnet PatternFill;
        DotnetBackgroundColor: Dotnet BackgroundColor;
        DotnetForegroundColor: Dotnet ForegroundColor;
        HexBinaryValue: Dotnet HexBinaryValue;
        PatternValues: Dotnet PatternValues;
        Map: DotNet Map;
        Enum: Dotnet Enum;
        UInt32Value: DotNet UInt32Value;
        Double: Dotnet DoubleValue;
        DotnetFontSize: DotNet FontSize;
        DotnetFont: DotNet Font;
        DotNetFontName: DotNet FontName;
        DotnetFontColor: DotNet FontColor;
        StringValue: DotNet StringValue;
        ownerXML: DotNet String;
    begin

        if IsBold and IsItalic then begin
            if IsDoubleUnderlined then begin
                Decorator := XlWrkShtWriter.DefaultBoldItalicDoubleUnderlinedCellDecorator;
                exit;
            end;
            if IsUnderlined then begin
                Decorator := XlWrkShtWriter.DefaultBoldItalicUnderlinedCellDecorator;
                exit;
            end;
        end;

        if IsBold and IsItalic then begin
            Decorator := XlWrkShtWriter.DefaultBoldItalicCellDecorator;
            exit;
        end;
        if IsBold then begin
            if IsDoubleUnderlined then begin
                Decorator := XlWrkShtWriter.DefaultBoldDoubleUnderlinedCellDecorator;
                exit;
            end;
            if IsUnderlined then begin
                Decorator := XlWrkShtWriter.DefaultBoldUnderlinedCellDecorator;
                exit;
            end;
        end;

        if IsBold then begin
            Decorator := XlWrkShtWriter.DefaultBoldCellDecorator;
            exit;
        end;

        if IsItalic then begin
            if IsDoubleUnderlined then begin
                Decorator := XlWrkShtWriter.DefaultItalicDoubleUnderlinedCellDecorator;
                exit;
            end;
            if IsUnderlined then begin
                Decorator := XlWrkShtWriter.DefaultItalicUnderlinedCellDecorator;
                exit;
            end;
        end;

        if IsItalic then begin
            Decorator := XlWrkShtWriter.DefaultItalicCellDecorator;
            exit;
        end;

        if IsDoubleUnderlined then
            Decorator := XlWrkShtWriter.DefaultDoubleUnderlinedCellDecorator
        else
            if IsUnderlined then
                Decorator := XlWrkShtWriter.DefaultUnderlinedCellDecorator
            else
                Decorator := XlWrkShtWriter.DefaultCellDecorator;
        //error(Decorator.Fill.PatternFill.BackgroundColor.Rgb.HexBinaryValue().ToString());
        If (FontColor <> '') Or (FontName <> '') or (FontSize <> 0) then begin
            DotnetFont := DotnetFont.Font();
            If FontColor <> '' then begin
                HexBinaryValue := HexBinaryValue.HexBinaryValue();
                HexBinaryValue.Value := FontColor;
                DotnetFontColor := DotNetFontColor.Color();
                DotnetFontColor.Rgb := HexBinaryValue;
                DotnetFont.Color := DotnetFontColor;
            end;
            If FontName <> '' then begin
                DotNetFontName := DotNetFontName.FontName();
                StringValue := StringValue.StringValue();
                StringValue.Value := FontName;
                DotNetFontName.Val := StringValue;
                DotnetFont.FontName := DotNetFontName;
            end;
            If FontSize <> 0 then begin
                DotnetFontSize := DotnetFontSize.FontSize();
                Double := Double.DoubleValue();
                Double.Value := FontSize;
                DotnetFontSize.Val := Double;
                DotnetFont.FontSize := DotnetFontSize;
            end;
            Decorator.Font := DotNetFont;
        end;

        If BackgroundColor <> '' then begin
            Fill := Decorator.Fill.CloneNode(true);
            ownerXML := '<x:patternFill xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" patternType="solid">' + '<x:fgColor rgb="' + BackgroundColor + '" /></x:patternFill>';
            PatterFill := PatterFill.PatternFill(ownerXML);
            Fill.PatternFill := PatterFill;
            Decorator.Fill := Fill;
        end;


    end;

    procedure SetColumnWidth(ColName: Text[10]; NewColWidth: Decimal)
    begin
        if not IsNull(XlWrkShtWriter) then
            XlWrkShtWriter.SetColumnWidth(ColName, NewColWidth);
    end;

    [Scope('OnPrem')]
    procedure CreateRangeName(RangeName: Text[30]; FromColumnNo: Integer; FromRowNo: Integer)
    var
        TempExcelBuf: Record "Excel Buffer 2" temporary;
        ToxlRowID: Text[10];
    begin
        SetCurrentKey("Row No.", "Column No.");
        if Find('+') then
            ToxlRowID := xlRowID;
        TempExcelBuf.Validate("Row No.", FromRowNo);
        TempExcelBuf.Validate("Column No.", FromColumnNo);

        XlWrkShtWriter.AddRange(
          RangeName,
          GetExcelReference(4) + TempExcelBuf.xlColID + GetExcelReference(4) + TempExcelBuf.xlRowID +
          ':' +
          GetExcelReference(4) + TempExcelBuf.xlColID + GetExcelReference(4) + ToxlRowID);
    end;

    procedure ReadSheet()
    begin
        ReadSheetContinous('', true);
    end;

    procedure ReadSheetContinous(SheetName: Text; CloseBookOnCompletion: Boolean)
    var
        ColumnList: List of [Integer];
        RowList: List of [Integer];
    begin
        ReadSheetContinous(SheetName, CloseBookOnCompletion, ColumnList, RowList, 0);
    end;

    procedure ReadSheetContinous(SheetName: Text; CloseBookOnCompletion: Boolean; ColumnList: List of [Integer]; RowList: List of [Integer]; MaxRowNo: Integer)
    var
        ExcelBufferDialogMgt: Codeunit "Excel Buffer Dialog Management";
        CellData: DotNet CellData;
        Enumerator: DotNet IEnumerator;
        RowCount: Integer;
        LastUpdate: DateTime;
        ReadData: Boolean;
    begin
        // Allows reading Excel files with more than one sheet without closing and reopening file
        if SheetName <> '' then
            SetActiveReaderSheet(SheetName);
        LastUpdate := CurrentDateTime;
        ExcelBufferDialogMgt.Open(Text007);
        DeleteAll();

        Enumerator := XlWrkShtReader.GetEnumerator();
        RowCount := XlWrkShtReader.RowCount;
        ReadData := Enumerator.MoveNext();
        while ReadData do begin
            CellData := Enumerator.Current;
            if CellData.HasValue() and ShouldReadCellData(CellData.ColumnNumber, CellData.RowNumber, ColumnList, RowList) then begin
                Validate("Row No.", CellData.RowNumber);
                Validate("Column No.", CellData.ColumnNumber);
                ParseCellValue(CellData.Value, CellData.Format);
                Insert();

                if not UpdateProgressDialog(ExcelBufferDialogMgt, LastUpdate, CellData.RowNumber, RowCount) then begin
                    CloseBook();
                    Error(Text035)
                end;
            end;
            ReadData := Enumerator.MoveNext();
            if MaxRowNo = CellData.RowNumber then
                ReadData := false;
        end;

        if CloseBookOnCompletion then
            CloseBook();
        ExcelBufferDialogMgt.Close();
    end;

    protected procedure ParseCellValue(Value: Text; FormatString: Text)
    var
        OutStream: OutStream;
        Decimal: Decimal;
        RoundingPrecision: Decimal;
        IsHandled: Boolean;
    begin
        // The format contains only en-US number separators, this is an OpenXML standard requirement
        // The algorithm sieves the data based on formatting as follows (the steps must run in this order)
        // 1. FormatString = '@' -> Text
        // 2. FormatString.Contains(':') -> Time
        // 3. FormatString.ContainsOneOf('y', 'm', 'd') && FormatString.DoesNotContain('Red') -> Date
        // 4. anything else -> Decimal

        IsHandled := false;
        OnBeforeParseCellValue(Rec, Value, FormatString, IsHandled);
        if IsHandled then
            exit;
        //TODO: Check if the format string is valid
        //NumberFormat := CopyStr(FormatString, 1, 30);

        Clear("Cell Value as Blob");
        if FormatString = '@' then begin
            "Cell Type" := "Cell Type"::Text;
            "Cell Value as Text" := CopyStr(Value, 1, MaxStrLen("Cell Value as Text"));

            if StrLen(Value) <= MaxStrLen("Cell Value as Text") then
                exit; // No need to store anything in the blob

            "Cell Value as Blob".CreateOutStream(OutStream, TEXTENCODING::Windows);
            OutStream.Write(Value);
            exit;
        end;

        Evaluate(Decimal, Value);

        if StrPos(FormatString, ':') <> 0 then begin
            // Excel Time is stored in OADate format
            "Cell Type" := "Cell Type"::Time;
            "Cell Value as Text" := Format(DT2Time(ConvertDateTimeDecimalToDateTime(Decimal)));
            exit;
        end;

        if ((StrPos(FormatString, 'y') <> 0) or
            (StrPos(FormatString, 'm') <> 0) or
            (StrPos(FormatString, 'd') <> 0)) and
           (StrPos(FormatString, 'Red') = 0)
        then begin
            "Cell Type" := "Cell Type"::Date;
            "Cell Value as Text" := Format(DT2Date(ConvertDateTimeDecimalToDateTime(Decimal)));
            exit;
        end;

        "Cell Type" := "Cell Type"::Number;
        RoundingPrecision := 0.000001;
        OnParseCellValueOnBeforeRoundDecimal(Rec, Decimal, RoundingPrecision);
        "Cell Value as Text" := Format(Round(Decimal, RoundingPrecision), 0, 1);
    end;

    [Scope('OnPrem')]
    procedure SelectSheetsName(FileName: Text): Text[250]
    var
        TempBlob: Codeunit "Temp Blob";
        InStr: InStream;
    begin
        if FileName = '' then
            Error(Text001);

        FileManagement.IsAllowedPath(FileName, false);
        FileManagement.BLOBImportFromServerFile(TempBlob, FileName);
        TempBlob.CreateInStream(InStr);
        exit(SelectSheetsNameStream(InStr));
    end;

    procedure SelectSheetsNameStream(FileStream: InStream): Text[250]
    var
        TempNameValueBuffer: Record "Name/Value Buffer" temporary;
        SelectedSheetName: Text[250];
    begin
        if GetSheetsNameListFromStream(FileStream, TempNameValueBuffer) then
            if TempNameValueBuffer.Count = 1 then
                SelectedSheetName := TempNameValueBuffer.Value
            else begin
                TempNameValueBuffer.FindFirst();
                if PAGE.RunModal(PAGE::"Name/Value Lookup", TempNameValueBuffer) = ACTION::LookupOK then
                    SelectedSheetName := TempNameValueBuffer.Value;
            end;

        exit(SelectedSheetName);
    end;

    procedure GetExcelReference(Which: Integer): Text[250]
    begin
        case Which of
            1:
                exit(Text013);
            // DO NOT TRANSLATE: &B is the Excel code to turn bold printing on or off for customized Header/Footer.
            2:
                exit(Text014);
            // DO NOT TRANSLATE: &D is the Excel code to print the current date in customized Header/Footer.
            3:
                exit(Text015);
            // DO NOT TRANSLATE: &P is the Excel code to print the page number in customized Header/Footer.
            4:
                exit('$');
            // DO NOT TRANSLATE: $ is the Excel code for absolute reference to cells.
            5:
                exit(Text016);
            // DO NOT TRANSLATE: A1 is the Excel reference of the first cell.
            6:
                exit(Text017);
            // DO NOT TRANSLATE: SUMIF is the name of the Excel function used to summarize values according to some conditions.
            7:
                exit(Text018);
            // DO NOT TRANSLATE: The #N/A Excel error value occurs when a value is not available to a function or formula.
            8:
                exit(Text019);
            // DO NOT TRANSLATE: GLAcc is used to define an Excel range name. You must refer to Excel rules to change this term.
            9:
                exit(Text020);
            // DO NOT TRANSLATE: Period is used to define an Excel range name. You must refer to Excel rules to change this term.
            10:
                exit(Text021);
            // DO NOT TRANSLATE: Budget is used to define an Excel worksheet name. You must refer to Excel rules to change this term.
            11:
                exit(Text022);
        // DO NOT TRANSLATE: CostAcc is used to define an Excel range name. You must refer to Excel rules to change this term.
        end;
    end;

    procedure ExportBudgetFilterToFormula(var ExcelBuf: Record "Excel Buffer 2"): Boolean
    var
        TempExcelBufFormula: Record "Excel Buffer 2" temporary;
        TempExcelBufFormula2: Record "Excel Buffer 2" temporary;
        FirstRow: Integer;
        LastRow: Integer;
        HasFormulaError: Boolean;
        ThisCellHasFormulaError: Boolean;
    begin
        FirstRow := 0;
        ExcelBuf.SetFilter(Formula, '<>%1', '');
        if ExcelBuf.FindSet() then
            repeat
                TempExcelBufFormula := ExcelBuf;
                TempExcelBufFormula.Insert();
            until ExcelBuf.Next() = 0;
        ExcelBuf.Reset();

        with TempExcelBufFormula do
            if FindFirst() then
                repeat
                    ThisCellHasFormulaError := false;
                    ExcelBuf.SetRange("Column No.", 1);
                    ExcelBuf.SetFilter("Row No.", '<>%1', "Row No.");
                    ExcelBuf.SetFilter("Cell Value as Text", Formula);
                    TempExcelBufFormula2 := TempExcelBufFormula;
                    if ExcelBuf.FindSet() then
                        repeat
                            if not Get(ExcelBuf."Row No.", "Column No.") then
                                ExcelBuf.Mark(true);
                        until ExcelBuf.Next() = 0;
                    TempExcelBufFormula := TempExcelBufFormula2;
                    ClearFormula();
                    ExcelBuf.SetRange("Cell Value as Text");
                    ExcelBuf.SetRange("Row No.");
                    if ExcelBuf.FindSet() then
                        repeat
                            if ExcelBuf.Mark() then begin
                                LastRow := ExcelBuf."Row No.";
                                if FirstRow = 0 then
                                    FirstRow := LastRow;
                            end else
                                if FirstRow <> 0 then begin
                                    if FirstRow = LastRow then
                                        ThisCellHasFormulaError := AddToFormula(xlColID + Format(FirstRow))
                                    else
                                        ThisCellHasFormulaError :=
                                          AddToFormula('SUM(' + xlColID + Format(FirstRow) + ':' + xlColID + Format(LastRow) + ')');
                                    FirstRow := 0;
                                    if ThisCellHasFormulaError then
                                        SetFormula(ExcelBuf.GetExcelReference(7));
                                end;
                        until ThisCellHasFormulaError or (ExcelBuf.Next() = 0);

                    if not ThisCellHasFormulaError and (FirstRow <> 0) then begin
                        if FirstRow = LastRow then
                            ThisCellHasFormulaError := AddToFormula(xlColID + Format(FirstRow))
                        else
                            ThisCellHasFormulaError :=
                              AddToFormula('SUM(' + xlColID + Format(FirstRow) + ':' + xlColID + Format(LastRow) + ')');
                        FirstRow := 0;
                        if ThisCellHasFormulaError then
                            SetFormula(ExcelBuf.GetExcelReference(7));
                    end;

                    ExcelBuf.Reset();
                    ExcelBuf.Get("Row No.", "Column No.");
                    ExcelBuf.SetFormula(GetFormula());
                    ExcelBuf.Modify();
                    HasFormulaError := HasFormulaError or ThisCellHasFormulaError;
                until Next() = 0;

        exit(HasFormulaError);
    end;

    procedure AddToFormula(Text: Text[30]): Boolean
    var
        Overflow: Boolean;
        LongFormula: Text[1000];
    begin
        LongFormula := GetFormula();
        if LongFormula = '' then
            LongFormula := '=';
        if LongFormula <> '=' then
            if StrLen(LongFormula) + 1 > MaxStrLen(LongFormula) then
                Overflow := true
            else
                LongFormula := LongFormula + '+';
        if StrLen(LongFormula) + StrLen(Text) > MaxStrLen(LongFormula) then
            Overflow := true
        else
            SetFormula(LongFormula + Text);
        exit(Overflow);
    end;

    procedure GetFormula(): Text[1000]
    begin
        exit(Formula + Formula2 + Formula3 + Formula4);
    end;

    procedure SetFormula(LongFormula: Text[1000])
    begin
        ClearFormula();
        if LongFormula = '' then
            exit;

        Formula := CopyStr(LongFormula, 1, MaxStrLen(Formula));
        if StrLen(LongFormula) > MaxStrLen(Formula) then
            Formula2 := CopyStr(LongFormula, MaxStrLen(Formula) + 1, MaxStrLen(Formula2));
        if StrLen(LongFormula) > MaxStrLen(Formula) + MaxStrLen(Formula2) then
            Formula3 := CopyStr(LongFormula, MaxStrLen(Formula) + MaxStrLen(Formula2) + 1, MaxStrLen(Formula3));
        if StrLen(LongFormula) > MaxStrLen(Formula) + MaxStrLen(Formula2) + MaxStrLen(Formula3) then
            Formula4 := CopyStr(LongFormula, MaxStrLen(Formula) + MaxStrLen(Formula2) + MaxStrLen(Formula3) + 1, MaxStrLen(Formula4));
    end;

    procedure ClearFormula()
    begin
        Formula := '';
        Formula2 := '';
        Formula3 := '';
        Formula4 := '';
    end;

    procedure NewRow()
    begin
        SetCurrent(CurrentRow + 1, 0);
    end;

    procedure AddColumn(Value: Variant; IsFormula: Boolean; CommentText: Text; IsBold: Boolean; IsItalics: Boolean; IsUnderline: Boolean; NumFormat: Text[30]; CellType: Option)
    begin
        AddColumnToBuffer(Rec, Value, IsFormula, CommentText, IsBold, IsItalics, IsUnderline, NumFormat, CellType);
    end;

    procedure AddInfoColumn(Value: Variant; IsFormula: Boolean; IsBold: Boolean; IsItalics: Boolean; IsUnderline: Boolean; NumFormat: Text[30]; CellType: Option)
    begin
        AddColumnToBuffer(TempInfoExcelBuf, Value, IsFormula, '', IsBold, IsItalics, IsUnderline, NumFormat, CellType);
    end;

    local procedure AddColumnToBuffer(var ExcelBuffer: Record "Excel Buffer 2"; Value: Variant; IsFormula: Boolean; CommentText: Text; IsBold: Boolean; IsItalics: Boolean; IsUnderline: Boolean; NumFormat: Text[250]; CellType: Option)
    begin
        if CurrentRow < 1 then
            NewRow();

        CurrentCol := CurrentCol + 1;
        ExcelBuffer.Init();
        ExcelBuffer.Validate("Row No.", CurrentRow);
        ExcelBuffer.Validate("Column No.", CurrentCol);
        if IsFormula then
            ExcelBuffer.SetFormula(Format(Value))
        else
            ExcelBuffer."Cell Value as Text" := Format(Value);
        ExcelBuffer.Comment := CopyStr(CommentText, 1, MaxStrLen(ExcelBuffer.Comment));
        ExcelBuffer.Bold := IsBold;
        ExcelBuffer.Italic := IsItalics;
        ExcelBuffer.Underline := IsUnderline;
        ExcelBuffer.NumberFormat := NumFormat;
        ExcelBuffer."Cell Type" := CellType;
        ExcelBuffer.Insert();
        OnAfterAddColumnToBuffer(ExcelBuffer, Value, IsFormula, CommentText, IsBold, IsItalics, IsUnderline, NumFormat, CellType);
    end;

    procedure EnterCell(var ExcelBuffer: Record "Excel Buffer 2"; RowNo: Integer; ColumnNo: Integer; Value: Variant; IsBold: Boolean; IsItalics: Boolean; IsUnderline: Boolean)
    begin
        ExcelBuffer.Init();
        ExcelBuffer.Validate("Row No.", RowNo);
        ExcelBuffer.Validate("Column No.", ColumnNo);

        case true of
            Value.IsDecimal or Value.IsInteger:
                ExcelBuffer.Validate("Cell Type", ExcelBuffer."Cell Type"::Number);
            Value.IsDate:
                ExcelBuffer.Validate("Cell Type", ExcelBuffer."Cell Type"::Date);
            else
                ExcelBuffer.Validate("Cell Type", ExcelBuffer."Cell Type"::Text);
        end;

        ExcelBuffer."Cell Value as Text" := CopyStr(Format(Value), 1, MaxStrLen(ExcelBuffer."Cell Value as Text"));
        ExcelBuffer.Bold := IsBold;
        ExcelBuffer.Italic := IsItalics;
        ExcelBuffer.Underline := IsUnderline;
        ExcelBuffer.Insert(true);
    end;

    procedure StartRange()
    var
        DummyExcelBuf: Record "Excel Buffer 2";
    begin
        DummyExcelBuf.Validate("Row No.", CurrentRow);
        DummyExcelBuf.Validate("Column No.", CurrentCol);

        RangeStartXlRow := DummyExcelBuf.xlRowID;
        RangeStartXlCol := DummyExcelBuf.xlColID;
    end;

    procedure EndRange()
    var
        DummyExcelBuf: Record "Excel Buffer 2";
    begin
        DummyExcelBuf.Validate("Row No.", CurrentRow);
        DummyExcelBuf.Validate("Column No.", CurrentCol);

        RangeEndXlRow := DummyExcelBuf.xlRowID;
        RangeEndXlCol := DummyExcelBuf.xlColID;
    end;

    procedure CreateRange(RangeName: Text[250])
    begin
        XlWrkShtWriter.AddRange(
          RangeName,
          GetExcelReference(4) + RangeStartXlCol + GetExcelReference(4) + RangeStartXlRow +
          ':' +
          GetExcelReference(4) + RangeEndXlCol + GetExcelReference(4) + RangeEndXlRow);
    end;

    procedure ClearNewRow()
    begin
        SetCurrent(0, 0);
    end;

    procedure SetUseInfoSheet()
    begin
        UseInfoSheet := true;
    end;

    procedure UTgetGlobalValue(globalVariable: Text[30]; var value: Variant)
    begin
        case globalVariable of
            'CurrentRow':
                value := CurrentRow;
            'CurrentCol':
                value := CurrentCol;
            'RangeStartXlRow':
                value := RangeStartXlRow;
            'RangeStartXlCol':
                value := RangeStartXlCol;
            'RangeEndXlRow':
                value := RangeEndXlRow;
            'RangeEndXlCol':
                value := RangeEndXlCol;
            'XlWrkSht':
                value := XlWrkShtWriter;
            'ExcelFile':
                value := FileNameServer;
            else
                Error(Text038, globalVariable);
        end;
    end;

    procedure SetCurrent(NewCurrentRow: Integer; NewCurrentCol: Integer)
    begin
        CurrentRow := NewCurrentRow;
        CurrentCol := NewCurrentCol;
    end;

    local procedure AddHyperLink(Link: Text[250])
    begin

        XlWrkShtWriter.AddHyperlink("Row No.", xlColID, Link);
    end;

    procedure CreateValidationRule(Range: Code[20])
    begin
        XlWrkShtWriter.AddRangeDataValidation(
          Range,
          GetExcelReference(4) + RangeStartXlCol + GetExcelReference(4) + RangeStartXlRow +
          ':' +
          GetExcelReference(4) + RangeEndXlCol + GetExcelReference(4) + RangeEndXlRow);
    end;

    procedure QuitExcel()
    begin
        CloseBook();
    end;

    procedure OpenExcel()
    begin
        if OpenUsingDocumentService('') then
            exit;

        FileManagement.DownloadHandler(FileNameServer, '', '', Text034, GetFriendlyFilename());
    end;

    [Scope('OnPrem')]
    procedure DownloadAndOpenExcel()
    begin
        OpenExcelWithName(GetFriendlyFilename());
    end;

    [Scope('OnPrem')]
    procedure OpenExcelWithName(FileName: Text)
    begin
        if FileName = '' then
            Error(Text001);

        if OpenUsingDocumentService(FileName) then
            exit;

        FileManagement.DownloadHandler(FileNameServer, '', '', Text034, FileName);
    end;

    local procedure OpenUsingDocumentService(FileName: Text) Result: Boolean
    var
        DocumentServiceMgt: Codeunit "Document Service Management";
        FileMgt: Codeunit "File Management";
        PathHelper: DotNet Path;
        DialogWindow: Dialog;
        DocumentUrl: Text;
        IsHandled: Boolean;
    begin
        IsHandled := false;
        OnBeforeOpenUsingDocumentService(FileNameServer, Filename, Result, IsHandled);
        if IsHandled then
            exit(Result);

        if not Exists(FileNameServer) then
            Error(Text003, FileNameServer);

        // if document service is configured we save the generated document to SharePoint and open it from there.
        if DocumentServiceMgt.IsConfigured() then begin
            if FileName = '' then
                FileName := 'Book.' + PathHelper.ChangeExtension(PathHelper.GetRandomFileName(), 'xlsx')
            else begin
                // if file is not applicable for the service it can not be opened using the document service.
                if not DocumentServiceMgt.IsServiceUri(FileName) then
                    exit(false);

                FileName := FileMgt.GetFileName(FileName);
            end;

            DialogWindow.Open(StrSubstNo(SavingDocumentMsg, FileName));
            DocumentUrl := DocumentServiceMgt.SaveFile(FileNameServer, FileName, Enum::"Doc. Sharing Conflict Behavior"::Replace);
            DocumentServiceMgt.OpenDocument(DocumentUrl);
            DialogWindow.Close();
            exit(true);
        end;

        exit(false);
    end;

    [Scope('OnPrem')]
    procedure CreateBookAndOpenExcel(FileName: Text; SheetName: Text[250]; ReportHeader: Text; CompanyName2: Text; UserID2: Text; Orientacion: Enum Orientacion)
    begin
        CreateBook(FileName, SheetName);
        WriteSheet(ReportHeader, CompanyName2, UserID2, Orientacion);
        CloseBook();
        OpenExcel();
    end;

    local procedure UpdateProgressDialog(var ExcelBufferDialogManagement: Codeunit "Excel Buffer Dialog Management"; var LastUpdate: DateTime; CurrentCount: Integer; TotalCount: Integer): Boolean
    var
        CurrentTime: DateTime;
    begin
        // Refresh at 100%, and every second in between 0% to 100%
        // Duration is measured in miliseconds -> 1 sec = 1000 ms
        CurrentTime := CurrentDateTime;
        if (CurrentCount = TotalCount) or (CurrentTime - LastUpdate >= 1000) then begin
            LastUpdate := CurrentTime;
            if not ExcelBufferDialogManagement.SetProgress(Round(CurrentCount / TotalCount * 10000, 1)) then
                exit(false);
        end;

        exit(true)
    end;

    local procedure GetFriendlyFilename(): Text
    begin
        if FriendlyName = '' then
            exit('Book1' + ExcelFileExtensionTok);

        exit(FileManagement.StripNotsupportChrInFileName(FriendlyName) + ExcelFileExtensionTok);
    end;

    procedure SetFriendlyFilename(Name: Text)
    begin
        FriendlyName := Name;
    end;

    procedure ConvertDateTimeDecimalToDateTime(DateTimeAsOADate: Decimal): DateTime
    var
        DotNetDateTime: DotNet DateTime;
        DateTimeResult: DateTime;
        DotNetDateTimeKind: DotNet DateTimeKind;
    begin
        DotNetDateTime := DotNetDateTime.FromOADate(DateTimeAsOADate);
        if ReadDateTimeInUtcDate then
            Evaluate(DateTimeResult, DotNetDateTime.ToString())
        else
            DateTimeResult := DotNetDateTime.DateTime(DotNetDateTime.Ticks, DotNetDateTimeKind.Local);
        exit(DateTimeResult);
    end;

    procedure SaveToStream(var ResultStream: OutStream; EraseFileAfterCompletion: Boolean)
    var
        TempBlob: Codeunit "Temp Blob";
        BlobStream: InStream;
    begin
        FileManagement.BLOBImportFromServerFile(TempBlob, FileNameServer);
        TempBlob.CreateInStream(BlobStream);
        CopyStream(ResultStream, BlobStream);
        if EraseFileAfterCompletion then
            FILE.Erase(FileNameServer);
    end;

    procedure GetSheetsNameListFromStream(FileStream: InStream; var TempNameValueBufferOut: Record "Name/Value Buffer" temporary) SheetsFound: Boolean
    var
        SheetNames: DotNet StringArray;
        SheetName: Text[250];
        i: Integer;
    begin
        XlWrkBkReader := XlWrkBkReader.Open(FileStream);
        TempNameValueBufferOut.Reset();
        TempNameValueBufferOut.DeleteAll();
        //SheetNames := SheetNames.ArrayList(XlWrkBkReader.SheetNames());
        SheetNames := (XlWrkBkReader.SheetNames());
        if IsNull(SheetNames) then
            exit(false);

        SheetsFound := SheetNames.Length > 0;

        if not SheetsFound then
            exit(false);
        foreach SheetName in SheetNames do begin
            if SheetName <> '' then begin
                i += 1;
                TempNameValueBufferOut.Init();
                TempNameValueBufferOut.ID := i;
                TempNameValueBufferOut.Name := Format(i + 1);
                TempNameValueBufferOut.Value := SheetName;
                TempNameValueBufferOut.Insert();
            end;
        end;

        // for i := 0 to SheetNames.Length - 1 do begin
        //     SheetName := SheetNames. Item(i);
        //     if SheetName <> '' then begin
        //         TempNameValueBufferOut.Init();
        //         TempNameValueBufferOut.ID := i;
        //         TempNameValueBufferOut.Name := Format(i + 1);
        //         TempNameValueBufferOut.Value := SheetName;
        //         TempNameValueBufferOut.Insert();
        //     end;
        // end;

        CloseBook();
    end;

    local procedure ShouldReadCellData(ColumnNo: Integer; RowNo: Integer; ColumnList: List of [Integer]; RowList: List of [Integer]): Boolean
    begin
        if (ColumnList.Count = 0) and (RowList.Count = 0) then
            exit(true);

        if ColumnList.Count = 0 then
            exit(RowList.Contains(RowNo));

        if RowList.Count = 0 then
            exit(ColumnList.Contains(ColumnNo));

        exit(ColumnList.Contains(ColumnNo) and RowList.Contains(RowNo));
    end;

    [IntegrationEvent(false, false)]
    local procedure OnAfterAddColumnToBuffer(var ExcelBuffer: Record "Excel Buffer 2"; Value: Variant; IsFormula: Boolean; CommentText: Text; IsBold: Boolean; IsItalics: Boolean; IsUnderline: Boolean; NumFormat: Text[250]; CellType: Option)
    begin
    end;

    [IntegrationEvent(false, false)]
    local procedure OnBeforeOpenUsingDocumentService(FileNameServer: Text; FileName: Text; var Result: Boolean; var IsHandled: Boolean)
    begin
    end;

    [IntegrationEvent(false, false)]
    local procedure OnBeforeParseCellValue(var ExcelBuffer: Record "Excel Buffer 2"; var Value: Text; var FormatString: Text; var IsHandled: Boolean)
    begin
    end;

    [IntegrationEvent(false, false)]
    local procedure OnBeforeWriteCellFormula(var Rec: Record "Excel Buffer 2"; var ExcelBuffer: Record "Excel Buffer 2"; var IsHandled: Boolean)
    begin
    end;

    [IntegrationEvent(false, false)]
    local procedure OnParseCellValueOnBeforeRoundDecimal(var ExcelBuffer: Record "Excel Buffer 2"; DecimalValue: Decimal; var RoundingPrecision: Decimal)
    begin
    end;

    [IntegrationEvent(false, false)]
    local procedure OnWriteSheetOnBeforeUseXmlTextWriter(var ExcelBuffer: Record "Excel Buffer 2"; var IsHandled: Boolean)
    begin
    end;

    [IntegrationEvent(false, false)]
    local procedure OnWriteSheetOnBeforeAddAndInitializeCommentsPart(var ExcelBuffer: Record "Excel Buffer 2"; var IsHandled: Boolean)
    begin
    end;

    [IntegrationEvent(false, false)]
    local procedure OnWriteCellValueOnBeforeSetCellValue(var ExcelBuffer: Record "Excel Buffer 2"; var CellTextValue: Text)
    begin
    end;
}

