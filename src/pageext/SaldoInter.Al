pageextension 93004 SaldoInterExt extends "Saldo Interempresas Det"
{

    procedure ExportExcel(var Filtros: Record "Filtros Informes";
    IdInforme: Integer;
    Var Destinatario: Record "Destinatarios Informes"; var ExcelStream: OutStream)
    var
        TempExcelBuffer: Record "Excel Buffer 2" temporary;
        ContabilidadLblEP: Label 'Saldo InterEmpresas';
        ExcelFileNameEPR: Text;
        RecordLink: Record "Record Link";
        RecordLinkMgt: Codeunit "Record Link Management";
        Empresas: Record "Empresas Informes";
        BgText: Text;
        Row: Integer;
        Col: Integer;
        Informes: Record "Informes";
        DesdeFecha: Date;
        HastaFecha: Date;
        DF: DateFormula;
        TypeHelper: Codeunit "Type Helper";
        Matrix: Codeunit "Matrix Management";
        Rf: Enum "Analysis Rounding Factor";
        Todo: Record "To-do";
        Contact: Record "Contact";
        InExcelStream: Instream;
        Continuar: Boolean;
        FechaTarea: Date;
        Campos: Record "Columnas Informes";
        RecRef: RecordRef;
        ColumnValues: array[100] of Decimal;
        ColumnCaptions: array[100] of Text[250];
        Valor: Variant;
        FieldRef: FieldRef;
        Id: RecordId;
        FieldT: FieldType;
        Fecha: Date;
        Campo: Integer;
        NoContrato: Code[20];
        Desde: Date;
        Hasta: Date;
        Periodos: Record "Periodos Informes";
        Control: Codeunit ControlInformes;
        HojasSeparadas: Boolean;
        Formatos: Record "Formato Columnas";
        TempBlob: Codeunit "Temp Blob";
        Base64Convert: Codeunit "Base64 Convert";
        PlantillaBase64: Text;
    begin

        TempExcelBuffer.Reset();
        TempExcelBuffer.DeleteAll();
        FechaTarea := CalcDate('1S', WorkDate());
        RecRef.Open(Database::"Saldo InterEmpresa", true);
        RecReftemp.Open(Database::"Saldo InterEmpresa", true);

        Informes.Get(IdInforme);
        Campos.SetRange(Id, Informes.Id);
        ExcelFileNameEPR := ConvertStr(Informes.Descripcion, ' ', '_');
        Empresas.SetRange(Id, Informes.Id);
        Empresas.SetRange(Incluir, true);
        If not Empresas.FindSet() then
            CrearCabecera(Informes.Id, TempExcelBuffer, Row, DesdeFecha, HastaFecha, FieldRef, RecReftemp, '')
        else begin
            If Filtros.FindSet() then
                repeat
                    Row += 1;
                    If Filtros.Desde <> DF then DesdeFecha := CalcDate(Filtros.Desde, WorkDate()) else DesdeFecha := 0D;
                    If Filtros.Hasta <> DF then HastaFecha := CalcDate(Filtros.Hasta, WorkDate()) else HastaFecha := Calcdate('99A', WorkDate());

                    FieldRef := RecReftemp.Field(Filtros.Campo);
                    if (filtros.Desde <> DF) or (Filtros.Hasta <> DF) then begin
                        FieldRef.SetRange(DesdeFecha, HastaFecha);
                        if FieldRef.Name = 'Posting Date' then begin
                            Desde := DesdeFecha;
                            Hasta := HastaFecha;
                        end;

                        //EnterCell(TempExcelBuffer, Row, 1, FieldRef.Caption, true, false, '', TempExcelBuffer."Cell Type"::Text);
                        //if DesdeFecha <> 0D then
                        //  EnterCell(TempExcelBuffer, Row, 2, CopyStr(TypeHelper.FormatDateWithCurrentCulture(DesdeFecha), 1, 250), Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Text);
                        //EnterCell(TempExcelBuffer, Row, 3, CopyStr(TypeHelper.FormatDateWithCurrentCulture(HastaFecha), 1, 250), Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Text);
                    end else begin
                        FieldRef.SetFilter(Filtros.Valor);
                        //EnterCell(TempExcelBuffer, Row, 1, FieldRef.Caption, true, false, '', TempExcelBuffer."Cell Type"::Text);
                        //EnterCell(TempExcelBuffer, Row, 2, Filtros.Valor, Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Text);
                    end;
                until Filtros.Next() = 0;
            Row += 1;
            if Destinatario."Campo Destinatario" <> 0 then begin
                FieldRef := RecReftemp.Field(Destinatario."Campo Destinatario");
                FieldRef.SetFilter(Destinatario.Valor);
            end;
        end;
        if Empresas.FindSet() then
            if not Empresas."Hojas Separadas" then
                CrearCabecera(Informes.Id, TempExcelBuffer, Row, DesdeFecha, HastaFecha, FieldRef, RecReftemp, Empresas.Empresa)
            else begin
                If Filtros.FindSet() then
                    repeat
                        Row += 1;
                        If Filtros.Desde <> DF then DesdeFecha := CalcDate(Filtros.Desde, WorkDate()) else DesdeFecha := 0D;
                        If Filtros.Hasta <> DF then HastaFecha := CalcDate(Filtros.Hasta, WorkDate()) else HastaFecha := Calcdate('99A', WorkDate());

                        FieldRef := RecReftemp.Field(Filtros.Campo);
                        if (filtros.Desde <> DF) or (Filtros.Hasta <> DF) then begin
                            FieldRef.SetRange(DesdeFecha, HastaFecha);
                            if FieldRef.Name = 'Posting Date' then begin
                                Desde := DesdeFecha;
                                Hasta := HastaFecha;
                            end;

                            //EnterCell(TempExcelBuffer, Row, 1, FieldRef.Caption, true, false, '', TempExcelBuffer."Cell Type"::Text);
                            // if DesdeFecha <> 0D then
                            //     EnterCell(TempExcelBuffer, Row, 2, CopyStr(TypeHelper.FormatDateWithCurrentCulture(DesdeFecha), 1, 250), Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Text);
                            // EnterCell(TempExcelBuffer, Row, 3, CopyStr(TypeHelper.FormatDateWithCurrentCulture(HastaFecha), 1, 250), Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Text);
                        end else begin
                            FieldRef.SetFilter(Filtros.Valor);
                            // EnterCell(TempExcelBuffer, Row, 1, FieldRef.Caption, true, false, '', TempExcelBuffer."Cell Type"::Text);
                            //  EnterCell(TempExcelBuffer, Row, 2, Filtros.Valor, Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Text);
                        end;
                    until Filtros.Next() = 0;
                if FieldRef.Name = 'Schedule Name' Then Schedulename := Filtros.Valor;
                Row += 1;
                if Destinatario."Campo Destinatario" <> 0 then begin
                    FieldRef := RecReftemp.Field(Destinatario."Campo Destinatario");
                    FieldRef.SetFilter(Destinatario.Valor);
                end;
            end;
        Row += 1;

        HojasSeparadas := false;
        Empresas.SetRange(Id, Informes.Id);
        Empresas.SetRange(Incluir, true);
        If not Empresas.FindSet() then begin
            Empresas.Init();
            Empresas.Empresa := CompanyName;
        end else begin
            HojasSeparadas := Empresas."Hojas Separadas";
            if HojasSeparadas then begin

                if (Informes."Plantilla Excel".HasValue) Or (Informes."Url Plantilla" <> '') then begin

                    TempExcelBuffer.UpdateBookStream(InExcelStream, ConvertStr(Empresas.HojaExcel, ' ', '_'), true);

                end else
                    TempExcelBuffer.CreateNewBook(ConvertStr(Empresas.Empresa, ' ', '_'));


            end;
        end;
        repeat


            if HojasSeparadas then begin
                TempExcelBuffer.SelectOrAddSheet(ConvertStr(Empresas.HojaExcel, ' ', '_'));
                Row := 0;
                CrearCabecera(Informes.Id, TempExcelBuffer, Row, DesdeFecha, HastaFecha, FieldRef, RecReftemp, Empresas.Empresa);
                Empresas.TestField("HojaExcel");
            end;
            "Periodos".SetRange(Id, Informes.Id);
            If not "Periodos".FindFirst() then begin
                "Periodos".Init();
                "Periodos".Id := Informes.Id;
                Periodos.Periodo := 'Ninguno';
                if Desde = 0D then Desde := DesdeFecha;
                if Hasta = 0D then Hasta := HastaFecha;
            end;
            repeat
                if ("Periodos".Periodo <> 'Ninguno') then begin
                    iF Periodos.Semana then begin
                        // Semana := Round((Workdate() - CalcDate('PA+1D-1A', WorkDate())) / 7, 1, '<');

                        HastaFecha := CalcDate(Format(Periodos.Hasta) + '+' + Format(Control.Semana(WorkDate())) + 'S', WorkDate());
                        If Periodos.Desde <> DF then DesdeFecha := CalcDate(Format(Periodos.Desde) + '+' + Format(Control.Semana(WorkDate())) + 'S', WorkDate());
                        FieldRef := RecReftemp.Field(Periodos.Campo);
                        if (Periodos.Desde <> DF) or (Periodos.Hasta <> DF) then begin
                            FieldRef.SetRange(DesdeFecha, HastaFecha);
                        end;
                    end else begin
                        If "Periodos".Desde <> DF then Desde := CalcDate("Periodos".Desde, WorkDate()) else DesdeFecha := 0D;
                        If "Periodos".Hasta <> DF then Hasta := CalcDate("Periodos".Hasta, WorkDate()) else HastaFecha := Calcdate('99A', WorkDate());
                    end;
                end;
                if Desde = 0D then Desde := DesdeFecha;
                if Hasta = 0D then Hasta := HastaFecha;
                If Desde = 0D Then Error('No se ha definido la fecha de inicio');
                Rec.SetRange("Filtro Fecha", Desde, Hasta);
                Campos.SetRange(Id, Informes.Id);
                Campos.SetRange(Include, true);
                RecTemp.DeleteAll();
                If Rec.FindFirst() then
                    repeat
                        RecRef.GetTable(Rec);
                        for Campo := 1 to RecRef.FieldCount do begin
                            If RecRef.FieldIndex(campo).Active then begin
                                FieldRef := RecRef.FieldIndex(Campo);
                                RecRefTemp.Fieldindex(Campo).Value := FieldRef.Value;

                            end;
                            //Campo += 1
                        end;
                        RecReftemp.Insert();
                    until Rec.Next() = 0;

                if RecReftemp.FindSet() then
                    repeat
                        Rec.SetPosition(RecReftemp.GetPosition());
                        Rec.Get(Rec."Source Type", Rec."Source No.");
                        Clear(ColumnValues);
                        RecTemp := Rec;
                        If RecTEMP.Insert Then;
                        DetalleColumnas(ColumnValues, ColumnCaptions, Empresas.Empresa, RecTemp);
                        RecReftemp.field(Rec.Fieldno("Saldo en otra empresa")).Value := Rectemp."Saldo en otra empresa";
                        RecReftemp.field(Rec.Fieldno("Saldo en otra empresa Albara")).Value := Rectemp."Saldo en otra empresa Albara";
                        RecReftemp.field(Rec.Fieldno("Saldo en otra Empresa Fact/Abo")).Value := Rectemp."Saldo en otra Empresa Fact/Abo";
                        if Informes."Crear Tarea" then
                            CreateTask(Destinatario.Valor
                               , CompanyName
                                , FechaTarea, Informes."Descripcion Tarea");

                        //
                        Row += 1;
                        if Campos.FindSet() then
                            repeat

                                rf := "Analysis Rounding Factor"::None;
                                If (Campos.Campo <> 0) And (Campos.Funcion = Campos.Funcion::" ") then begin
                                    FieldRef := RecRef.Field(Campos.Campo);
                                    FieldT := FieldRef.Type;
                                    Valor := DevuelveCampo(Campos.Campo);
                                end else begin
                                    FieldT := FieldType::Text;
                                    //Importe,Vendedor,GetTotImp,ImporteIva,GetImpBorFac,GetImpBorAbo,GetImpFac,GetImpAbo,GetTotCont
                                    case Campos.Funcion of
                                        Funciones::Cliente_Proveedor:
                                            begin
                                                FieldT := FieldType::Text;
                                                Valor := Cliente_Proveedor(RecrefTemp.Field(Campos.Campo).Value, Empresas.Empresa, RecrefTemp.Field(Rec.FieldNo("Source Type")).Value);
                                            end;
                                        funciones::Cadena:
                                            begin
                                                FieldT := FieldType::Text;
                                                Valor := Cadena(RecrefTemp.Field(Campos.Campo).Value, Empresas.Empresa);
                                            end;
                                        // Funciones::Vendedor:
                                        //     begin
                                        //         FieldT := FieldType::Text;
                                        //         Valor := Vendedor(RecrefTemp.Field(Rec.FieldNo("Salesperson Code")).Value, RecrefTemp.Field(Rec.FieldNo("Periodo de Pago")).Value);
                                        //     end;
                                        Funciones::"Año":
                                            begin
                                                FieldT := FieldType::Integer;
                                                Valor := Date2DMY(RecrefTemp.Field(Campos.Campo).Value, 3);

                                            end;
                                        Funciones::"Mes":
                                            begin
                                                FieldT := FieldType::Integer;
                                                Valor := Date2DMY(RecrefTemp.Field(Campos.Campo).Value, 2);
                                            end;
                                        Funciones::"Semana":
                                            begin
                                                FieldT := FieldType::Integer;
                                                Valor := Control.Semana(RecrefTemp.Field(Campos.Campo).Value);
                                            end;
                                        Funciones::Columna:
                                            begin
                                                FieldT := FieldType::Decimal;
                                                //Valor := Format(ColumnValues[Campos.Campo], 0, '<Precision,2:3><Standard Format,0>');
                                                Valor := ColumnValues[Campos.Campo];
                                            end;
                                        else
                                            Valor := '';
                                    end;
                                end;
                                If Not Formatos.Get(campos.Id, campos.Id_campo, false) then begin
                                    Formatos.Init();
                                    If FieldT = FieldType::Decimal then
                                        Formatos."Formato Columna" := '_-* #,##0.00_-;-* #,##0.00_-;_-* "-"_-;_-@_-';

                                end;
                                Case FieldT of
                                    FieldT::Date:
                                        begin
                                            if Valor.IsDate then Fecha := Valor else Fecha := 0D;
                                            iF Fecha <> 0D then
                                                EnterCell(TempExcelBuffer, Row, Campos.Orden, CopyStr(TypeHelper.FormatDateWithCurrentCulture(Fecha), 1, 250), Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Date, Formatos."Fuente", Formatos."Tamaño", Formatos."Color Fuente", Formatos."Color Fondo")
                                            else
                                                EnterCell(TempExcelBuffer, Row, Campos.Orden, '', Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Text, Formatos."Fuente", Formatos."Tamaño", Formatos."Color Fuente", Formatos."Color Fondo");
                                        END;
                                    FieldT::Time:
                                        EnterCell(TempExcelBuffer, Row, Campos.Orden, Format(Valor), Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Time, Formatos."Fuente", Formatos."Tamaño", Formatos."Color Fuente", Formatos."Color Fondo");
                                    FieldT::Integer:
                                        EnterCell(TempExcelBuffer, Row, Campos.Orden, Format(Valor), Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Number, Formatos."Fuente", Formatos."Tamaño", Formatos."Color Fuente", Formatos."Color Fondo");
                                    FieldT::Decimal:
                                        EnterCell(TempExcelBuffer, Row, Campos.Orden, Format(Valor), Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Number, Formatos."Fuente", Formatos."Tamaño", Formatos."Color Fuente", Formatos."Color Fondo");
                                    FieldT::Option:
                                        EnterCell(TempExcelBuffer, Row, Campos.Orden, Valor, Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Text, Formatos."Fuente", Formatos."Tamaño", Formatos."Color Fuente", Formatos."Color Fondo");
                                    FieldT::Code:
                                        EnterCell(TempExcelBuffer, Row, Campos.Orden, Valor, Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Text, Formatos."Fuente", Formatos."Tamaño", Formatos."Color Fuente", Formatos."Color Fondo");
                                    FieldT::Text:
                                        EnterCell(TempExcelBuffer, Row, Campos.Orden, Valor, Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Text, Formatos."Fuente", Formatos."Tamaño", Formatos."Color Fuente", Formatos."Color Fondo");
                                    FieldT::Boolean:
                                        EnterCell(TempExcelBuffer, Row, Campos.Orden, Format(Valor), Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Text, Formatos."Fuente", Formatos."Tamaño", Formatos."Color Fuente", Formatos."Color Fondo");
                                    FieldT::RecordId:
                                        EnterCell(TempExcelBuffer, Row, Campos.Orden, Valor, Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Text, Formatos."Fuente", Formatos."Tamaño", Formatos."Color Fuente", Formatos."Color Fondo");
                                    FieldT::Blob:
                                        EnterCell(TempExcelBuffer, Row, Campos.Orden, Valor, Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Text, Formatos."Fuente", Formatos."Tamaño", Formatos."Color Fuente", Formatos."Color Fondo");
                                    FieldT::Guid:
                                        EnterCell(TempExcelBuffer, Row, Campos.Orden, Valor, Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Text, Formatos."Fuente", Formatos."Tamaño", Formatos."Color Fuente", Formatos."Color Fondo");

                                End;


                            until Campos.Next() = 0;


                    //end;
                    until RecrefTemp.Next() = 0;

                RecReftemp.DeleteAll();
            until "Periodos".Next() = 0;
            if HojasSeparadas then begin

                TempExcelBuffer.WriteSheet(ConvertStr(Empresas.HojaExcel, ' ', '_'), Empresas.Empresa, UserId, Informes."Orientación");
                TempExcelBuffer.DeleteAll();
            end;
        until Empresas.Next() = 0;
        Informes.CalcFields("Plantilla Excel");
        if (Informes."Plantilla Excel".HasValue) Or (Informes."Url Plantilla" <> '') then begin
            if Informes."Plantilla Excel".HasValue then
                Informes."Plantilla Excel".CreateInStream(InExcelStream);
            Control.UrlPlantilla(gUrlPlantilla, Informes, PlantillaBase64, false);
            if Not Informes."Formato Json" then
                TempExcelBuffer.UpdateBookStream(InExcelStream, ConvertStr(Informes.Descripcion, ' ', '_'), true);

        end else begin
            if Informes."Formato Json" then
                PlantillaBase64 := ''
            else
                TempExcelBuffer.CreateNewBook(ConvertStr(Informes.Descripcion, ' ', '_'));
        end;
        if Not Informes."Formato Json" then begin
            TempExcelBuffer.WriteSheet(ConvertStr(Informes.Descripcion, ' ', '_'), CompanyName, UserId, Informes."Orientación");
            TempExcelBuffer.CloseBook();
            TempExcelBuffer.SetFriendlyFilename(StrSubstNo(ExcelFileNameEPR, ConvertStr(Informes.Descripcion, ' ', '_'), CurrentDateTime, UserId));
            TempExcelBuffer.SaveToStream(ExcelStream, true);
        end else begin
            TempExcelBuffer.ModifyAll("Sheet Name", ConvertStr(Informes.Descripcion, ' ', '_'));
            PlantillaBase64 := Control.JsonExcel(TempExcelBuffer, PlantillaBase64,
            gUrlPlantilla);
            Base64Convert.FromBase64(PlantillaBase64, ExcelStream);
        end;
    end;

    Procedure DevuelveCampo(Campo: Integer) Valor: Variant
    var
        MyFieldRef: FieldRef;
    begin


        MyFieldRef := RecRefTemp.Field(Campo);
        If MyFieldRef.Type = FieldType::Option Then begin
            Exit(MyFieldRef.GetEnumValueNameFromOrdinalValue(MyFieldRef.Value));
        end;

        Exit(MyFieldRef.Value);
    end;

    local procedure EnterCell(
         Var TempExcelBuf: Record "Excel Buffer 2" temporary;
         RowNo: Integer;
         ColumnNo: Integer;
         CellValue: Text[250];
         Bold: Boolean;
         Italic: Boolean;
         UnderLine: Boolean;
         DobleUnderLine: Boolean;
         NumberFormat: Text;
         CellType: Option; Fuente: Text[30]; Tamaño: Integer; Color: Text; ColorFondo: Text)
    begin
        TempExcelBuf.Init();
        TempExcelBuf.Validate("Row No.", RowNo);
        TempExcelBuf.Validate("Column No.", ColumnNo);
        TempExcelBuf."Cell Value as Text" := CellValue;
        TempExcelBuf.Formula := '';
        TempExcelBuf.Bold := Bold;
        TempExcelBuf.Italic := Italic;
        TempExcelBuf.Underline := UnderLine;
        TempExcelBuf."Double Underline" := DobleUnderLine;
        TempExcelBuf.NumberFormat := NumberFormat;
        TempExcelBuf."Cell Type" := CellType;
        TempExcelBuf."Font Name" := Fuente;
        TempExcelBuf."Font Size" := Tamaño;
        TempExcelBuf."Font Color" := Color;
        TempExcelBuf."Background Color" := ColorFondo;
        TempExcelBuf.Insert();
    end;


    local procedure Vendedor(SalespersonCode: Code[10]; Empresa: Text): Variant
    Var
        Vendedor: Record "Salesperson/Purchaser";
    begin
        Vendedor.ChangeCompany(Empresa);
        if Vendedor.Get(SalespersonCode) then
            exit(Vendedor.Name)
        else
            exit('');
    end;

    procedure CreateTask(Vendedor: Code[20]; Empresa: Text; Fecha: Date; Descripcion: Text[250])
    var
        "To-do": Record "To-do";
        Cont: Record Contact;
        TempAttendee: Record Attendee temporary;
        RMSetup: Record "Marketing Setup";
        Ser: Record "No. Series Line";
        TempEndDateTime: DateTime;
        Contrato: Record "Sales Header";
    begin
        //Error('No se puede crear la tarea');
        "To-do".ChangeCompany(Empresa);
        "To-do".Init();
        "To-do"."Contact No." := '';

        Cont.ChangeCompany(Empresa);
        if Cont.Get("To-do"."Contact No.") then
            "To-do"."Contact Company No." := Cont."Company No."
        else
            Clear("To-do"."Contact Company No.");
        if ("To-do"."No." <> '') and
            ("To-do"."No." = "To-do"."Organizer To-do No.") and
            ("To-do".Type <> "To-do".Type::Meeting)
        then begin
            TempAttendee.ChangeCompany(Empresa);
            TempAttendee.CreateAttendee(
            TempAttendee,
                    "To-do"."No.", 20000, TempAttendee."Attendance Type"::Required,
                    TempAttendee."Attendee Type"::Contact,
                    "To-do"."Contact No.", false);
            "To-do".CreateSubTask(TempAttendee, "To-do");
        end;


        "To-do".SetRange("Contact No.", '');
        if Vendedor <> '' then begin
            "To-do"."Salesperson Code" := Vendedor;

        end;

        "To-do".Description := Descripcion;
        "To-do"."Descripción Visita" := Descripcion;
        "To-do"."Salesperson Code" := Vendedor;
        "To-do"."Segment No." := '';
        "To-do".Type := "To-do".Type::Meeting;
        "To-do"."Date" := Fecha;
        "To-do"."Start Time" := 110000T;
        "To-do".Duration := 60 * 1000 * 30;
        "To-Do"."All Day Event" := false;
        TempEndDateTime := CreateDateTime(Fecha, "To-Do"."Start Time") + "To-Do".Duration;

        "To-Do"."Ending Date" := DT2Date(TempEndDateTime);
        if "To-Do"."All Day Event" then
            "To-Do"."Ending Time" := 0T
        else
            "To-Do"."Ending Time" := DT2Time(TempEndDateTime);
        "To-do".Status := "To-do".Status::"Not Started";
        "To-do".Priority := "To-do".Priority::Normal;
        RMSetup.ChangeCompany(Empresa);
        RMSetup.Get();
        RMSetup.TestField("To-do Nos.");
        "To-do"."No. Series" := RMSetup."To-do Nos.";

        "To-do"."Team Code" := '';
        "To-do"."Organizer To-do No." := "To-do"."No.";
        "To-do"."Last Date Modified" := Today;
        "To-do"."Last Time Modified" := Time;
        // hata que no se insertre, voy incrementando el contador
        "To-do"."No." := Copystr(Format(CurrentDateTime), 1, MaxStrLen("To-do"."No."));
        "To-do"."Organizer To-do No." := "To-do"."No.";
        "To-do".Insert;



    end;

    var
        RecReftemp: RecordRef;
        ScheduleName: Text;
        RecTemp: Record "Saldo InterEmpresa" temporary;
        gUrlPlantilla: Text;


    local procedure CrearCabecera(idInformes: Integer; var TempExcelBuffer: Record "Excel Buffer 2" temporary; var Row: Integer; var DesdeFecha: Date; var HastaFecha: Date; var FieldRef: FieldRef; var RecRefTemp: RecordRef; pemPresa: Text[30])
    var
        DF: DateFormula;
        Destinatario: Record "Destinatarios Informes";
        Empresas: Record "Empresas Informes";
        Informes: Record "Informes";
        TypeHelper: Codeunit "Type Helper";
        Campos: Record "Columnas Informes";
        Formatos: Record "Formato Columnas";
        Filtros: Record "Filtros Informes";
        ColumnCaptions: array[100] of Text[250];
        Desde: Date;
        Hasta: Date;
        FiltroFecha: Text;
    begin
        Campos.Reset();
        Informes.Get(idInformes);
        Destinatario.SetRange(Id, Informes.Id);
        Filtros.SetRange(Id, Informes.Id);
        Campos.SetRange(Id, Informes.Id);
        Empresas.SetRange(Id, Informes.Id);

        Row := 1;
        EnterCell(TempExcelBuffer, Row, 1, StrSubstNo('%1 de %2', Informes.Descripcion, DT2Date(Informes."Earliest Start Date/Time")), true, false, false, false, '', TempExcelBuffer."Cell Type"::Text, '', 0, '', '');
        Row += 1;
        EnterCell(TempExcelBuffer, Row, 1, 'Filtros:', true, false, false, false, '', TempExcelBuffer."Cell Type"::Text, '', 0, '', '');
        If Filtros.FindSet() then
            repeat
                Row += 1;
                If Filtros.Desde <> DF then DesdeFecha := CalcDate(Filtros.Desde, WorkDate()) else DesdeFecha := 0D;
                If Filtros.Hasta <> DF then HastaFecha := CalcDate(Filtros.Hasta, WorkDate()) else HastaFecha := Calcdate('99A', WorkDate());
                FieldRef := RecReftemp.Field(Filtros.Campo);

                if (filtros.Desde <> DF) or (Filtros.Hasta <> DF) then begin
                    FieldRef.SetRange(DesdeFecha, HastaFecha);
                    Rec.SetRange("Filtro Fecha", DesdeFecha, HastaFecha);
                    Desde := DesdeFecha;
                    Hasta := HastaFecha;
                    FiltroFecha := Rec.GetFilter("Filtro Fecha");
                    //'VERSION(1) SORTING(Tipo,Nº mov.) WHERE(Fecha vencimiento=FILTER(1925-04-11..2024-04-11),Cód. forma pago=FILTER(PAG. FIRMA|PAGARE))'
                    EnterCell(TempExcelBuffer, Row, 1, FieldRef.Caption, true, false, false, false, '', TempExcelBuffer."Cell Type"::Text, '', 0, '', '');
                    if DesdeFecha <> 0D then
                        EnterCell(TempExcelBuffer, Row, 2, CopyStr(TypeHelper.FormatDateWithCurrentCulture(DesdeFecha), 1, 250), false, false, false, false, '', TempExcelBuffer."Cell Type"::Text, '', 0, '', '');
                    EnterCell(TempExcelBuffer, Row, 3, CopyStr(TypeHelper.FormatDateWithCurrentCulture(HastaFecha), 1, 250), false, false, false, false, '', TempExcelBuffer."Cell Type"::Text, '', 0, '', '');
                end else begin
                    FieldRef.SetFilter(Filtros.Valor);
                    if FieldRef.Name = 'Document Type' then begin
                        Rec.SetFilter("Document Type", Filtros.Valor);
                    end;
                    EnterCell(TempExcelBuffer, Row, 1, FieldRef.Caption, true, false, false, false, '', TempExcelBuffer."Cell Type"::Text, '', 0, '', '');
                    EnterCell(TempExcelBuffer, Row, 2, Filtros.Valor, false, false, false, false, '', TempExcelBuffer."Cell Type"::Text, '', 0, '', '');
                end;
                if FieldRef.Name = 'Schedule Name' Then Schedulename := Filtros.Valor;
            //CreaFiltroCampo(TenantRecorId, RecReftemp.Number, Filtros.Campo);
            until Filtros.Next() = 0;
        Row += 1;
        FieldRef := RecReftemp.Field(Destinatario."Campo Destinatario");
        FieldRef.SetFilter(Destinatario.Valor);

        Row += 1;

        EnterCell(TempExcelBuffer, Row, 1, FieldRef.Caption, true, false, false, false, '', TempExcelBuffer."Cell Type"::Text, '', 0, '', '');
        EnterCell(TempExcelBuffer, Row, 2, Destinatario.Valor, false, false, false, false, '', TempExcelBuffer."Cell Type"::Text, '', 0, '', '');


        Calcular(true, pEmpresa, FiltroFecha);
        DesdeFecha := Desde;
        HastaFecha := Hasta;
        Campos.SetRange(Id, Informes.Id);
        Campos.SetRange(Include, true);
        Row += 1;
        if Campos.FindSet() then
            repeat
                If not formatos.Get(campos.Id, campos.Id_campo, true) then begin
                    Formatos.Init();
                    Formatos.Bold := true
                end;
                If (Campos.Funcion = Campos.Funcion::Columna) and (Campos.Titulo = '') then begin
                    DetalleColumnasCaption(ColumnCaptions);

                    EnterCell(TempExcelBuffer, Row, Campos.Orden, ColumnCaptions[Campos.campo], Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Text, Formatos."Fuente", Formatos."Tamaño", Formatos."Color Fuente", Formatos."Color Fondo");
                end else
                    EnterCell(TempExcelBuffer, Row, Campos.Orden, Campos.Titulo, Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Text, Formatos."Fuente", Formatos."Tamaño", Formatos."Color Fuente", Formatos."Color Fondo");
                if Campos."Ancho Columna" <> 0 then
                    TempExcelBuffer.SetColumnWidth(Campos.LetraColumna(Campos.Orden), Campos."Ancho Columna");
            until Campos.Next() = 0;
        Empresas.SetRange(Incluir, true);
        If Empresas.FindFirst() then
            If Empresas."Columna Excel" <> 0 then begin

                EnterCell(TempExcelBuffer, Row, Empresas."Columna Excel", 'Empresa', true, false, false, false, '', TempExcelBuffer."Cell Type"::Text, '', 0, '', '');
            end;

    end;

    local procedure Cliente_Proveedor(Value: Text; Empresa: Text[30]; SourceType: Enum "Gen. Journal Source Type"): Text
    var
        Cliente: Record "Customer";
        Proveedor: Record "Vendor";
    begin
        Cliente.ChangeCompany(Empresa);
        Proveedor.ChangeCompany(Empresa);
        If SourceType = SourceType::Customer then
            if Cliente.Get(Value) then
                exit(Cliente.Name)
            else
                exit('')
        else if SourceType = SourceType::Vendor then
            if Proveedor.Get(Value) then
                exit(Proveedor.Name)
            else
                exit('');

    end;

    procedure Cadena(SourceCode: Code[20]; Empresa: Text[30]): Text
    var
        Customer: Record Customer;

    begin
        Customer.ChangeCompany(Empresa);
        If Customer.Get(SourceCode) then
            exit(Customer."Cod cadena");

    end;



}