pageextension 93001 MovContaExt extends "MovContabilidad"
{

    procedure ExportExcel(var Filtros: Record "Filtros Informes"; Var Destinatario: Record "Destinatarios Informes"; var ExcelStream: OutStream)
    var
        TempExcelBuffer: Record "Excel Buffer 2" temporary;
        ContabilidadLblEP: Label 'Movs. contabilidad';
        ExcelFileNameEPR: Text;
        RecordLink: Record "Record Link";
        RecordLinkMgt: Codeunit "Record Link Management";
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
        Formatos: Record "Formato Columnas";
        RecRef: RecordRef;
        Contrato: Record "Sales Header";
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
    begin

        TempExcelBuffer.Reset();
        TempExcelBuffer.DeleteAll();
        FechaTarea := CalcDate('1S', WorkDate());
        RecRef.Open(17, true);
        RecReftemp.Open(17, true);

        Informes.Get(Filtros.Id);
        ExcelFileNameEPR := ConvertStr(Informes.Descripcion, ' ', '_');
        if Destinatario."Nombre Informe" <> '' then
            ExcelFileNameEPR := ConvertStr(Destinatario."Nombre Informe", ' ', '_');
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
                    if FieldRef.Name = 'Posting Date' then begin
                        Desde := DesdeFecha;
                        Hasta := HastaFecha;
                    end;

                    EnterCell(TempExcelBuffer, Row, 1, FieldRef.Caption, true, false, false, false, '', TempExcelBuffer."Cell Type"::Text, '', 0, '', '');
                    if DesdeFecha <> 0D then
                        EnterCell(TempExcelBuffer, Row, 2, CopyStr(TypeHelper.FormatDateWithCurrentCulture(DesdeFecha), 1, 250), false, false, false, false, '', TempExcelBuffer."Cell Type"::Text, '', 0, '', '');
                    EnterCell(TempExcelBuffer, Row, 3, CopyStr(TypeHelper.FormatDateWithCurrentCulture(HastaFecha), 1, 250), false, false, false, false, '', TempExcelBuffer."Cell Type"::Text, '', 0, '', '');
                end else begin
                    FieldRef.SetFilter(Filtros.Valor);
                    EnterCell(TempExcelBuffer, Row, 1, FieldRef.Caption, true, false, false, false, '', TempExcelBuffer."Cell Type"::Text, '', 0, '', '');
                    EnterCell(TempExcelBuffer, Row, 2, Filtros.Valor, false, false, false, false, '', TempExcelBuffer."Cell Type"::Text, '', 0, '', '');
                end;
            until Filtros.Next() = 0;
        Row += 1;
        FieldRef := RecReftemp.Field(Destinatario."Campo Destinatario");
        FieldRef.SetFilter(Destinatario.Valor);
        Row += 1;



        Campos.SetRange(Id, Informes.Id);
        Campos.SetRange(Include, true);
        if Campos.FindSet() then
            repeat
                If Not Formatos.Get(campos.Id, campos.Id_campo, true) then begin
                    Formatos.Init();
                    Formatos.Bold := true;
                end;
                EnterCell(TempExcelBuffer, Row, Campos.Orden, Campos.Titulo, Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", '', TempExcelBuffer."cell type"::Text, Formatos."Fuente", Formatos.Tamaño, Formatos."Color Fuente", Formatos."Color Fondo");
                if Campos."Ancho Columna" <> 0 then
                    TempExcelBuffer.SetColumnWidth(Campos.LetraColumna(Campos.Orden), Campos."Ancho Columna");
            until Campos.Next() = 0;
        "Periodos".SetRange(Id, Informes.Id);
        If not "Periodos".FindFirst() then begin
            "Periodos".Init();
            "Periodos".Id := Informes.Id;
            Periodos.Periodo := 'Ninguno';
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
            procesar(true, Desde, Hasta);


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
                    if Informes."Crear Tarea" then
                        CreateTask(Destinatario.Valor
                           , CompanyName
                            , FechaTarea, Informes."Descripcion Tarea");
                    Row += 1;
                    //

                    if Campos.FindSet() then
                        repeat

                            rf := "Analysis Rounding Factor"::None;
                            If Campos.Campo <> 0 then begin
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
                                            Valor := Cliente_Proveedor(RecrefTemp.Field(Rec.FieldNo("Source No.")).Value, RecrefTemp.Field(Rec.FieldNo("Periodo de Pago")).Value,
                                            RecrefTemp.Field(Rec.Fieldno("Source Type")).Value);
                                        end;
                                    Funciones::Vendedor:
                                        begin
                                            FieldT := FieldType::Text;
                                            Valor := Vendedor(RecrefTemp.Field(Rec.FieldNo("Salesperson Code")).Value, RecrefTemp.Field(Rec.FieldNo("Periodo de Pago")).Value);
                                        end;
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
                                    else
                                        Valor := '';
                                end;
                            end;
                            If Not Formatos.Get(campos.Id, campos.Id_campo, false) then begin
                                Formatos.Init();
                                if FieldT = FieldT::Decimal then
                                    Formatos."Formato Columna" := '_-* #,##0.00_-;-* #,##0.00_-;_-* "-"_-;_-@_-';
                            end;
                            Case FieldT of
                                FieldT::Date:
                                    begin
                                        if Valor.IsDate then Fecha := Valor else Fecha := 0D;
                                        iF Fecha <> 0D then
                                            EnterCell(TempExcelBuffer, Row, Campos.Orden, CopyStr(TypeHelper.FormatDateWithCurrentCulture(Fecha), 1, 250), Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Date, Formatos.Fuente, Formatos.Tamaño, Formatos."Color Fuente", Formatos."Color Fondo")
                                        else
                                            EnterCell(TempExcelBuffer, Row, Campos.Orden, '', Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Text, Formatos.Fuente, Formatos.Tamaño, Formatos."Color Fuente", Formatos."Color Fondo");
                                    END;
                                FieldT::Time:
                                    EnterCell(TempExcelBuffer, Row, Campos.Orden, Format(Valor), Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Time, Formatos.Fuente, Formatos.Tamaño, Formatos."Color Fuente", Formatos."Color Fondo");
                                FieldT::Integer:
                                    EnterCell(TempExcelBuffer, Row, Campos.Orden, Format(Valor), Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Number, Formatos.Fuente, Formatos.Tamaño, Formatos."Color Fuente", Formatos."Color Fondo");
                                FieldT::Decimal:
                                    EnterCell(TempExcelBuffer, Row, Campos.Orden, Format(Valor), Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Number, Formatos.Fuente, Formatos.Tamaño, Formatos."Color Fuente", Formatos."Color Fondo");
                                // EnterCell(TempExcelBuffer, Row, Campos.Orden, Matrix.FormatAmount(Valor, Rf, False), false, false, '', TempExcelBuffer."Cell Type"::Number);
                                FieldT::Option:
                                    EnterCell(TempExcelBuffer, Row, Campos.Orden, Valor, Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Text, Formatos.Fuente, Formatos.Tamaño, Formatos."Color Fuente", Formatos."Color Fondo");
                                FieldT::Code:
                                    EnterCell(TempExcelBuffer, Row, Campos.Orden, Valor, Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Text, Formatos.Fuente, Formatos.Tamaño, Formatos."Color Fuente", Formatos."Color Fondo");
                                FieldT::Text:
                                    EnterCell(TempExcelBuffer, Row, Campos.Orden, Valor, Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Text, Formatos.Fuente, Formatos.Tamaño, Formatos."Color Fuente", Formatos."Color Fondo");
                                FieldT::Boolean:
                                    EnterCell(TempExcelBuffer, Row, Campos.Orden, Format(Valor), Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Text, Formatos.Fuente, Formatos.Tamaño, Formatos."Color Fuente", Formatos."Color Fondo");
                                FieldT::RecordId:
                                    EnterCell(TempExcelBuffer, Row, Campos.Orden, Valor, Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Text, Formatos.Fuente, Formatos.Tamaño, Formatos."Color Fuente", Formatos."Color Fondo");
                                FieldT::Blob:
                                    EnterCell(TempExcelBuffer, Row, Campos.Orden, Valor, Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Text, Formatos.Fuente, Formatos.Tamaño, Formatos."Color Fuente", Formatos."Color Fondo");
                                FieldT::Guid:
                                    EnterCell(TempExcelBuffer, Row, Campos.Orden, Valor, Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Text, Formatos.Fuente, Formatos.Tamaño, Formatos."Color Fuente", Formatos."Color Fondo");

                            End;


                        until Campos.Next() = 0;


                //end;
                until RecrefTemp.Next() = 0;

            RecReftemp.DeleteAll();
        until "Periodos".Next() = 0;
        Informes.CalcFields("Plantilla Excel");
        if Informes."Plantilla Excel".HasValue then begin
            Informes."Plantilla Excel".CreateInStream(InExcelStream);
            TempExcelBuffer.UpdateBookStream(InExcelStream, ContabilidadLblEP, true);

        end else
            TempExcelBuffer.CreateNewBook(ExcelFileNameEPR);
        TempExcelBuffer.WriteSheet(ContabilidadLblEP, CompanyName, UserId, Informes."Orientación");
        TempExcelBuffer.CloseBook();
        TempExcelBuffer.SetFriendlyFilename(StrSubstNo(ExcelFileNameEPR, CurrentDateTime, UserId));
        TempExcelBuffer.SaveToStream(ExcelStream, true);
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



}