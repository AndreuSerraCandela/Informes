pageextension 93000 ContratosExt extends "Lista Contratos x Empresa"
{

    procedure ExportExcel(var Filtros: Record "Filtros Informes"; Var Destinatario: Record "Destinatarios Informes"; var ExcelStream: OutStream)
    var
        TempExcelBuffer: Record "Excel Buffer" temporary;
        ContratosLblEP: Label 'Contratos';
        ExcelFileNameEPR: Label 'Contratos_%1_%2';
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
        Campos: Record "Campos Informes";
        RecRef: RecordRef;
        Valor: Variant;
        FieldRef: FieldRef;
        Id: RecordId;
        FieldT: FieldType;
        Fecha: Date;
    begin
        TempExcelBuffer.Reset();
        TempExcelBuffer.DeleteAll();
        FechaTarea := CalcDate('1S', WorkDate());
        procesar(true);
        RecRef.Open(36, true);
        Informes.Get(Filtros.Id);
        Row := 1;
        EnterCell(TempExcelBuffer, Row, 1, StrSubstNo('%1 de %2', Informes.Descripcion, Informes."Próxima Fecha"), true, false, '', TempExcelBuffer."Cell Type"::Text);
        Row += 1;
        EnterCell(TempExcelBuffer, Row, 1, 'Filtros:', true, false, '', TempExcelBuffer."Cell Type"::Text);
        If Filtros.FindSet() then
            repeat
                Row += 1;
                If Filtros.Desde <> DF then DesdeFecha := CalcDate(Filtros.Desde, WorkDate()) else DesdeFecha := 0D;
                If Filtros.Hasta <> DF then HastaFecha := CalcDate(Filtros.Hasta, WorkDate()) else HastaFecha := Calcdate('99A', WorkDate());

                Case
                    Filtros.Campo Of
                    Rec.FieldNo("Sell-to Customer No."):
                        begin
                            Rec.SetRange("Sell-to Customer No.", Filtros.Valor);
                            EnterCell(TempExcelBuffer, Row, 1, Rec.FieldCaption("Sell-to Customer Name"), true, false, '', TempExcelBuffer."Cell Type"::Text);
                            EnterCell(TempExcelBuffer, Row, 2, Filtros.Valor, false, false, '', TempExcelBuffer."Cell Type"::Text);
                        end;
                    // Rec.FieldNo("No."):
                    //     Rec.SetFilter("No.", '%1', Filtros.Valor);
                    Rec.FieldNo("Salesperson Code"):
                        begin
                            Rec.SetFILTER("Salesperson Code", '%1', Filtros.Valor);
                            EnterCell(TempExcelBuffer, Row, 1, Rec.FieldCaption("Salesperson Code"), true, false, '', TempExcelBuffer."Cell Type"::Text);
                            EnterCell(TempExcelBuffer, Row, 2, Filtros.Valor, false, false, '', TempExcelBuffer."Cell Type"::Text);
                        end;
                    Rec.FieldNo("Posting Date"):
                        begin
                            Rec.SetRange("Posting Date", DesdeFecha, HastaFecha);
                            EnterCell(TempExcelBuffer, Row, 1, Rec.FieldCaption("Posting Date"), true, false, '', TempExcelBuffer."Cell Type"::Text);
                            if DesdeFecha <> 0D then
                                EnterCell(TempExcelBuffer, Row, 2, CopyStr(TypeHelper.FormatDateWithCurrentCulture(DesdeFecha), 1, 250), false, false, '', TempExcelBuffer."Cell Type"::Text);
                            EnterCell(TempExcelBuffer, Row, 3, CopyStr(TypeHelper.FormatDateWithCurrentCulture(HastaFecha), 1, 250), false, false, '', TempExcelBuffer."Cell Type"::Text);
                        end;
                    Rec.FieldNo("Fecha inicial proyecto"):
                        begin
                            Rec.SetRange("Fecha inicial proyecto", DesdeFecha, HastaFecha);
                            EnterCell(TempExcelBuffer, Row, 1, Rec.FieldCaption("Fecha inicial proyecto"), false, false, '', TempExcelBuffer."Cell Type"::Text);
                            if DesdeFecha <> 0D then
                                EnterCell(TempExcelBuffer, Row, 2, CopyStr(TypeHelper.FormatDateWithCurrentCulture(DesdeFecha), 1, 250), false, false, '', TempExcelBuffer."Cell Type"::Text);
                            EnterCell(TempExcelBuffer, Row, 3, CopyStr(TypeHelper.FormatDateWithCurrentCulture(HastaFecha), 1, 250), false, false, '', TempExcelBuffer."Cell Type"::Text);
                        end;
                    Rec.FieldNo("Fecha fin proyecto"):
                        begin
                            Rec.SetRange("Fecha fin proyecto", DesdeFecha, HastaFecha);
                            EnterCell(TempExcelBuffer, Row, 1, Rec.FieldCaption("Fecha fin proyecto"), true, false, '', TempExcelBuffer."Cell Type"::Text);
                            if DesdeFecha <> 0D then
                                EnterCell(TempExcelBuffer, Row, 2, CopyStr(TypeHelper.FormatDateWithCurrentCulture(DesdeFecha), 1, 250), false, false, '', TempExcelBuffer."Cell Type"::Text);
                            EnterCell(TempExcelBuffer, Row, 3, CopyStr(TypeHelper.FormatDateWithCurrentCulture(HastaFecha), 1, 250), false, false, '', TempExcelBuffer."Cell Type"::Text);
                        end;
                    Rec.FieldNo("Fecha Estado"):
                        begin
                            Rec.SetRange("Fecha Estado", DesdeFecha, HastaFecha);
                            EnterCell(TempExcelBuffer, Row, 1, Rec.FieldCaption("Fecha Estado"), true, false, '', TempExcelBuffer."Cell Type"::Text);
                            if DesdeFecha <> 0D then
                                EnterCell(TempExcelBuffer, Row, 2, CopyStr(TypeHelper.FormatDateWithCurrentCulture(DesdeFecha), 1, 250), false, false, '', TempExcelBuffer."Cell Type"::Text);
                            EnterCell(TempExcelBuffer, Row, 3, CopyStr(TypeHelper.FormatDateWithCurrentCulture(HastaFecha), 1, 250), false, false, '', TempExcelBuffer."Cell Type"::Text);
                        end;
                    rec.FieldNo("Fecha renovacion"):
                        begin
                            Rec.SetRange("Fecha renovacion", DesdeFecha, HastaFecha);
                            EnterCell(TempExcelBuffer, Row, 1, Rec.FieldCaption("Fecha renovacion"), true, false, '', TempExcelBuffer."Cell Type"::Text);
                            if DesdeFecha <> 0D then
                                EnterCell(TempExcelBuffer, Row, 2, CopyStr(TypeHelper.FormatDateWithCurrentCulture(DesdeFecha), 1, 250), false, false, '', TempExcelBuffer."Cell Type"::Text);
                            EnterCell(TempExcelBuffer, Row, 3, CopyStr(TypeHelper.FormatDateWithCurrentCulture(HastaFecha), 1, 250), false, false, '', TempExcelBuffer."Cell Type"::Text);
                        end;
                    rec.FieldNo("Ofrecida ampliación"):
                        begin
                            case Filtros.Valor of
                                'Sí':
                                    Rec.SetRange("Ofrecida ampliación", true);
                                'No':
                                    Rec.SetRange("Ofrecida ampliación", false);
                                else
                                    Rec.Setfilter("Ofrecida ampliación", Filtros.Valor);
                            end;
                            EnterCell(TempExcelBuffer, Row, 1, Rec.FieldCaption("Ofrecida ampliación"), true, false, '', TempExcelBuffer."Cell Type"::Text);
                            EnterCell(TempExcelBuffer, Row, 2, Filtros.Valor, false, false, '', TempExcelBuffer."Cell Type"::Text);
                        end;
                    Rec.FieldNo(Tipo):
                        begin
                            case Filtros.Valor of
                                'Anual':
                                    Rec.SetRange(Tipo, Rec.Tipo::Anual);
                                'Por Campaña':
                                    Rec.SetRange(Tipo, Rec.Tipo::"Por Campaña");
                                'Otros':
                                    Rec.SetRange(Tipo, Rec.Tipo::Otros);
                                'Por Temporada':
                                    Rec.SetRange(Tipo, Rec.Tipo::"Por Temporada");
                                'Reserva':
                                    Rec.SetRange(Tipo, Rec.Tipo::Reserva);
                                'Propuesta':
                                    Rec.SetRange(Tipo, Rec.Tipo::Propuesta);
                                else
                                    Rec.Setfilter(Tipo, Filtros.Valor);
                            end;
                            EnterCell(TempExcelBuffer, Row, 1, Rec.FieldCaption(Tipo), true, false, '', TempExcelBuffer."Cell Type"::Text);
                            EnterCell(TempExcelBuffer, Row, 2, Filtros.Valor, false, false, '', TempExcelBuffer."Cell Type"::Text);
                        end;
                    Rec.FieldNo(Estado):
                        begin
                            case Filtros.Valor of
                                'Pendiente de Firma':
                                    Rec.SetRange(Estado, Rec.Estado::"Pendiente de Firma");
                                'Firmado':
                                    Rec.SetRange(Estado, Rec.Estado::Firmado);
                                'Anulado':
                                    Rec.SetRange(Estado, Rec.Estado::Anulado);
                                'Cancelado':
                                    Rec.SetRange(Estado, Rec.Estado::Cancelado);
                                'Sin Montar':
                                    Rec.SetRange(Estado, Rec.Estado::"Sin Montar");
                                else
                                    Rec.Setfilter(Estado, Filtros.Valor);
                            end;
                            EnterCell(TempExcelBuffer, Row, 1, Rec.FieldCaption(Estado), true, false, '', TempExcelBuffer."Cell Type"::Text);
                            EnterCell(TempExcelBuffer, Row, 2, Filtros.Valor, false, false, '', TempExcelBuffer."Cell Type"::Text);

                        end;
                    Rec.FieldNo(Renovado):
                        begin
                            case Filtros.Valor of
                                'Sí':
                                    Rec.SetRange(Renovado, true);
                                'No':
                                    Rec.SetRange(Renovado, false);
                                else
                                    Rec.Setfilter(Renovado, Filtros.Valor);
                            end;
                            EnterCell(TempExcelBuffer, Row, 1, Rec.FieldCaption(Renovado), true, false, '', TempExcelBuffer."Cell Type"::Text);
                            EnterCell(TempExcelBuffer, Row, 2, Filtros.Valor, false, false, '', TempExcelBuffer."Cell Type"::Text);

                        end;
                End;
            until Filtros.Next() = 0;
        Case
           Destinatario."Campo Destinatario" Of
            Rec.FieldNo("Salesperson Code"):
                begin
                    Row += 1;
                    Rec.SetFILTER("Salesperson Code", '%1', Destinatario.Valor);
                    EnterCell(TempExcelBuffer, Row, 1, Rec.FieldCaption("Salesperson Code"), true, false, '', TempExcelBuffer."Cell Type"::Text);
                    EnterCell(TempExcelBuffer, Row, 2, Destinatario.Valor, false, false, '', TempExcelBuffer."Cell Type"::Text);
                end;
        end;
        Row += 1;



        // TempExcelBuffer.AddColumn(Rec.FieldCaption("No."), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        // TempExcelBuffer.AddColumn(Rec.FieldCaption(Description), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        //var ExcelBuffer: Record "Excel Buffer"; RowNo: Integer; ColumnNo: Integer; Value: Variant; IsBold: Boolean; IsItalics: Boolean; IsUnderline: Boolean)
        //Crear estos títulos: Nº Contrato	Cliente	Nombre	Importe	Descripción	Agente	Comentario	Inicio	Fin
        Campos.SetRange(Id, Informes.Id);
        if Campos.FindSet() then
            repeat

                EnterCell(TempExcelBuffer, Row, Campos.Orden, Campos.Titulo, true, false, '', TempExcelBuffer."Cell Type"::Text);

            until Campos.Next() = 0;



        if Rec.FindSet() then
            repeat
                if Informes."Crear Tarea" then
                    CreateTaskFromSalesHeader(Rec, Rec."Empresa del Cliente", FechaTarea, Informes."Descripcion Tarea");
                Row += 1;
                //
                TotalesDocumentos(Rec."Nº Proyecto", Rec."Nº Contrato", Rec."Empresa del Cliente");
                if Campos.FindSet() then
                    repeat

                        rf := "Analysis Rounding Factor"::None;
                        If Campos.Campo <> 0 then begin
                            FieldRef := RecRef.Field(Campos.Campo);
                            FieldT := FieldRef.Type;
                            Id := Rec.RecordId;
                            Valor := DevuelveValorCampo(Id, Campos.Campo);
                        end else begin
                            FieldT := FieldType::Text;
                            //Importe,Vendedor,GetTotImp,ImporteIva,GetImpBorFac,GetImpBorAbo,GetImpFac,GetImpAbo,GetTotCont
                            case Campos.Funcion of
                                Funciones::Importe:
                                    begin
                                        FieldT := FieldType::Decimal;
                                        Valor := Importe(Rec."Nº Contrato", Rec."Empresa del Cliente");
                                    end;
                                Funciones::ImporteIva:
                                    begin
                                        FieldT := FieldType::Decimal;
                                        Valor := ImporteIva(Rec."Nº Contrato", Rec."Empresa del Cliente");

                                    end;
                                Funciones::Vendedor:
                                    begin
                                        FieldT := FieldType::Text;
                                        Valor := Vendedor(Rec."Salesperson Code");
                                    end;
                                Funciones::GetTotImp:
                                    begin
                                        FieldT := FieldType::Decimal;
                                        Valor := GetTotImp();

                                    end;
                                Funciones::GetImpBorFac:
                                    begin
                                        FieldT := FieldType::Decimal;
                                        Valor := GetImpBorFac();
                                    end;
                                Funciones::GetImpBorAbo:
                                    begin
                                        FieldT := FieldType::Decimal;
                                        Valor := GetImpBorAbo();
                                    end;
                                Funciones::GetImpFac:
                                    begin
                                        FieldT := FieldType::Decimal;
                                        Valor := GetImpFac();
                                    end;
                                Funciones::GetImpAbo:
                                    begin
                                        FieldT := FieldType::Decimal;
                                        Valor := GetImpAbo();
                                    end;
                                Funciones::GetTotCont:
                                    begin
                                        FieldT := FieldType::Decimal;
                                        Valor := GetTotCont();
                                    end;
                                else
                                    Valor := '';
                            end;
                        end;
                        Case FieldT of
                            FieldT::Date:
                                begin
                                    if Valor.IsDate then Fecha := Valor else Fecha := 0D;
                                    iF Fecha <> 0D then
                                        EnterCell(TempExcelBuffer, Row, Campos.Orden, CopyStr(TypeHelper.FormatDateWithCurrentCulture(Fecha), 1, 250), false, false, '', TempExcelBuffer."Cell Type"::Date)
                                    else
                                        EnterCell(TempExcelBuffer, Row, Campos.Orden, '', false, false, '', TempExcelBuffer."Cell Type"::Text);
                                END;
                            FieldT::Time:
                                EnterCell(TempExcelBuffer, Row, Campos.Orden, Format(Valor), false, false, '', TempExcelBuffer."Cell Type"::Time);
                            FieldT::Integer:
                                EnterCell(TempExcelBuffer, Row, Campos.Orden, Valor, false, false, '', TempExcelBuffer."Cell Type"::Number);
                            FieldT::Decimal:
                                EnterCell(TempExcelBuffer, Row, Campos.Orden, Matrix.FormatAmount(Valor, Rf, False), false, false, '', TempExcelBuffer."Cell Type"::Number);
                            FieldT::Option:
                                EnterCell(TempExcelBuffer, Row, Campos.Orden, Valor, false, false, '', TempExcelBuffer."Cell Type"::Text);
                            FieldT::Code:
                                EnterCell(TempExcelBuffer, Row, Campos.Orden, Valor, false, false, '', TempExcelBuffer."Cell Type"::Text);
                            FieldT::Text:
                                EnterCell(TempExcelBuffer, Row, Campos.Orden, Valor, false, false, '', TempExcelBuffer."Cell Type"::Text);
                            FieldT::Boolean:
                                EnterCell(TempExcelBuffer, Row, Campos.Orden, Valor, false, false, '', TempExcelBuffer."Cell Type"::Text);
                            FieldT::RecordId:
                                EnterCell(TempExcelBuffer, Row, Campos.Orden, Valor, false, false, '', TempExcelBuffer."Cell Type"::Text);
                            FieldT::Blob:
                                EnterCell(TempExcelBuffer, Row, Campos.Orden, Valor, false, false, '', TempExcelBuffer."Cell Type"::Text);
                            FieldT::Guid:
                                EnterCell(TempExcelBuffer, Row, Campos.Orden, Valor, false, false, '', TempExcelBuffer."Cell Type"::Text);

                        End;


                    until Campos.Next() = 0;
            // EnterCell(TempExcelBuffer, Row, 1, Rec."Nº Contrato", false, false, '', TempExcelBuffer."Cell Type"::Text);
            // EnterCell(TempExcelBuffer, Row, 2, Rec."Sell-to Customer No.", false, false, '', TempExcelBuffer."Cell Type"::Text);
            // EnterCell(TempExcelBuffer, Row, 3, Rec."Sell-to Customer Name", false, false, '', TempExcelBuffer."Cell Type"::Text);
            // //FormatAmount(Value: Decimal; RoundingFactor: Enum "Analysis Rounding Factor"; AddCurrency: Boolean)

            // EnterCell(TempExcelBuffer, Row, 4, Matrix.FormatAmount(Importe(Rec."Nº Contrato", Rec."Empresa del Cliente"), Rf, False), false, false, '', TempExcelBuffer."Cell Type"::Number);
            // EnterCell(TempExcelBuffer, Row, 5, Rec."Posting Description", false, false, '', TempExcelBuffer."Cell Type"::Text);
            // EnterCell(TempExcelBuffer, Row, 6, Vendedor(Rec."Salesperson Code"), false, false, '', TempExcelBuffer."Cell Type"::Text);
            // EnterCell(TempExcelBuffer, Row, 7, Rec."Comentario Cabecera", false, false, '', TempExcelBuffer."Cell Type"::Text);
            // EnterCell(TempExcelBuffer, Row, 8, CopyStr(TypeHelper.FormatDateWithCurrentCulture(Rec."Fecha inicial proyecto"), 1, 250), false, false, '', TempExcelBuffer."Cell Type"::Date);
            // EnterCell(TempExcelBuffer, Row, 9, CopyStr(TypeHelper.FormatDateWithCurrentCulture(Rec."Fecha fin proyecto"), 1, 250), false, false, '', TempExcelBuffer."Cell Type"::Date);


            until Rec.Next() = 0;
        if Informes."Crear Tarea" then Rec.ModifyAll("Ofrecida ampliación", true);
        Informes.CalcFields("Plantilla Excel");
        if Informes."Plantilla Excel".HasValue then begin
            Informes."Plantilla Excel".CreateInStream(InExcelStream);
            TempExcelBuffer.UpdateBookStream(InExcelStream, ContratosLblEP, true);

        end else
            TempExcelBuffer.CreateNewBook(ContratosLblEP);
        TempExcelBuffer.WriteSheet(ContratosLblEP, CompanyName, UserId);
        TempExcelBuffer.CloseBook();
        TempExcelBuffer.SetFriendlyFilename(StrSubstNo(ExcelFileNameEPR, CurrentDateTime, UserId));
        TempExcelBuffer.SaveToStream(ExcelStream, true);
    end;

    local procedure EnterCell(
        Var TempExcelBuf: Record "Excel Buffer" temporary;
        RowNo: Integer;
        ColumnNo: Integer;
        CellValue: Text[250];
        Bold: Boolean;
        UnderLine: Boolean;
        NumberFormat: Text[30];
        CellType: Option)
    begin
        TempExcelBuf.Init();
        TempExcelBuf.Validate("Row No.", RowNo);
        TempExcelBuf.Validate("Column No.", ColumnNo);
        TempExcelBuf."Cell Value as Text" := CellValue;
        TempExcelBuf.Formula := '';
        TempExcelBuf.Bold := Bold;
        TempExcelBuf.Underline := UnderLine;
        TempExcelBuf.NumberFormat := NumberFormat;
        TempExcelBuf."Cell Type" := CellType;
        TempExcelBuf.Insert();
    end;

    procedure CreateTaskFromSalesHeader(SalesHeader: Record "Sales Header"; Empresa: Text; Fecha: Date; Descripcion: Text[250])
    var
        "To-do": Record "To-do";
        Cont: Record Contact;
        TempAttendee: Record Attendee temporary;
        RMSetup: Record "Marketing Setup";
        Ser: Record "No. Series Line";
        TempEndDateTime: DateTime;
        Contrato: Record "Sales Header";
    begin
        "To-do".ChangeCompany(Empresa);
        "To-do".Init();
        "To-do"."Contact No." := SalesHeader."Sell-to Contact No.";

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


        "To-do".SetRange("Contact No.", SalesHeader."Sell-to Contact No.");
        if SalesHeader."Salesperson Code" <> '' then begin
            "To-do"."Salesperson Code" := SalesHeader."Salesperson Code";

        end;
        if SalesHeader."Campaign No." <> '' then begin
            "To-do"."Campaign No." := SalesHeader."Campaign No.";

        end;
        "To-do".Description := Descripcion +' '+ SalesHeader."Nº Contrato";
        "To-do"."Descripción Visita" := Descripcion+' ' + SalesHeader."Nº Contrato";
        "To-do"."Salesperson Code" := SalesHeader."Salesperson Code";
        "To-do"."Campaign No." := SalesHeader."Campaign No.";
        "To-do"."Opportunity No." := SalesHeader."Opportunity No.";
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
        "To-do"."No." := SalesHeader."Nº Contrato";
        "To-do"."Organizer To-do No." := "To-do"."No.";
        "To-do".Insert;



    end;



}