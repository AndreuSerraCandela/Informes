pageextension 93000 ContratosExt extends "Lista Contratos x Empresa"
{

    procedure ExportExcel(var Filtros: Record "Filtros Informes"; Var Destinatario: Record "Destinatarios Informes"; var ExcelStream: OutStream)
    var
        TempExcelBuffer: Record "Excel Buffer" temporary;
        ContratosLblEP: Label 'Contratos';
        ExcelFileNameEPR: Text;//Label 'Contratos_%1_%2';
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
        RecRef: RecordRef;
        Contrato: Record "Sales Header";
        Valor: Variant;
        FieldRef: FieldRef;
        Id: RecordId;
        FieldT: FieldType;
        Fecha: Date;
        Campo: Integer;
        NoContrato: Code[20];
        "Periodos": Record "Periodos Informes";
        Control: Codeunit ControlInformes;
        No: Text;
    begin
        TempExcelBuffer.Reset();
        TempExcelBuffer.DeleteAll();
        FechaTarea := CalcDate('1S', WorkDate());
        RecRef.Open(36, true);
        RecReftemp.Open(36, true);
        Informes.Get(Filtros.Id);
        ExcelFileNameEPR := ConvertStr(Informes.Descripcion, ' ', '_');
        if Destinatario."Nombre Informe" <> '' then
            ExcelFileNameEPR := ConvertStr(Destinatario."Nombre Informe", ' ', '_');
        Row := 1;
        EnterCell(TempExcelBuffer, Row, 1, StrSubstNo('%1 de %2', Informes.Descripcion, DT2Date(Informes."Earliest Start Date/Time")), true, false, '', TempExcelBuffer."Cell Type"::Text);
        Row += 1;
        EnterCell(TempExcelBuffer, Row, 1, 'Filtros:', true, false, '', TempExcelBuffer."Cell Type"::Text);
        If Filtros.FindSet() then
            repeat
                Row += 1;
                If Filtros.Desde <> DF then DesdeFecha := CalcDate(Filtros.Desde, WorkDate()) else DesdeFecha := 0D;
                If Filtros.Hasta <> DF then HastaFecha := CalcDate(Filtros.Hasta, WorkDate()) else HastaFecha := Calcdate('99A', WorkDate());
                FieldRef := RecReftemp.Field(Filtros.Campo);
                if (filtros.Desde <> DF) or (Filtros.Hasta <> DF) then begin
                    FieldRef.SetRange(DesdeFecha, HastaFecha);
                    EnterCell(TempExcelBuffer, Row, 1, FieldRef.Caption, true, false, '', TempExcelBuffer."Cell Type"::Text);
                    if DesdeFecha <> 0D then
                        EnterCell(TempExcelBuffer, Row, 2, CopyStr(TypeHelper.FormatDateWithCurrentCulture(DesdeFecha), 1, 250), false, false, '', TempExcelBuffer."Cell Type"::Text);
                    EnterCell(TempExcelBuffer, Row, 3, CopyStr(TypeHelper.FormatDateWithCurrentCulture(HastaFecha), 1, 250), false, false, '', TempExcelBuffer."Cell Type"::Text);
                end else begin
                    FieldRef.SetFilter(Filtros.Valor);
                    EnterCell(TempExcelBuffer, Row, 1, FieldRef.Caption, true, false, '', TempExcelBuffer."Cell Type"::Text);
                    EnterCell(TempExcelBuffer, Row, 2, Filtros.Valor, false, false, '', TempExcelBuffer."Cell Type"::Text);
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

                EnterCell(TempExcelBuffer, Row, Campos.Orden, Campos.Titulo, true, false, '', TempExcelBuffer."Cell Type"::Text);

            until Campos.Next() = 0;

        "Periodos".SetRange(Id, Informes.Id);
        If not "Periodos".FindFirst() then begin
            "Periodos".Init();
            "Periodos".Id := Informes.Id;
            Periodos.Periodo := 'Ninguno';
        end;
        REPEAT
            If Periodos.Campo2 <> 0 then begin
                iF Periodos.Semana2 then begin


                    HastaFecha := CalcDate(Format(Periodos.Hasta2) + '+' + Format(Control.Semana(WorkDate())) + 'S', WorkDate());
                    If Periodos.Desde2 <> DF then DesdeFecha := CalcDate(Format(Periodos.Desde2) + '+' + Format(Control.Semana(WorkDate())) + 'S', WorkDate());
                    FieldRef := RecReftemp.Field(Periodos.Campo2);
                    if (Periodos.Desde2 <> DF) or (Periodos.Hasta2 <> DF) then begin
                        FieldRef.SetRange(DesdeFecha, HastaFecha);
                    end;
                end else begin
                    If Periodos.Desde2 <> DF then DesdeFecha := CalcDate(Periodos.Desde2, WorkDate()) else DesdeFecha := 0D;
                    If Periodos.Hasta2 <> DF then HastaFecha := CalcDate(Periodos.Hasta2, WorkDate()) else HastaFecha := Calcdate('99A', WorkDate());
                    FieldRef := RecReftemp.Field(Periodos.Campo2);
                    if (Periodos.Desde2 <> DF) or (Periodos.Hasta2 <> DF) then begin
                        FieldRef.SetRange(DesdeFecha, HastaFecha);
                    end;
                end;
            end;
            iF Periodos.Semana then begin


                HastaFecha := CalcDate(Format(Periodos.Hasta) + '+' + Format(Control.Semana(WorkDate())) + 'S', WorkDate());
                If Periodos.Desde <> DF then DesdeFecha := CalcDate(Format(Periodos.Desde) + '+' + Format(Control.Semana(WorkDate())) + 'S', WorkDate());
                FieldRef := RecReftemp.Field(Periodos.Campo);
                if (Periodos.Desde <> DF) or (Periodos.Hasta <> DF) then begin
                    FieldRef.SetRange(DesdeFecha, HastaFecha);
                end;
            end else begin
                If Periodos.Desde <> DF then DesdeFecha := CalcDate(Periodos.Desde, WorkDate()) else DesdeFecha := 0D;
                If Periodos.Hasta <> DF then HastaFecha := CalcDate(Periodos.Hasta, WorkDate()) else HastaFecha := Calcdate('99A', WorkDate());
                FieldRef := RecReftemp.Field(Periodos.Campo);
                if (Periodos.Desde <> DF) or (Periodos.Hasta <> DF) then begin
                    FieldRef.SetRange(DesdeFecha, HastaFecha);
                end;
            end;

            procesar(true, DesdeFecha, HastaFecha);




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
                        CreateTaskFromSalesHeader(RecrefTemp.Field(Rec.FieldNo(Rec."Nº Contrato")).Value
                            , RecrefTemp.Field(Rec.FieldNo(Rec."Sell-to Contact No.")).Value
                            , RecrefTemp.Field(Rec.FieldNo(Rec."Salesperson Code")).Value
                            , RecrefTemp.Field(Rec.FieldNo(Rec."Opportunity No.")).Value
                            , RecrefTemp.Field(Rec.FieldNo(Rec."Campaign No.")).Value
                            , RecrefTemp.Field(Rec.FieldNo(Rec."Empresa del Cliente")).Value
                            , FechaTarea, Informes."Descripcion Tarea");
                    Row += 1;
                    //
                    TotalesDocumentos(RecrefTemp.Field(Rec.FieldNo("Nº Proyecto")).Value,
                     RecrefTemp.Field(Rec.FieldNo("Nº Contrato")).Value, RecrefTemp.Field(Rec.FieldNo("Empresa del Cliente")).Value);

                    if Campos.FindSet() then
                        repeat

                            rf := "Analysis Rounding Factor"::None;
                            If (Campos.Campo <> 0) and (Campos.Funcion = Campos.Funcion::" ") then begin
                                FieldRef := RecRef.Field(Campos.Campo);
                                FieldT := FieldRef.Type;
                                Valor := DevuelveCampo(Campos.Campo);
                            end else begin
                                FieldT := FieldType::Text;
                                //Importe,Vendedor,GetTotImp,ImporteIva,GetImpBorFac,GetImpBorAbo,GetImpFac,GetImpAbo,GetTotCont
                                case Campos.Funcion of
                                    Funciones::Importe:
                                        begin
                                            FieldT := FieldType::Decimal;
                                            Valor := Importe(RecrefTemp.Field(Rec.FieldNo("Nº Contrato")).Value, RecrefTemp.Field(Rec.FieldNo("Empresa del Cliente")).Value);
                                        end;
                                    Funciones::ImporteIva:
                                        begin
                                            FieldT := FieldType::Decimal;
                                            Valor := ImporteIva(RecrefTemp.Field(Rec.FieldNo("Nº Contrato")).Value, RecrefTemp.Field(Rec.FieldNo("Empresa del Cliente")).Value);

                                        end;
                                    Funciones::Vendedor:
                                        begin
                                            FieldT := FieldType::Text;
                                            Valor := Vendedor(RecrefTemp.Field(Rec.FieldNo("Salesperson Code")).Value, RecrefTemp.Field(Rec.FieldNo("Empresa del Cliente")).Value);
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
                                    funciones::Diferencia:
                                        begin
                                            FieldT := FieldType::Decimal;
                                            Valor := GetDiferencia();
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
                                    EnterCell(TempExcelBuffer, Row, Campos.Orden, Format(Valor), false, false, '', TempExcelBuffer."Cell Type"::Number);
                                FieldT::Decimal:
                                    EnterCell(TempExcelBuffer, Row, Campos.Orden, Matrix.FormatAmount(Valor, Rf, False), false, false, '', TempExcelBuffer."Cell Type"::Number);
                                FieldT::Option:
                                    EnterCell(TempExcelBuffer, Row, Campos.Orden, Valor, false, false, '', TempExcelBuffer."Cell Type"::Text);
                                FieldT::Code:
                                    EnterCell(TempExcelBuffer, Row, Campos.Orden, Valor, false, false, '', TempExcelBuffer."Cell Type"::Text);
                                FieldT::Text:
                                    EnterCell(TempExcelBuffer, Row, Campos.Orden, Valor, false, false, '', TempExcelBuffer."Cell Type"::Text);
                                FieldT::Boolean:
                                    EnterCell(TempExcelBuffer, Row, Campos.Orden, Format(Valor), false, false, '', TempExcelBuffer."Cell Type"::Text);
                                FieldT::RecordId:
                                    EnterCell(TempExcelBuffer, Row, Campos.Orden, Valor, false, false, '', TempExcelBuffer."Cell Type"::Text);
                                FieldT::Blob:
                                    EnterCell(TempExcelBuffer, Row, Campos.Orden, Valor, false, false, '', TempExcelBuffer."Cell Type"::Text);
                                FieldT::Guid:
                                    EnterCell(TempExcelBuffer, Row, Campos.Orden, Valor, false, false, '', TempExcelBuffer."Cell Type"::Text);

                            End;


                        until Campos.Next() = 0;

                    Contrato.ChangeCompany(RecrefTemp.Field(Rec.FieldNo("Empresa del Cliente")).Value);
                    NoContrato := RecrefTemp.Field(Rec.FieldNo("Nº Contrato")).Value;
                    Contrato.Get(Contrato."Document Type"::Order, NoContrato);
                    Contrato."Ofrecida ampliación" := true;
                    if Informes."Crear Tarea" then Contrato.Modify;
                //end;
                until RecrefTemp.Next() = 0;
            //RecReftemp.Close();
            RecReftemp.DeleteAll();
        until Periodos.Next() = 0;
        Informes.CalcFields("Plantilla Excel");
        if Informes."Plantilla Excel".HasValue then begin
            Informes."Plantilla Excel".CreateInStream(InExcelStream);
            TempExcelBuffer.UpdateBookStream(InExcelStream, ContratosLblEP, true);

        end else
            TempExcelBuffer.CreateNewBook(ExcelFileNameEPR);
        TempExcelBuffer.WriteSheet(ContratosLblEP, CompanyName, UserId);
        TempExcelBuffer.CloseBook();
        TempExcelBuffer.SetFriendlyFilename(StrSubstNo(ExcelFileNameEPR, CurrentDateTime, UserId));
        TempExcelBuffer.SaveToStream(ExcelStream, true);
        RecReftemp.Close();
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

    procedure CreateTaskFromSalesHeader(
        NoContrato: Code[20];
        ContactNo: Code[20];
        SalespersonCode: Code[20];
        OportunityNo: Code[20];
        CampaingNo: Code[20];
        Empresa: Text;
        Fecha: Date;
        Descripcion: Text[250])
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
        if "to-do".Get(NoContrato) then
            exit;
        "To-do".Init();
        "To-do"."Contact No." := ContactNo;

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


        "To-do".SetRange("Contact No.", ContactNo);
        if SalespersonCode <> '' then begin
            "To-do"."Salesperson Code" := SalespersonCode;

        end;
        if CampaingNo <> '' then begin
            "To-do"."Campaign No." := CampaingNo;

        end;
        "To-do".Description := Descripcion + ' ' + NoContrato;
        "To-do"."Descripción Visita" := Descripcion + ' ' + NoContrato;
        "To-do"."Salesperson Code" := SalespersonCode;
        "To-do"."Campaign No." := CampaingNo;
        "To-do"."Opportunity No." := OportunityNo;
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
        "To-do"."No." := NoContrato;
        "To-do"."Organizer To-do No." := "To-do"."No.";
        "To-do".Insert;



    end;

    var
        RecReftemp: RecordRef;


}