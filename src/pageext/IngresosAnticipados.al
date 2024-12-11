pageextension 93005 IngresosAnticipadosExt extends "Ingresos Anticipados"
{
    procedure ExportExcel(var Filtros: Record "Filtros Informes"; Idinforme: Integer;
    Var Destinatario: Record "Destinatarios Informes"; var ExcelStream: OutStream)
    var
        TempExcelBuffer: Record "Excel Buffer 2" temporary;
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
        "Periodos": Record "Periodos Informes";
        Control: Codeunit ControlInformes;
        No: Text;
        Vinculo: Text;
        LinkContrato: Record "Sales Header";
        LinkCliente: Record "Customer";
        LinkProyecto: Record "Job";
        Albaranes485: Record "Sales Header";
        Facturas485: Record "Sales Header";
        Abonos485: Record "Sales Header";
        Contabilidad485: Record "G/L Entry";
        NombreEmpresa: Text[30];
        T485: Decimal;
        F485: Decimal;
        A485: Decimal;
        DL485: Decimal;
        TempBlob: Codeunit "Temp Blob";
        Base64Convert: Codeunit "Base64 Convert";
        PlantillaBase64: Text;
    begin
        TempExcelBuffer.Reset();
        TempExcelBuffer.DeleteAll();
        FechaTarea := CalcDate('1S', WorkDate());
        RecRef.Open(36, true);
        RecReftemp.Open(36, true);
        Informes.Get(Idinforme);
        ExcelFileNameEPR := ConvertStr(Informes.Descripcion, ' ', '_');
        if Destinatario."Nombre Informe" <> '' then
            ExcelFileNameEPR := ConvertStr(Destinatario."Nombre Informe", ' ', '_');
        Row := 1;
        EnterCell(TempExcelBuffer, Row, 1, StrSubstNo('%1 de %2', Informes.Descripcion, DT2Date(Informes."Earliest Start Date/Time")), true, false, false, false, '', TempExcelBuffer."Cell Type"::Text, '', 0, '', '', '');
        Row += 1;
        EnterCell(TempExcelBuffer, Row, 1, 'Filtros:', true, false, false, false, '', TempExcelBuffer."Cell Type"::Text, '', 0, '', '', '');
        If Filtros.FindSet() then
            repeat
                Row += 1;
                If Filtros.Desde <> DF then DesdeFecha := CalcDate(Filtros.Desde, WorkDate()) else DesdeFecha := 0D;
                If Filtros.Hasta <> DF then HastaFecha := CalcDate(Filtros.Hasta, WorkDate()) else HastaFecha := Calcdate('99A', WorkDate());
                FieldRef := RecReftemp.Field(Filtros.Campo);
                if (filtros.Desde <> DF) or (Filtros.Hasta <> DF) then begin
                    FieldRef.SetRange(DesdeFecha, HastaFecha);
                    EnterCell(TempExcelBuffer, Row, 1, FieldRef.Caption, true, false, false, false, '', TempExcelBuffer."Cell Type"::Text, '', 0, '', '', '');
                    if DesdeFecha <> 0D then
                        EnterCell(TempExcelBuffer, Row, 2, CopyStr(TypeHelper.FormatDateWithCurrentCulture(DesdeFecha), 1, 250), false, false, false, false, '', TempExcelBuffer."Cell Type"::Text, '', 0, '', '', '');
                    EnterCell(TempExcelBuffer, Row, 3, CopyStr(TypeHelper.FormatDateWithCurrentCulture(HastaFecha), 1, 250), false, false, false, false, '', TempExcelBuffer."Cell Type"::Text, '', 0, '', '', '');
                end else begin
                    FieldRef.SetFilter(Filtros.Valor);
                    EnterCell(TempExcelBuffer, Row, 1, FieldRef.Caption, true, false, false, false, '', TempExcelBuffer."Cell Type"::Text, '', 0, '', '', '');
                    EnterCell(TempExcelBuffer, Row, 2, Filtros.Valor, false, false, false, false, '', TempExcelBuffer."Cell Type"::Text, '', 0, '', '', '');
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
                Vinculo := '';
                If Not Formatos.Get(campos.Id, campos.Id_campo, true) then begin
                    Formatos.Init();
                    Formatos.Bold := true;

                end;
                EnterCell(TempExcelBuffer, Row, Campos.Orden, Campos.Titulo, Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", '', TempExcelBuffer."Cell Type"::Text, Formatos.Fuente, Formatos."Tamaño", Formatos."Color Fuente", Formatos."Color Fondo", Vinculo);
                if Campos."Ancho Columna" <> 0 then
                    TempExcelBuffer.SetColumnWidth(Campos.LetraColumna(Campos.Orden), Campos."Ancho Columna");
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
                    If RecReftemp.Insert() then;
                until Rec.Next() = 0;
            RecReftemp.SetView('sorting("Document Type", "Bill-To Customer No.")');

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
                                            Valor := GetTotCont2();
                                        end;
                                    Funciones::GetTotContNew:
                                        begin
                                            FieldT := FieldType::Decimal;
                                            Valor := GetTotCont(RecrefTemp.Field(Rec.FieldNo("Estado")).Value, RecrefTemp.Field(Rec.FieldNo("Fecha Inicial Proyecto")).Value,
                                            RecrefTemp.Field(Rec.FieldNo("Posting Date")).Value);
                                        end;
                                    funciones::Diferencia:
                                        begin
                                            FieldT := FieldType::Decimal;
                                            Valor := GetDiferencia();
                                        end;
                                    Funciones::"Año":
                                        begin
                                            FieldT := FieldType::Integer;
                                            Fecha := RecrefTemp.Field(Campos.Campo).Value;
                                            If fecha <> 0D then
                                                Valor := Date2DMY(RecrefTemp.Field(Campos.Campo).Value, 3)
                                            else
                                                valor := '';

                                        end;
                                    Funciones::"Mes":
                                        begin
                                            FieldT := FieldType::Integer;
                                            Fecha := RecrefTemp.Field(Campos.Campo).Value;
                                            If fecha <> 0D then
                                                Valor := Date2DMY(RecrefTemp.Field(Campos.Campo).Value, 2)
                                            else
                                                valor := '';
                                        end;
                                    Funciones::"Semana":
                                        begin
                                            FieldT := FieldType::Integer;
                                            Fecha := RecrefTemp.Field(Campos.Campo).Value;
                                            If fecha <> 0D then
                                                Valor := Control.Semana(RecrefTemp.Field(Campos.Campo).Value)
                                            else
                                                valor := '';
                                        end;
                                    Funciones::"Diferencia 485":
                                        begin
                                            FieldT := FieldType::Decimal;
                                            //Rec."Prepmt. Payment Discount %" + Rec."Currency Factor" - Rec."VAT Base Discount %" + Rec."Invoice Discount Value"
                                            T485 := RecrefTemp.Field(Rec.FieldNo("Prepmt. Payment Discount %")).Value;
                                            F485 := RecrefTemp.Field(Rec.FieldNo("Currency Factor")).Value;
                                            DL485 := RecrefTemp.Field(Rec.FieldNo("Invoice Discount Value")).Value;
                                            A485 := RecrefTemp.Field(Rec.FieldNo("VAT Base Discount %")).Value;
                                            Valor := T485 + F485 - A485 + DL485;
                                        end;

                                    else
                                        Valor := '';
                                end;
                            end;
                            Vinculo := '';
                            If Not Formatos.Get(campos.Id, campos.Id_campo, false) then begin
                                Formatos.Init();
                                if FieldT = FieldType::Decimal then
                                    Formatos."Formato Columna" := '_-* #,##0.00_-;-* #,##0.00_-;';
                            end;
                            If Formatos."Insertar Vínculo" then begin

                                LinkCliente.ChangeCompany(RecrefTemp.Field(Rec.FieldNo("Empresa del Cliente")).Value);
                                LinkContrato.ChangeCompany(RecrefTemp.Field(Rec.FieldNo("Empresa del Cliente")).Value);
                                LinkProyecto.ChangeCompany(RecrefTemp.Field(Rec.FieldNo("Empresa del Cliente")).Value);
                                Albaranes485.ChangeCompany(RecrefTemp.Field(Rec.FieldNo("Empresa del Cliente")).Value);
                                Facturas485.ChangeCompany(RecrefTemp.Field(Rec.FieldNo("Empresa del Cliente")).Value);
                                Abonos485.ChangeCompany(RecrefTemp.Field(Rec.FieldNo("Empresa del Cliente")).Value);
                                Contabilidad485.ChangeCompany(RecrefTemp.Field(Rec.FieldNo("Empresa del Cliente")).Value);
                                NombreEmpresa := RecrefTemp.Field(Rec.FieldNo("Empresa del Cliente")).Value;
                                Case RecrefTemp.Field(Campos.Campo).Name of
                                    'Nº Contrato':
                                        begin
                                            if Not LinkContrato.Get(LinkContrato."Document Type"::Order, RecrefTemp.Field(Rec.FieldNo("Nº Contrato")).Value) then
                                                LinkContrato.Init()
                                            else
                                                LinkContrato.SetRange("No.", LinkContrato."No.");
                                            Vinculo := GetUrl(ClientType::Web, RecrefTemp.Field(Rec.FieldNo("Empresa del Cliente")).Value, ObjectType::Page, Page::"Ficha Contrato Venta", LinkContrato);
                                        end;
                                    'Nº Proyecto':
                                        begin
                                            if Not LinkProyecto.Get(RecrefTemp.Field(Rec.FieldNo("Nº Proyecto")).Value) then
                                                LinkProyecto.Init()
                                            else
                                                LinkProyecto.SetRange("No.", LinkProyecto."No.");
                                            Vinculo := GetUrl(ClientType::Web, RecrefTemp.Field(Rec.FieldNo("Empresa del Cliente")).Value, ObjectType::Page, Page::"Job Card", LinkProyecto);
                                        end;
                                    'Sell-to Contact No.':
                                        begin
                                            if Not LinkCliente.Get(RecrefTemp.Field(Rec.FieldNo("Sell-to Contact No.")).Value) then
                                                LinkCliente.Init()
                                            else
                                                LinkCliente.SetRange("No.", LinkCliente."No.");
                                            Vinculo := GetUrl(ClientType::Web, RecrefTemp.Field(Rec.FieldNo("Empresa del Cliente")).Value, ObjectType::Page, Page::"Customer Card", LinkCliente);
                                        end;



                                    'Prepmt. Payment Discount %':
                                        begin
                                            Contabilidad485.SetRange("G/L Account No.", '485', '4859999999');
                                            Contabilidad485.SetRange("Job No.", RecrefTemp.Field(Rec.FieldNo("Nº Proyecto")).Value);

                                            //https://bc220.malla.es/BC220/?company=Malla%20Publicidad&page=20&filter=%27G%2fL%20Entry%27.%27Job%20No.%27%20IS%20%27PR24-M0052%27%20AND%20%27G%2fL%20Entry%27.%27G%2fL%20Account%20No.%27%20IS%20%27485..486%27&dc=0&bookmark=C_EQAAAACH8Tsx;
                                            Vinculo := 'https://bc220.malla.es/BC220/?company=' + NombreEmpresa +
                                            '&page=20&filter=%27G%2fL%20Entry%27.%27Job%20No.%27%20IS%20%27'
                                             + Format(RecrefTemp.Field(Rec.FieldNo("Nº Proyecto")).Value) +
                                             '%27%20AND%20%27G%2fL%20Entry%27.%27G%2fL%20Account%20No.%27%20IS%20%27485..486%27';

                                        end;
                                    'Currency Factor':
                                        begin
                                            Facturas485.SetRange("Document Type", Facturas485."Document Type"::Order);
                                            Facturas485.SetRange("Nº Proyecto", RecrefTemp.Field(Rec.FieldNo("Nº Proyecto")).Value);
                                            Vinculo := GetUrl(ClientType::Web, RecrefTemp.Field(Rec.FieldNo("Empresa del Cliente")).Value, ObjectType::Page, Page::"Lista documentos venta MLL", Facturas485);
                                        end;
                                    'VAT Base Discount %':
                                        begin
                                            Abonos485.SetRange("Document Type", Abonos485."Document Type"::Order);
                                            Abonos485.SetRange("Nº Proyecto", RecrefTemp.Field(Rec.FieldNo("Nº Proyecto")).Value);
                                            Vinculo := GetUrl(ClientType::Web, RecrefTemp.Field(Rec.FieldNo("Empresa del Cliente")).Value, ObjectType::Page, Page::"Lista documentos venta MLL", Abonos485);
                                        end;
                                    'Invoice Discount Value':

                                        begin
                                            Albaranes485.SetRange("Document Type", Albaranes485."Document Type"::Order);
                                            Albaranes485.SetRange("Nº Proyecto", RecrefTemp.Field(Rec.FieldNo("Nº Proyecto")).Value);
                                            Vinculo := GetUrl(ClientType::Web, RecrefTemp.Field(Rec.FieldNo("Empresa del Cliente")).Value, ObjectType::Page, Page::"Lista documentos venta MLL", Albaranes485);
                                        end;
                                end;

                            end;

                            //end;
                            Case FieldT of
                                FieldT::Date:
                                    begin
                                        if Valor.IsDate then Fecha := Valor else Fecha := 0D;
                                        iF Fecha <> 0D then
                                            EnterCell(TempExcelBuffer, Row, Campos.Orden, CopyStr(TypeHelper.FormatDateWithCurrentCulture(Fecha), 1, 250), Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Date, Formatos.Fuente, Formatos."Tamaño", Formatos."Color Fuente", Formatos."Color Fondo", Vinculo)
                                        else
                                            EnterCell(TempExcelBuffer, Row, Campos.Orden, '', Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Text, Formatos.Fuente, Formatos."Tamaño", Formatos."Color Fuente", Formatos."Color Fondo", Vinculo);
                                    END;
                                FieldT::Time:
                                    EnterCell(TempExcelBuffer, Row, Campos.Orden, Format(Valor), Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Time, Formatos.Fuente, Formatos."Tamaño", Formatos."Color Fuente", Formatos."Color Fondo", Vinculo);
                                FieldT::Integer:
                                    EnterCell(TempExcelBuffer, Row, Campos.Orden, Format(Valor), Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Number, Formatos.Fuente, Formatos."Tamaño", Formatos."Color Fuente", Formatos."Color Fondo", Vinculo);
                                FieldT::Decimal:
                                    EnterCell(TempExcelBuffer, Row, Campos.Orden, Format(Valor), Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Number, Formatos.Fuente, Formatos."Tamaño", Formatos."Color Fuente", Formatos."Color Fondo", Vinculo);
                                //EnterCell(TempExcelBuffer, Row, Campos.Orden, Matrix.FormatAmount(Valor, Rf, False), false, false, '', TempExcelBuffer."Cell Type"::Number);
                                FieldT::Option:
                                    EnterCell(TempExcelBuffer, Row, Campos.Orden, Valor, Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Text, Formatos.Fuente, Formatos."Tamaño", Formatos."Color Fuente", Formatos."Color Fondo", Vinculo);
                                FieldT::Code:
                                    EnterCell(TempExcelBuffer, Row, Campos.Orden, Valor, Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Text, Formatos.Fuente, Formatos."Tamaño", Formatos."Color Fuente", Formatos."Color Fondo", Vinculo);
                                FieldT::Text:
                                    EnterCell(TempExcelBuffer, Row, Campos.Orden, Valor, Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Text, Formatos.Fuente, Formatos."Tamaño", Formatos."Color Fuente", Formatos."Color Fondo", Vinculo);
                                FieldT::Boolean:
                                    EnterCell(TempExcelBuffer, Row, Campos.Orden, Format(Valor), Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Text, Formatos.Fuente, Formatos."Tamaño", Formatos."Color Fuente", Formatos."Color Fondo", Vinculo);
                                FieldT::RecordId:
                                    EnterCell(TempExcelBuffer, Row, Campos.Orden, Valor, Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Text, Formatos.Fuente, Formatos."Tamaño", Formatos."Color Fuente", Formatos."Color Fondo", Vinculo);
                                FieldT::Blob:
                                    EnterCell(TempExcelBuffer, Row, Campos.Orden, Valor, Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Text, Formatos.Fuente, Formatos."Tamaño", Formatos."Color Fuente", Formatos."Color Fondo", Vinculo);
                                FieldT::Guid:
                                    EnterCell(TempExcelBuffer, Row, Campos.Orden, Valor, Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Text, Formatos.Fuente, Formatos."Tamaño", Formatos."Color Fuente", Formatos."Color Fondo", Vinculo);

                            End;


                        until Campos.Next() = 0;

                    Contrato.ChangeCompany(RecrefTemp.Field(Rec.FieldNo("Empresa del Cliente")).Value);
                    NoContrato := RecrefTemp.Field(Rec.FieldNo("Nº Contrato")).Value;
                    If Contrato.Get(Contrato."Document Type"::Order, NoContrato) Then begin
                        Contrato."Ofrecida ampliación" := true;
                        if Informes."Crear Tarea" then Contrato.Modify;
                    end;
                //end;
                until RecrefTemp.Next() = 0;
            //RecReftemp.Close();
            RecReftemp.DeleteAll();
        until Periodos.Next() = 0;
        Informes.CalcFields("Plantilla Excel");
        if (Informes."Plantilla Excel".HasValue) Or (Informes."Url Plantilla" <> '') then begin
            if Informes."Plantilla Excel".HasValue then
                Informes."Plantilla Excel".CreateInStream(InExcelStream);
            Control.UrlPlantilla(gUrlPlantilla, Informes, PlantillaBase64, false);
            if Not Informes."Formato Json" then
                TempExcelBuffer.UpdateBookStream(InExcelStream, ContratosLblEP, true);

        end else begin
            if Informes."Formato Json" then
                PlantillaBase64 := ''
            else
                TempExcelBuffer.CreateNewBook(ExcelFileNameEPR);
        end;
        if Not Informes."Formato Json" then begin
            TempExcelBuffer.WriteSheet(ContratosLblEP, CompanyName, UserId, Informes."Orientación");
            TempExcelBuffer.CloseBook();
            TempExcelBuffer.SetFriendlyFilename(StrSubstNo(ExcelFileNameEPR, CurrentDateTime, UserId));
            TempExcelBuffer.SaveToStream(ExcelStream, true);
        end else begin
            TempExcelBuffer.ModifyAll("Sheet Name", ContratosLblEP);
            PlantillaBase64 := Control.JsonExcel(TempExcelBuffer, PlantillaBase64,
            gUrlPlantilla);
            Base64Convert.FromBase64(PlantillaBase64, ExcelStream);
        end;
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
         Var TempExcelBuf: Record "Excel Buffer 2" temporary;
         RowNo: Integer;
         ColumnNo: Integer;
         CellValue: Text[250];
         Bold: Boolean;
         Italic: Boolean;
         UnderLine: Boolean;
         DobleUnderLine: Boolean;
         NumberFormat: Text;
         CellType: Option; Fuente: Text[30]; Tamaño: Integer; Color: Text; ColorFondo: Text; Vinculo: Text)
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
        TempExcelBuf.Vinculo := Vinculo;
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
        gUrlPlantilla: Text;

}