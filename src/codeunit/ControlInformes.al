Codeunit 7001130 ControlInformes
{
    trigger OnRun()
    var

    begin

        imprimirInformes(0, CurrentDateTime, false);
    end;


    procedure imprimirInformes(IdInforme: Integer; ProximaFecha: Datetime; NoEnviar: Boolean)
    var
        Informe: Record "Informes";
        Destinatario: Record "Destinatarios Informes";
        Filtros: Record "Filtros Informes";
        Contratos: Page "Lista Contratos x Empresa";
        Conta: Page "MovContabilidad";
        Out: OutStream;
        ficheros: Record Ficheros temporary;
        Secuencia: Integer;
        MillisecondsToAdd: Integer;
        NoOfMinutes: Integer;
        Intstream: InStream;
    begin
        NoOfMinutes := 5;
        MillisecondsToAdd := NoOfMinutes;
        MillisecondsToAdd := MillisecondsToAdd * 60000;
        If IdInforme <> 0 tHEN
            Informe.SetRange("ID", IdInforme);
        If ProximaFecha <> 0DT tHEN
            Informe.SetRange("Earliest Start Date/Time", CurrentDateTime - MillisecondsToAdd, CurrentDateTime + MillisecondsToAdd);
        Informe.SetRange(Ejecutandose, false);

        If Informe.FindSet() then begin
            repeat
                Informe.Ejecutandose := true;
                Informe.Modify();
                Commit();
                Destinatario.Reset;
                Destinatario.SetRange("No enviar", false);
                Destinatario.SetRange("ID", Informe."ID");
                if Destinatario.FindSet() then begin
                    repeat
                        Filtros.Reset;
                        Filtros.SetRange("ID", Informe."ID");
                        if Not Filtros.FindSet() then Filtros.Init;
                        ficheros.Reset();
                        If ficheros.FindLast() then Secuencia := ficheros.Secuencia + 1 else Secuencia := 1;
                        ficheros.Secuencia := Secuencia;
                        ficheros."Nombre fichero" := Informe.Descripcion + '.xlsx';
                        ficheros.Proceso := 'ENVIARXLS';
                        repeat
                            ficheros.Secuencia := Secuencia;
                            Secuencia += 1;
                        Until ficheros.Insert();
                        ficheros.CalcFields(Fichero);
                        ficheros.Fichero.CreateOutStream(out);
                        case Informe.Informe Of
                            Informes::"Contratos x Empresa":
                                begin
                                    Clear(Contratos);
                                    Contratos.ExportExcel(Filtros, Destinatario, out);
                                end;
                            informes::"Estadisticas Contabilidad":
                                begin
                                    Clear(Conta);
                                    Conta.ExportExcel(Filtros, Destinatario, out);
                                end;
                            Informes::Tablas:
                                begin
                                    ExportExcel(Filtros, Destinatario, out);
                                end;
                            Informes::"Web Service":
                                begin

                                    if ExportExcelWeb(Filtros, Destinatario, out) = 'Retry' then begin
                                        //RecReftemp.Close();
                                        ExportExcelWeb(Filtros, Destinatario, out);
                                    end;

                                end;

                        end;
                        ficheros.Modify();
                        Commit();
                        if NoEnviar then begin
                            ficheros.CalcFields(Fichero);
                            ficheros.Fichero.CreateInStream(Intstream);
                            DownloadFromStream(Intstream, 'Guardar', 'C:\Temp', 'ALL Files (*.*)|*.*', ficheros."Nombre fichero");
                        end else
                            EnviaCorreoComercial(Destinatario."e-mail", ficheros, Informe.Descripcion, destinatario."Nombre Informe", Informe."ID");
                    // end;
                    until Destinatario.Next() = 0;
                end;
                IF ProximaFecha <> 0DT tHEN begin
                    Informe."Earliest Start Date/Time" := Informe.CalcNextRunTimeForRecurringReport(Informe, Informe."Earliest Start Date/Time");


                end;
                Informe.Ejecutandose := false;
                Informe.Modify;

            until Informe.Next() = 0;
        end;
    end;

    local procedure EnviaCorreoComercial(SalesPersonMail: Text; var ficheros: Record Ficheros; Informe: Text; InformeDestinatario: Text; IdInforme: Integer)


    var
        Mail: Codeunit Mail;
        Body: Text;
        Customer: Record 18;
        BigText: Text;
        REmail: Record "Email Item" temporary;
        emilesc: Enum "Email Scenario";
        rInf: Record "Company Information";
        Funciones: Codeunit "Funciones Correo PDF";
        AttachmentStream: InStream;
        out: OutStream;
        Secuencia: Integer;
        Informes: Record "Informes";
        workdescription: Text;
    begin
        rInf.Get();
        BigText := ('Estimado:');

        if InformeDestinatario <> '' then
            Informe := ConvertStr(InformeDestinatario, ' ', '_') + '_' + Format(Today(), 0, '<Year4><Month,2><Day,2>');
        //(FORMAT(cr,0,'<CHAR>') + FORMAT(lf,0,'<CHAR>')
        BigText := BigText + '<br> </br>';
        BigText := BigText + '<br> </br>';
        Informes.get(IdInforme);
        workdescription := Informes.GetDescripcionAmpliada();
        if workdescription = '' then
            BigText := BigText + 'Adjuntamos el Informe: ' + Informe;


        BigText := BigText + '<br> </br>';
        BigText := BigText + workdescription;
        //BigText:=('<br> </br>';
        BigText := BigText + '<br> </br>';
        BigText := BigText + '<br> </br>';


        BigText := BigText + '<br> </br>';
        BigText := BigText + '<br> </br>';
        BigText := BigText + ('Aprovechamos la ocasión para enviarte un cordial saludo');
        BigText := BigText + '<br> </br>';
        BigText := BigText + '<br> </br>';
        BigText := BigText + ('Atentamente');
        BigText := BigText + '<br> </br>';
        BigText := BigText + ('Dpto. Tranformación digital');
        BigText := BigText + '<br> </br>';

        BigText := BigText + (rInf.Name);
        //"Plaintext Formatted":=TRUE;
        // SendMsg.AppendBody(BigText);
        // CLEAR(BigText);
        BigText := BigText + '<br> </br>';
        BigText := BigText + '<br> </br>';
        BigText := BigText + '<img src="emailFoot.png" />';
        BigText := BigText + '<br> </br>';
        BigText := BigText + '<br> </br>';
        BigText := BigText + '<font face="Franklin Gothic Book" sice=2 color=Blue>';
        BigText := BigText + ('<b>SI NO DESEA RECIBIR MAS INFORMACION, CONTESTE ESTE E-MAIL INDICANDOLO EXPRESAMENTE</b>');
        BigText := BigText + '</font>';
        BigText := BigText + '<br> </br>';
        BigText := BigText + '<font face="Franklin Gothic Book" size=1 color=Blue>';
        BigText := BigText + ('Según la LOPD 15/199, su dirección de correo electrónico junto a los demás datos personales');
        BigText := BigText + (' que Ud. nos ha facilitado, constan en un fichero titularidad de ');
        BigText := BigText + (rInf.Name + ', cuyas finalidades son mantener la');
        BigText := BigText + (' gestión de las comunicaciones con sus clientes y con aquellas personas que solicitan');
        BigText := BigText + (' información, así como la gestión y atención de los correos entrantes o sugerencias que');
        BigText := BigText + (' se formulen a través de esta cuenta derivados de su actividad. Podrá ejercitar los derechos');
        BigText := BigText + (' de acceso, cancelación, rectificación y oposición,  dirigiéndose, por escrito a ');
        BigText := BigText + (rInf.Name + ' . ' + rInf.Address + '. ' + rInf."Post Code" + '. ' + rInf.City + '. España');

        BigText := BigText + '<br> </br>';
        //SendMsg.AppendBody(BigText);
        //CLEAR(BigText);
        BigText := BigText + ('Este correo y sus archivos asociados son privados y confidenciales y va');
        BigText := BigText + (' dirigido exclusivamente a su destinatario. Si recibe este correo sin ser');
        BigText := BigText + (' el destinatario del mismo, le rogamos proceda a su eliminación y lo ponga');
        BigText := BigText + (' en conocimiento del emisor. La difusión por cualquier medio del contenido de este');
        BigText := BigText + (' correo podría ser sancionada conforme a lo previsto en las leyes españolas. ');
        BigText := BigText + ('No se autoriza la utilización con fines comerciales o para su incorporación a ficheros');
        BigText := BigText + (' automatizados de las direcciones del emisor o del destinatario.');
        BigText := BigText + '</font>';
        //REmail.Subject := 'Pago contrato ' + NContrato;
        REmail.AddAttachment(Funciones.CargaPie(), 'emailfoot.png');
        REmail."Send to" := SalesPersonMail;
        if StrPos(SalesPersonMail, Informes.Bcc) <> 0 then
            Informes.bcc := '';
        If Informes.bcc <> '' then begin
            if StrPos(SalesPersonMail, 'andreuserra@malla.es') = 0 then
                REmail."Send BCC" := Informes.bcc + ';andreuserra@malla.es'
            else
                REmail."Send BCC" := Informes.bcc;
        end else begin
            if StrPos(SalesPersonMail, 'andreuserra@malla.es') = 0 then
                REmail."Send BCC" := 'andreuserra@malla.es';
        end;
        REmail.SetBodyText(BigText);
        REmail."From Name" := UserId;
        REmail.Subject := Informe;


        ficheros.CalcFields(Fichero);
        ficheros.Fichero.CreateInStream(AttachmentStream);
        REmail.AddAttachment(AttachmentStream, Informe + '.xlsx');

        // if REmail."From Address" <> '' Then
        //     REmail."Send BCC" := REmail."From Address" else
        //     REmail."Send BCC" := BCC();
        REmail.Send(true, emilesc::Informes);
        ficheros.delete;

    end;

    procedure ExportExcel(var Filtros: Record "Filtros Informes"; Var Destinatario: Record "Destinatarios Informes"; var ExcelStream: OutStream)
    var
        TempExcelBuffer: Record "Excel Buffer" temporary;
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
        Valor: Variant;
        FieldRef: FieldRef;
        Id: RecordId;
        FieldT: FieldType;
        Fecha: Date;
        Campo: Integer;
        Tarea: Code[20];
        Contacto: Code[20];
        Vendedor: Code[20];
        ExcelFileNameEPR: Text;//Label '%1_%2_%3';
        RecrefTemp: RecordRef;

    begin

        TempExcelBuffer.Reset();
        TempExcelBuffer.DeleteAll();
        FechaTarea := CalcDate('1S', WorkDate());
        Informes.Get(Filtros.Id);
        Campos.SetRange(Id, Informes.Id);
        Campos.SetFilter(Table, '<>%1', 0);
        Campos.FindFirst();
        RecReftemp.Open(Campos.Table);
        Campos.Reset();
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

        EnterCell(TempExcelBuffer, Row, 1, FieldRef.Caption, true, false, '', TempExcelBuffer."Cell Type"::Text);
        EnterCell(TempExcelBuffer, Row, 2, Destinatario.Valor, false, false, '', TempExcelBuffer."Cell Type"::Text);
        Row += 1;
        Campos.SetRange(Id, Informes.Id);
        Campos.SetRange(Include, true);
        if Campos.FindSet() then
            repeat

                EnterCell(TempExcelBuffer, Row, Campos.Orden, Campos.Titulo, true, false, '', TempExcelBuffer."Cell Type"::Text);

            until Campos.Next() = 0;



        if RecReftemp.FindSet() then
            repeat
                if Informes."Crear Tarea" then begin
                    Campos.SetRange("Field Name", 'Document No.');
                    Campos.FindSet();
                    Tarea := RecrefTemp.Field(Campos.Campo).Value;
                    Campos.SetRange("Field Name", 'Sell-to Contact No.');
                    Campos.FindSet();
                    Contacto := RecrefTemp.Field(Campos.Campo).Value;
                    Campos.SetRange("Field Name", 'Salesperson Code');
                    Campos.FindSet();
                    Vendedor := RecrefTemp.Field(Campos.Campo).Value;
                    Campos.SetRange("Field Name");
                    CreateTask(Tarea, Contacto, Vendedor, '', '', CompanyName
                        , FechaTarea, Informes."Descripcion Tarea");
                end;
                Row += 1;
                //

                if Campos.FindSet() then
                    repeat

                        rf := "Analysis Rounding Factor"::None;
                        If Campos.Campo <> 0 then begin
                            FieldRef := RecReftemp.Field(Campos.Campo);
                            FieldT := FieldRef.Type;
                            Valor := DevuelveCampo(Campos.Campo, RecrefTemp);
                        end else begin
                            FieldT := FieldType::Text;
                            //Importe,Vendedor,GetTotImp,ImporteIva,GetImpBorFac,GetImpBorAbo,GetImpFac,GetImpAbo,GetTotCont
                            // case Campos.Funcion of
                            //     Funciones::Importe:
                            //         begin
                            //             FieldT := FieldType::Decimal;
                            //             Valor := Importe(RecrefTemp.Field(Rec.FieldNo("Nº Contrato")).Value, RecrefTemp.Field(Rec.FieldNo("Empresa del Cliente")).Value);
                            //         end;
                            //     Funciones::ImporteIva:
                            //         begin
                            //             FieldT := FieldType::Decimal;
                            //             Valor := ImporteIva(RecrefTemp.Field(Rec.FieldNo("Nº Contrato")).Value, RecrefTemp.Field(Rec.FieldNo("Empresa del Cliente")).Value);

                            //         end;
                            //     Funciones::Vendedor:
                            //         begin
                            //             FieldT := FieldType::Text;
                            //             Valor := Vendedor(RecrefTemp.Field(Rec.FieldNo("Salesperson Code")).Value);
                            //         end;
                            //     Funciones::GetTotImp:
                            //         begin
                            //             FieldT := FieldType::Decimal;
                            //             Valor := GetTotImp();

                            //         end;
                            //     Funciones::GetImpBorFac:
                            //         begin
                            //             FieldT := FieldType::Decimal;
                            //             Valor := GetImpBorFac();
                            //         end;
                            //     Funciones::GetImpBorAbo:
                            //         begin
                            //             FieldT := FieldType::Decimal;
                            //             Valor := GetImpBorAbo();
                            //         end;
                            //     Funciones::GetImpFac:
                            //         begin
                            //             FieldT := FieldType::Decimal;
                            //             Valor := GetImpFac();
                            //         end;
                            //     Funciones::GetImpAbo:
                            //         begin
                            //             FieldT := FieldType::Decimal;
                            //             Valor := GetImpAbo();
                            //         end;
                            //     Funciones::GetTotCont:
                            //         begin
                            //             FieldT := FieldType::Decimal;
                            //             Valor := GetTotCont();
                            //         end;
                            //     else
                            //         Valor := '';
                            // end;
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

            //end;
            until RecrefTemp.Next() = 0;
        RecrefTemp.Close();
        Informes.CalcFields("Plantilla Excel");
        ExcelFileNameEPR := ConvertStr(Informes.Descripcion, ' ', '_');
        if Destinatario."Nombre Informe" <> '' then
            ExcelFileNameEPR := ConvertStr(Destinatario."Nombre Informe", ' ', '_');
        if Informes."Plantilla Excel".HasValue then begin
            Informes."Plantilla Excel".CreateInStream(InExcelStream);

            TempExcelBuffer.UpdateBookStream(InExcelStream, ConvertStr(Informes.Descripcion, ' ', '_'), true);

        end else
            TempExcelBuffer.CreateNewBook(ExcelFileNameEPR);
        TempExcelBuffer.WriteSheet(ConvertStr(Informes.Descripcion, ' ', '_'), CompanyName, UserId);
        TempExcelBuffer.CloseBook();
        TempExcelBuffer.SetFriendlyFilename(StrSubstNo(ExcelFileNameEPR, ConvertStr(Informes.Descripcion, ' ', '_'), CurrentDateTime, UserId));
        TempExcelBuffer.SaveToStream(ExcelStream, true);
    end;

    procedure ExportExcelWeb(var Filtros: Record "Filtros Informes"; Var Destinatario: Record "Destinatarios Informes"; var ExcelStream: OutStream): text
    var
        TempExcelBuffer: Record "Excel Buffer" temporary;
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
        Valor: Variant;
        FieldRef: FieldRef;
        Id: RecordId;
        FieldT: FieldType;

        Campo: Integer;
        Tarea: Code[20];
        Contacto: Code[20];
        Vendedor: Code[20];
        ExcelFileNameEPR: Label '%1_%2_%3';
        OdataUtility: Codeunit "ODataUtility";
        WebService: Codeunit "Web Service Management";
        ServiceParamSelect: Text;
        ServiceParamfilter: Text;
        TenantWebServiceFilter: Record "Tenant Web Service Filter";
        TenantWebService: Record "Tenant Web Service";
        TenantWebServiceColumns: Record "Tenant Web Service Columns";
        Url: Text;
        C: Label '''';
        Jsontext: Text;
        RestApi: Codeunit ControlInformes;
        ResuestType: Option Get,patch,put,post,delete;
        JsonObj: JsonObject;
        JsonArrayToken: JsonToken;
        JsonArray: JsonArray;
        jsonTokenFila: JsonToken;
        empplyee: Record "Employee";
        SalesPerson: Record "Salesperson/Purchaser";
        Periodos: Record "Periodos Informes";
        Empresas: Record "Empresas Informes";
        cUST: Record "Customer";
        FiltroTexto: Text;
        FiltroTexto2: Text;
        HojasSeparadas: Boolean;
        Fila: Integer;
        DateValue: Date;
        Primeravez: Boolean;
        RecReftemp: RecordRef;
    begin

        TempExcelBuffer.Reset();
        TempExcelBuffer.DeleteAll();
        FechaTarea := CalcDate('1S', WorkDate());
        Informes.Get(Filtros.Id);
        Campos.SetRange(Id, Informes.Id);

        Empresas.SetRange(Id, Informes.Id);
        Campos.SetFilter(Table, '<>%1', 0);

        Campos.FindFirst();
        RecReftemp.Close();
        RecReftemp.Open(Campos.Table);
        Campos.SetRange(Table);
        Campos.ModifyAll(Table, RecReftemp.Number);
        If not Empresas.FindSet() then
            CrearCabecera(Informes.Id, TempExcelBuffer, Row, DesdeFecha, HastaFecha, FieldRef, TenantWebService.RecordId, RecReftemp)
        else begin
            If Filtros.FindSet() then
                repeat
                    Row += 1;
                    If Filtros.Desde <> DF then DesdeFecha := CalcDate(Filtros.Desde, WorkDate()) else DesdeFecha := 0D;
                    If Filtros.Hasta <> DF then HastaFecha := CalcDate(Filtros.Hasta, WorkDate()) else HastaFecha := Calcdate('99A', WorkDate());
                    FieldRef := RecReftemp.Field(Filtros.Campo);

                    if (filtros.Desde <> DF) or (Filtros.Hasta <> DF) then begin
                        FieldRef.SetRange(DesdeFecha, HastaFecha);
                        //'VERSION(1) SORTING(Tipo,Nº mov.) WHERE(Fecha vencimiento=FILTER(1925-04-11..2024-04-11),Cód. forma pago=FILTER(PAG. FIRMA|PAGARE))'

                    end else begin
                        FieldRef.SetFilter(Filtros.Valor);

                    end;
                //CreaFiltroCampo(TenantRecorId, RecReftemp.Number, Filtros.Campo);
                until Filtros.Next() = 0;
            Row += 1;
            FieldRef := RecReftemp.Field(Destinatario."Campo Destinatario");
            FieldRef.SetFilter(Destinatario.Valor);
        end;
        if Empresas.FindSet() then
            if not Empresas."Hojas Separadas" then
                CrearCabecera(Informes.Id, TempExcelBuffer, Row, DesdeFecha, HastaFecha, FieldRef, TenantWebService.RecordId, RecReftemp)
            else begin
                If Filtros.FindSet() then
                    repeat
                        Row += 1;
                        If Filtros.Desde <> DF then DesdeFecha := CalcDate(Filtros.Desde, WorkDate()) else DesdeFecha := 0D;
                        If Filtros.Hasta <> DF then HastaFecha := CalcDate(Filtros.Hasta, WorkDate()) else HastaFecha := Calcdate('99A', WorkDate());
                        FieldRef := RecReftemp.Field(Filtros.Campo);

                        if (filtros.Desde <> DF) or (Filtros.Hasta <> DF) then begin
                            FieldRef.SetRange(DesdeFecha, HastaFecha);
                            //'VERSION(1) SORTING(Tipo,Nº mov.) WHERE(Fecha vencimiento=FILTER(1925-04-11..2024-04-11),Cód. forma pago=FILTER(PAG. FIRMA|PAGARE))'

                        end else begin
                            FieldRef.SetFilter(Filtros.Valor);

                        end;
                    //CreaFiltroCampo(TenantRecorId, RecReftemp.Number, Filtros.Campo);
                    until Filtros.Next() = 0;
            end;

        WebService.CreateTenantWebService(Informes."Tipo Objeto", Informes."Id Objeto", Informes.Descripcion, true);
        TenantWebService.get(Informes."Tipo Objeto", Informes.Descripcion);

        //PasaFiltros

        Primeravez := true;
        //ServiceParamfilter := '$filter=';
        if Filtros.FindFirst() then
            repeat

                if Informes."Tipo Objeto" = Informes."Tipo Objeto"::Page then
                    WebService.CreateTenantWebServiceColumnForPage(TenantWebService.RecordId, filtros.Campo, campos.Table);
                FieldRef := RecReftemp.Field(Filtros.Campo);
                If FieldRef.Class = Fieldref.Class::FlowFilter then begin
                    if not Primeravez then
                        ServiceParamfilter := ServiceParamfilter + ' and ';
                    Primeravez := false;
                    if FieldRef.Type = FieldRef.Type::Date then begin
                        If FieldRef.GetRangeMin() <> FieldRef.GetRangeMax then
                            ServiceParamfilter += ExternalizeName(FieldRef.Name) + ' ge ' + '''' + Format(FieldRef.GetRangeMin(), 0, '<Month,2>/<Day,2>/<Year4>') + '''' + ' and ' + ExternalizeName(FieldRef.Name) + ' le ' + '''' + Format(FieldRef.GetRangeMax(), 0, '<Month,2>/<Day,2>/<Year4>') + ''''
                        else
                            ServiceParamfilter += ExternalizeName(FieldRef.Name) + ' eq ' + '''' + FieldRef.GetFilter() + '''';
                    end else
                        ServiceParamfilter += ExternalizeName(FieldRef.Name) + ' eq ' + '''' + FieldRef.GetFilter() + '''';


                end;
            until Filtros.Next() = 0;
        FiltroTexto := ServiceParamfilter;
        CreateTenantWebServiceFilterFromRecordRef(TenantWebServiceFilter, RecReftemp, TenantWebService.RecordId);

        OdataUtility.GenerateODataV4FilterText(TenantWebService."Service Name", Informes."Tipo Objeto", ServiceParamfilter);
        if FiltroTexto <> '' then begin
            if ServiceParamfilter = ''
            then
                ServiceParamfilter := '$filter=' + FiltroTexto
            else
                ServiceParamfilter += ' and ' + FiltroTexto;
        end;

        HojasSeparadas := false;
        Empresas.SetRange(Id, Informes.Id);

        If not Empresas.FindSet() then begin
            Empresas.Init();
            Empresas.Empresa := CompanyName;
        end else begin
            HojasSeparadas := Empresas."Hojas Separadas";
            if HojasSeparadas then begin

                if Informes."Plantilla Excel".HasValue then begin
                    Informes."Plantilla Excel".CreateInStream(InExcelStream);
                    TempExcelBuffer.UpdateBookStream(InExcelStream, ConvertStr(Empresas.Empresa, ' ', '_'), true);

                end else
                    TempExcelBuffer.CreateNewBook(ConvertStr(Empresas.Empresa, ' ', '_'));


            end;
        end;
        repeat

            if HojasSeparadas then begin

                Row := 0;
                CrearCabecera(Informes.Id, TempExcelBuffer, Row, DesdeFecha, HastaFecha, FieldRef, TenantWebService.RecordId, RecReftemp);
                Empresas.TestField("HojaExcel");
            end;
            Periodos.SetRange(Id, Informes.Id);
            if not Periodos.FindSet() then begin
                Periodos.Init();
                Periodos.Periodo := 'Ninguno';

            end;
            repeat
                if Periodos.Periodo <> 'Ninguno' then begin
                    iF Periodos.Semana then begin
                        HastaFecha := CalcDate(Format(Periodos.Hasta) + '+' + Format(Semana(WorkDate())) + 'S', WorkDate());
                        If Periodos.Desde <> DF then DesdeFecha := CalcDate(Format(Periodos.Desde) + '+' + Format(Semana(WorkDate())) + 'S', WorkDate());
                        FieldRef := RecReftemp.Field(Periodos.Campo);
                        if (Periodos.Desde <> DF) or (Periodos.Hasta <> DF) then begin
                            FieldRef.SetRange(DesdeFecha, HastaFecha);
                        end;
                    end else begin
                        If Periodos.Desde <> DF then DesdeFecha := CalcDate(Periodos.Desde, WorkDate()) else DesdeFecha := 0D;
                        If Periodos.Hasta <> DF then HastaFecha := CalcDate(Periodos.Hasta, WorkDate()) else HastaFecha := Calcdate('99A', WorkDate());
                    end;
                    FieldRef := RecReftemp.Field(Periodos.Campo);
                    if (Periodos.Desde <> DF) or (Periodos.Hasta <> DF) then
                        FieldRef.SetRange(DesdeFecha, HastaFecha);
                    CreateTenantWebServiceFilterFromRecordRef(TenantWebServiceFilter, RecReftemp, TenantWebService.RecordId);
                    OdataUtility.GenerateODataV4FilterText(TenantWebService."Service Name", Informes."Tipo Objeto", ServiceParamfilter);
                end;
                if FiltroTexto <> '' then begin
                    if ServiceParamfilter = '' then
                        ServiceParamfilter := '$filter=' + FiltroTexto
                    else
                        ServiceParamfilter += ' and ' + FiltroTexto;
                end;
                url := BuildUrl('https://bc220.malla.es/powerbi/ODataV4/Company(' + c + Empresas.Empresa + c + ')/' + Informes.Descripcion, ServiceParamSelect, ServiceParamfilter);
                Clear(RestApi);
                Clear(Jsontext);
                Jsontext := RestApi.RestApi(url, ResuestType::GET, '', 'pi', 'Ib6343ds.');
                if Jsontext = 'Retry' then begin
                    exit('Retry');
                    Clear(RestApi);
                    Jsontext := RestApi.RestApi(url, ResuestType::GET, '', 'pi', 'Ib6343ds.');
                end;
                JsonObj.ReadFrom(Jsontext);
                JsonObj.Get('value', JsonArrayToken);
                JsonArray := JsonArrayToken.AsArray();
                foreach jsonTokenFila in jsonarray do begin
                    // if RecReftemp.FindSet() then
                    //     repeat
                    JsonObj := jsonTokenFila.AsObject();
                    if Informes."Crear Tarea" then begin
                        Campos.SetRange(Id_Campo, Informes."Campo Tarea");
                        Campos.FindSet();

                        Tarea := DevuelveCampo(Campos."Field Name", JsonObj, FieldT);
                        Campos.SetFilter("Field Name", '%1', '*Contact No.');
                        If Campos.FindSet() Then begin
                            FieldRef := RecReftemp.Field(Campos.Campo);
                            FieldT := FieldRef.Type;
                            Contacto := DevuelveCampo(Campos."Field Name", JsonObj, FieldT);

                        end else
                            Contacto := '';
                        If Destinatario.FindFirst() then
                            repeat
                                empplyee.Get(Destinatario.Empleado);
                                SalesPerson.SetRange("Code", empplyee."Salespers./Purch. Code");
                                If SalesPerson.FindSet() Then
                                    Vendedor := SalesPerson.Code;
                                CreateTask(Tarea, Contacto, Vendedor, '', '', CompanyName
                                    , FechaTarea, Informes."Descripcion Tarea");
                            until Destinatario.Next() = 0;
                    end;
                    Row += 1;
                    if Empresas."Columna Excel" <> 0 then
                        EnterCell(TempExcelBuffer, Row, Empresas."Columna Excel", Empresas.Empresa, false, false, '', TempExcelBuffer."Cell Type"::Text);
                    //
                    Campos.SetRange(Include, true);

                    if Campos.FindSet() then
                        repeat

                            rf := "Analysis Rounding Factor"::None;
                            If (Campos.Campo <> 0) and (Campos.Funcion = Campos.Funcion::" ") then begin
                                if RecReftemp.FieldExist(Campos.Campo) then begin

                                    FieldRef := RecReftemp.Field(Campos.Campo);
                                    FieldT := FieldRef.Type;
                                end else
                                    FieldT := FieldType::Text;
                                TextValue := DevuelveCampo(Campos."Field Name", JsonObj, Fieldt);
                            end else begin
                                If Campos.Campo <> 0 Then begin
                                    if RecReftemp.FieldExist(Campos.Campo) then begin
                                        FieldRef := RecReftemp.Field(Campos.Campo);
                                        FieldT := FieldRef.Type;
                                    end else
                                        FieldT := FieldType::Text;
                                    Case Campos.Funcion of
                                        Campos.Funcion::Cliente_Proveedor:
                                            begin
                                                FieldT := FieldType::Text;
                                                TextValue := DevuelveCampo(Campos."Field Name", JsonObj, Fieldt);
                                                cUST.ChangeCompany(Empresas.Empresa);
                                                iF cUST.Get(TextValue) then
                                                    TextValue := cUST.Name;
                                            end;
                                        Campos.Funcion::"Año":
                                            begin
                                                FieldT := FieldType::Integer;
                                                Evaluate(DateValue, DevuelveCampo(Campos."Field Name", JsonObj, Fieldt));
                                                TextValue := Format(Date2DMY(DateValue, 3));

                                            end;
                                        Campos.Funcion::"Mes":
                                            begin
                                                FieldT := FieldType::Integer;
                                                Evaluate(DateValue, DevuelveCampo(Campos."Field Name", JsonObj, Fieldt));
                                                TextValue := Format(Date2DMY(DateValue, 2));
                                            end;
                                        Campos.Funcion::"Semana":
                                            begin
                                                FieldT := FieldType::Integer;
                                                Evaluate(DateValue, DevuelveCampo(Campos."Field Name", JsonObj, Fieldt));
                                                TextValue := Format(Semana(DateValue));
                                            end;
                                        Campos.Funcion::Vendedor:
                                            begin

                                                FieldT := FieldType::Text;
                                                TextValue := DevuelveCampo(Campos."Field Name", JsonObj, Fieldt);
                                                TextValue := Comercial(Empresas.Empresa, TextValue);

                                            end;
                                    End;

                                end;
                                FieldT := FieldType::Text;

                            end;
                            Case FieldT of
                                FieldT::Date:
                                    begin


                                        iF Fecha <> 0D then
                                            EnterCell(TempExcelBuffer, Row, Campos.Orden, CopyStr(TypeHelper.FormatDateWithCurrentCulture(Fecha), 1, 250), false, false, '', TempExcelBuffer."Cell Type"::Date)
                                        else
                                            EnterCell(TempExcelBuffer, Row, Campos.Orden, '', false, false, '', TempExcelBuffer."Cell Type"::Text);
                                    END;
                                FieldT::Time:
                                    EnterCell(TempExcelBuffer, Row, Campos.Orden, TextValue, false, false, '', TempExcelBuffer."Cell Type"::Time);
                                FieldT::Integer:
                                    EnterCell(TempExcelBuffer, Row, Campos.Orden, TextValue, false, false, '', TempExcelBuffer."Cell Type"::Number);
                                FieldT::Decimal:
                                    EnterCell(TempExcelBuffer, Row, Campos.Orden, TextValue, false, false, '', TempExcelBuffer."Cell Type"::Number);
                                FieldT::Option:
                                    EnterCell(TempExcelBuffer, Row, Campos.Orden, TextValue, false, false, '', TempExcelBuffer."Cell Type"::Text);
                                FieldT::Code:
                                    EnterCell(TempExcelBuffer, Row, Campos.Orden, TextValue, false, false, '', TempExcelBuffer."Cell Type"::Text);
                                FieldT::Text:
                                    EnterCell(TempExcelBuffer, Row, Campos.Orden, TextValue, false, false, '', TempExcelBuffer."Cell Type"::Text);
                                FieldT::Boolean:
                                    EnterCell(TempExcelBuffer, Row, Campos.Orden, TextValue, false, false, '', TempExcelBuffer."Cell Type"::Text);
                                FieldT::RecordId:
                                    EnterCell(TempExcelBuffer, Row, Campos.Orden, TextValue, false, false, '', TempExcelBuffer."Cell Type"::Text);
                                FieldT::Blob:
                                    EnterCell(TempExcelBuffer, Row, Campos.Orden, TextValue, false, false, '', TempExcelBuffer."Cell Type"::Text);
                                FieldT::Guid:
                                    EnterCell(TempExcelBuffer, Row, Campos.Orden, TextValue, false, false, '', TempExcelBuffer."Cell Type"::Text);

                            End;


                        until Campos.Next() = 0;

                end;
            until Periodos.Next() = 0;
            if HojasSeparadas then begin
                TempExcelBuffer.SelectOrAddSheet(ConvertStr(Empresas.HojaExcel, ' ', '_'));
                TempExcelBuffer.WriteSheet(ConvertStr(Empresas.HojaExcel, ' ', '_'), Empresas.Empresa, UserId);
                TempExcelBuffer.DeleteAll();
            end;

        until Empresas.Next() = 0;
        Informes.CalcFields("Plantilla Excel");
        if Informes."Plantilla Excel".HasValue then begin
            Informes."Plantilla Excel".CreateInStream(InExcelStream);
            if not HojasSeparadas then
                TempExcelBuffer.UpdateBookStream(InExcelStream, ConvertStr(Informes.Descripcion, ' ', '_'), true);

        end else
            if not HojasSeparadas then
                TempExcelBuffer.CreateNewBook(ConvertStr(Informes.Descripcion, ' ', '_'));
        if not HojasSeparadas then
            TempExcelBuffer.WriteSheet(ConvertStr(Informes.Descripcion, ' ', '_'), CompanyName, UserId);
        TempExcelBuffer.CloseBook();
        TempExcelBuffer.SetFriendlyFilename(StrSubstNo(ExcelFileNameEPR, ConvertStr(Informes.Descripcion, ' ', '_'), CurrentDateTime, UserId));
        TempExcelBuffer.SaveToStream(ExcelStream, true);
        exit('OK');
    end;

    Procedure DevuelveCampo(Campo: Integer; var RecRefTemp: RecordRef) Valor: Variant
    var
        MyFieldRef: FieldRef;
    begin


        MyFieldRef := RecRefTemp.Field(Campo);
        If MyFieldRef.Type = FieldType::Option Then begin
            Exit(MyFieldRef.GetEnumValueNameFromOrdinalValue(MyFieldRef.Value));
        end;

        Exit(MyFieldRef.Value);
    end;




    Procedure DevuelveCampo(Campo: Text; JsonObj: JsonObject; Tipo: FieldType) ValorText: Text
    var
        MyFieldRef: JsonToken;
        Valor: Variant;
    begin

        Campo := ExternalizeName(Campo);
        if JsonObj.Get(Campo, MyFieldRef) then begin
            case Tipo of
                Fieldtype::Date:
                    begin
                        Valor := MyFieldRef.AsValue().AsDate();
                        Fecha := 0D;
                        if Valor.IsDate then Fecha := Valor;
                        exit('');
                    end;
                FieldType::DateTime:
                    begin
                        Valor := MyFieldRef.AsValue().AsDateTime();
                        exit(Format(Valor));
                    end;

                FieldType::Time:
                    begin
                        Valor := MyFieldRef.AsValue().AsTime();
                        exit(Format(Valor));
                    end;
                FieldType::Integer:
                    begin
                        Valor := MyFieldRef.AsValue().AsInteger();
                        exit(Format(Valor));
                    end;
                FieldType::Decimal:
                    begin
                        Valor := MyFieldRef.AsValue().AsDecimal();
                        exit(Format(Valor));
                    end;
                FieldType::Option:
                    begin
                        Valor := MyFieldRef.AsValue().AsText();
                        exit(Format(Valor));
                    end;
                FieldType::Code:
                    begin
                        Valor := MyFieldRef.AsValue().AsText();
                        exit(Format(Valor));
                    end;
                FieldType::Text:
                    begin
                        Valor := MyFieldRef.AsValue().AsText();
                        exit(Format(Valor));
                    end;
                FieldType::Boolean:
                    begin
                        Valor := MyFieldRef.AsValue().AsBoolean();
                        exit(Format(Valor));
                    end;
                FieldType::RecordId:
                    begin
                        Valor := MyFieldRef.AsValue().AsText();
                        exit(Format(Valor));
                    end;
                FieldType::Blob:
                    begin
                        Valor := MyFieldRef.AsValue().AsText();
                        exit(Format(Valor));
                    end;
                FieldType::Guid:
                    begin
                        Valor := MyFieldRef.AsValue().AsText();
                        exit(Format(Valor));
                    end;

            end;
        end;


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

    procedure CreateTask(
        Tarea: Code[20];
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
        if "to-do".Get(Tarea) then
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
        "To-do".Description := Descripcion;
        "To-do"."Descripción Visita" := Descripcion;
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
        "To-do"."No." := Tarea;
        "To-do"."Organizer To-do No." := "To-do"."No.";
        "To-do".Insert;



    end;

    procedure GenerateSelectText(var SelectTextParam: Text; Id: Integer)
    var
        TenantWebServiceColumns: Record "Columnas Informes";
        FirstColumn: Boolean;
    begin

        FirstColumn := true;
        TenantWebServiceColumns.SetRange(Id, Id);
        TenantWebServiceColumns.SetRange(Include, true);

        if TenantWebServiceColumns.Find('-') then begin
            SelectTextParam := '$select=';
            repeat
                if not FirstColumn then
                    SelectTextParam += ','
                else
                    FirstColumn := false;

                SelectTextParam += ExternalizeName(TenantWebServiceColumns."Field Name");
            until TenantWebServiceColumns.Next() = 0;
        end;
    end;

    procedure ExternalizeName(Name: Text) ConvertedName: Text
    var
        CurrentPosition: Integer;
    begin
        ConvertedName := Name;

        // Mimics the behavior of the compiler when converting a field or web service name to OData.
        CurrentPosition := StrPos(ConvertedName, '%');
        while CurrentPosition > 0 do begin
            ConvertedName := DelStr(ConvertedName, CurrentPosition, 1);
            ConvertedName := InsStr(ConvertedName, 'Percent', CurrentPosition);
            CurrentPosition := StrPos(ConvertedName, '%');
        end;

        CurrentPosition := 1;

        while CurrentPosition <= StrLen(ConvertedName) do begin
            if ConvertedName[CurrentPosition] in [' ', '\', '/', '''', '"', '.', '(', ')', '-', ':'] then
                if CurrentPosition > 1 then begin
                    if ConvertedName[CurrentPosition - 1] = '_' then begin
                        ConvertedName := DelStr(ConvertedName, CurrentPosition, 1);
                        CurrentPosition -= 1;
                    end else
                        ConvertedName[CurrentPosition] := '_';
                end else
                    ConvertedName[CurrentPosition] := '_';

            CurrentPosition += 1;
        end;

        ConvertedName := RemoveTrailingUnderscore(ConvertedName);
    end;

    local procedure RemoveTrailingUnderscore(Input: Text): Text
    begin
        Input := DelChr(Input, '>', '_');
        exit(Input);
    end;

    local procedure CrearCabecera(idInformes: Integer; var TempExcelBuffer: Record "Excel Buffer" temporary; var Row: Integer; DesdeFecha: Date; HastaFecha: Date; var FieldRef: FieldRef; TenantRecorId: RecordId; var RecRefTemp: RecordRef)
    var
        DF: DateFormula;
        Destinatario: Record "Destinatarios Informes";
        Empresas: Record "Empresas Informes";
        Informes: Record "Informes";
        TypeHelper: Codeunit "Type Helper";
        Campos: Record "Columnas Informes";
        Filtros: Record "Filtros Informes";
    begin
        Campos.Reset();
        Informes.Get(idInformes);
        Destinatario.SetRange(Id, Informes.Id);
        Filtros.SetRange(Id, Informes.Id);
        Campos.SetRange(Id, Informes.Id);
        Empresas.SetRange(Id, Informes.Id);
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
                    //'VERSION(1) SORTING(Tipo,Nº mov.) WHERE(Fecha vencimiento=FILTER(1925-04-11..2024-04-11),Cód. forma pago=FILTER(PAG. FIRMA|PAGARE))'
                    EnterCell(TempExcelBuffer, Row, 1, FieldRef.Caption, true, false, '', TempExcelBuffer."Cell Type"::Text);
                    if DesdeFecha <> 0D then
                        EnterCell(TempExcelBuffer, Row, 2, CopyStr(TypeHelper.FormatDateWithCurrentCulture(DesdeFecha), 1, 250), false, false, '', TempExcelBuffer."Cell Type"::Text);
                    EnterCell(TempExcelBuffer, Row, 3, CopyStr(TypeHelper.FormatDateWithCurrentCulture(HastaFecha), 1, 250), false, false, '', TempExcelBuffer."Cell Type"::Text);
                end else begin
                    FieldRef.SetFilter(Filtros.Valor);
                    EnterCell(TempExcelBuffer, Row, 1, FieldRef.Caption, true, false, '', TempExcelBuffer."Cell Type"::Text);
                    EnterCell(TempExcelBuffer, Row, 2, Filtros.Valor, false, false, '', TempExcelBuffer."Cell Type"::Text);
                end;
                CreaFiltroCampo(TenantRecorId, RecReftemp.Number, Filtros.Campo);
            until Filtros.Next() = 0;
        Row += 1;
        FieldRef := RecReftemp.Field(Destinatario."Campo Destinatario");
        FieldRef.SetFilter(Destinatario.Valor);

        Row += 1;

        EnterCell(TempExcelBuffer, Row, 1, FieldRef.Caption, true, false, '', TempExcelBuffer."Cell Type"::Text);
        EnterCell(TempExcelBuffer, Row, 2, Destinatario.Valor, false, false, '', TempExcelBuffer."Cell Type"::Text);
        Row += 1;
        Campos.SetRange(Id, Informes.Id);
        Campos.SetRange(Include, true);
        if Campos.FindSet() then
            repeat

                EnterCell(TempExcelBuffer, Row, Campos.Orden, Campos.Titulo, true, false, '', TempExcelBuffer."Cell Type"::Text);

            until Campos.Next() = 0;
        If Empresas.FindFirst() then
            If Empresas."Columna Excel" <> 0 then
                EnterCell(TempExcelBuffer, Row, Empresas."Columna Excel", 'Empresa', true, false, '', TempExcelBuffer."Cell Type"::Text);
    end;


    procedure BuildUrl(ServiceRootUrlParam: Text; SelectTextParam: Text; FilterTextParam: Text): Text
    var
        ODataUrl: Text;
        preSelectTextConjunction: Text;
    begin
        if StrPos(ServiceRootUrlParam, '?tenant=') > 0 then
            preSelectTextConjunction := '&'
        else
            preSelectTextConjunction := '?';

        if (StrLen(SelectTextParam) > 0) and (StrLen(FilterTextParam) > 0) then
            ODataUrl := ServiceRootUrlParam + preSelectTextConjunction + SelectTextParam + '&' + FilterTextParam
        else
            if StrLen(SelectTextParam) > 0 then
                ODataUrl := ServiceRootUrlParam + preSelectTextConjunction + SelectTextParam
            else
                if StrLen(FilterTextParam) > 0 then
                    ODataUrl := ServiceRootUrlParam + preSelectTextConjunction + FilterTextParam
                else
                    // FilterText is based on SelectText, so it doesn't make sense to have only the FilterText.
                    ODataUrl := ServiceRootUrlParam;

        exit(ODataUrl);
    end;

    procedure CreateTenantWebServiceFilterFromRecordRef(var TenantWebServiceFilter: Record "Tenant Web Service Filter"; var RecordRef: RecordRef; TenantWebServiceRecordId: RecordID)
    begin
        TenantWebServiceFilter.SetRange(TenantWebServiceID, TenantWebServiceRecordId);
        TenantWebServiceFilter.DeleteAll();

        TenantWebServiceFilter.Init();
        TenantWebServiceFilter."Entry ID" := 0;
        TenantWebServiceFilter."Data Item" := RecordRef.Number();
        TenantWebServiceFilter.TenantWebServiceID := TenantWebServiceRecordId;
        TenantWebServiceFilter.Insert();
        SetTenantWebServiceFilter(TenantWebServiceFilter, RecordRef.GetView(), TenantWebServiceRecordId, RecordRef.Number());

        TenantWebServiceFilter.CalcFields(Filter);
        If not TenantWebServiceFilter.Filter.HasValue then
            Error('No se ha podido generar el filtro para el informe');
    end;





    procedure SetTenantWebServiceFilter(var TenantWebServiceFilter: Record "Tenant Web Service Filter"; FilterText: Text; TenantWebServiceRecordId: RecordID; Id: Integer)
    var
        WriteOutStream: OutStream;
    begin
        TenantWebServiceFilter.CalcFields(Filter);
        TenantWebServiceFilter.Filter.CreateOutStream(WriteOutStream);
        //if StrPos(FilterText, 'WHERE') <> 0 then
        //  FilterText := CopyStr(FilterText, 1, StrPos(FilterText, 'WHERE') - 1) + 'WHERE (' + filtros + ')';

        WriteOutStream.WriteText(FilterText);
        TenantWebServiceFilter.Modify();
        RemoveUnselectedColumnsFromFilter(TenantWebServiceRecordId, Id, FilterText);
    end;

    procedure RemoveUnselectedColumnsFromFilter(TenantWebServiceRecordId: RecordId; DataItemNumber: Integer; DataItemView: Text): Text
    var
        TenantWebServiceColumns: Record "Tenant Web Service Columns";
        BaseRecordRef: RecordRef;
        UpdatedRecordRef: RecordRef;
        BaseFieldRef: FieldRef;
        UpdatedFieldRef: FieldRef;
    begin
        BaseRecordRef.Open(DataItemNumber);
        BaseRecordRef.SetView(DataItemView);
        UpdatedRecordRef.Open(DataItemNumber);

        TenantWebServiceColumns.SetRange(TenantWebServiceID, TenantWebServiceRecordId);
        TenantWebServiceColumns.SetRange("Data Item", DataItemNumber);
        if TenantWebServiceColumns.FindSet() then begin
            repeat
                if BaseRecordRef.FieldExist(TenantWebServiceColumns."Field Number") then begin
                    BaseFieldRef := BaseRecordRef.Field(TenantWebServiceColumns."Field Number");
                    UpdatedFieldRef := UpdatedRecordRef.Field(TenantWebServiceColumns."Field Number");
                    UpdatedFieldRef.SetFilter(BaseFieldRef.GetFilter());
                end;
            until TenantWebServiceColumns.Next() = 0;
        end else begin
            TenantWebServiceColumns.SetRange("Data Item");
            TenantWebServiceColumns.ModifyAll("Data Item", DataItemNumber);
        end;

        exit(UpdatedRecordRef.GetView());
    end;

    procedure CreaFiltroCampo(TenantWebServiceRecordId: RecordId; DataItemNumber: Integer; FiledNumber: Integer)
    var
        TenantWebServiceColumns: Record "Tenant Web Service Columns";
        BaseRecordRef: RecordRef;
        UpdatedRecordRef: RecordRef;
        BaseFieldRef: FieldRef;
        UpdatedFieldRef: FieldRef;
    begin

        BaseRecordRef.Open(DataItemNumber);
        TenantWebServiceColumns.SetRange(TenantWebServiceID, TenantWebServiceRecordId);
        TenantWebServiceColumns.SetRange("Data Item", DataItemNumber);
        TenantWebServiceColumns.SetRange("Field Number", FiledNumber);
        if Not TenantWebServiceColumns.FindSet() then begin
            TenantWebServiceColumns.Init();
            TenantWebServiceColumns.TenantWebServiceID := TenantWebServiceRecordId;
            TenantWebServiceColumns."Data Item" := DataItemNumber;
            TenantWebServiceColumns."Field Number" := FiledNumber;
            UpdatedFieldRef := BaseRecordRef.Field(TenantWebServiceColumns."Field Number");
            TenantWebServiceColumns."Field Name" := UpdatedFieldRef.Name;
            TenantWebServiceColumns."Field Caption" := UpdatedFieldRef.Caption;
            TenantWebServiceColumns.Insert();
        end;


    end;

    var

        Fecha: date;
        TextValue: Text;
        Client: HttpClient;
        RequestHeaders: HttpHeaders;
        RequestContent: HttpContent;
        ResponseMessage: HttpResponseMessage;
        RequestMessage: HttpRequestMessage;
        ResponseText: Text;
        contentHeaders: HttpHeaders;
        MEDIA_TYPE: Label 'application/json';



    /// <summary>
    /// RestApi.
    /// </summary>
    /// <param name="url">Text.</param>
    /// <param name="RequestType">Option Get,patch,put,post,delete.</param>
    /// <param name="payload">Text.</param>
    /// <returns>Return value of type Text.</returns>

    procedure RestApi(url: Text; RequestType: Option Get,patch,put,post,delete; payload: Text; User: Text; Pass: Text): Text
    var
        Ok: Boolean;
        Respuesta: Text;
    begin
        RequestHeaders := Client.DefaultRequestHeaders();
        RequestHeaders.Add('Authorization', 'Basic cGk6SWI2MzQzZHMu');
        //CreateBasicAuthHeader(User, Pass, RequestHeaders); debería devilver lo mismo


        case RequestType of
            RequestType::Get:
                If Not Client.Get(URL, ResponseMessage) Then
                    exit('Retry');

            RequestType::patch:
                begin
                    RequestContent.WriteFrom(payload);

                    RequestContent.GetHeaders(contentHeaders);
                    contentHeaders.Clear();
                    contentHeaders.Add('Content-Type', 'application/json-patch+json');

                    RequestMessage.Content := RequestContent;

                    RequestMessage.SetRequestUri(URL);
                    RequestMessage.Method := 'PATCH';

                    if not client.Send(RequestMessage, ResponseMessage) then exit('Retry');
                end;
            RequestType::post:
                begin
                    RequestContent.WriteFrom(payload);

                    RequestContent.GetHeaders(contentHeaders);
                    contentHeaders.Clear();
                    contentHeaders.Add('Content-Type', 'application/json');

                    if not Client.Post(URL, RequestContent, ResponseMessage) then exit('Retry');
                end;
            RequestType::delete:
                If not Client.Delete(URL, ResponseMessage) then
                    exit('Retry');
        end;

        ResponseMessage.Content().ReadAs(ResponseText);
        exit(ResponseText);

    end;

    procedure Semana(Fecha: Date): Integer
    var
        rFecha: Record Date;
    begin
        rFecha.RESET;
        rFecha.SETRANGE("Period Type", rFecha."Period Type"::Week);
        rFecha.SETFILTER("Period Start", '<=%1', "Fecha");
        rFecha.SETFILTER("Period End", '>=%1', "Fecha");
        if rFecha.FIND('-') THEN
            Exit(rFecha."Period No.");
    end;

    Procedure Comercial(Empresa: Text; EntryNo: Text): Text
    var
        R21: Record 21;
        r18: Record 18;
        r13: Record 13;
        a: Integer;
    begin
        if not Evaluate(a, EntryNo) then
            exit('');
        r13.ChangeCompany(Empresa);

        r21.CHANGECOMPANY(Empresa);
        if r21.GET(EntryNo) THEN BEGIN
            if r13.GET(r21."Salesperson Code") THEN EXIT(r13.Name);
        END;

    end;


}
