Codeunit 7001130 ControlInformes
{
    trigger OnRun()
    var
        informes: Record "Informes";
    begin
        imprimirInformes(0, CurrentDateTime, false);
    end;


    procedure imprimirInformes(IdInforme: Integer; ProximaFecha: Datetime; NoEnviar: Boolean)
    var
        Informe: Record "Informes";
        Destinatario: Record "Destinatarios Informes";
        Filtros: Record "Filtros Informes";
        Contratos: Page "Lista Contratos x Empresa";
        Anticipos: Page "Ingresos Anticipados";
        Conta: Page "MovContabilidad";
        Panrorama: Page "PanoramaBy";
        Saldo: Page "Saldo Interempresas Det";
        Out: OutStream;
        ficheros: Record Ficheros temporary;
        Secuencia: Integer;
        MillisecondsToAdd: Integer;
        NoOfMinutes: Integer;
        Intstream: InStream;
        HoraInicio: Boolean;
        HoraFin: Boolean;
        Base64: Text;
        Base64Convert: Codeunit "Base64 Convert";
        DocAttch: Record "Document Attachment" temporary;
        Id: Integer;
        Control: Codeunit "ControlInformes";
        UrlPlantilla: Text;
        PlantillaBase64: Text;
        FiltroDestinatario: Boolean;
        Primero: Boolean;
        UrlExcel: Text;
        RecReftemp: RecordRef;
    begin
        NoOfMinutes := 5;
        MillisecondsToAdd := NoOfMinutes;
        MillisecondsToAdd := MillisecondsToAdd * 60000;
        If IdInforme <> 0 tHEN
            Informe.SetRange("ID", IdInforme);
        If ProximaFecha <> 0DT tHEN
            Informe.SetRange("Fecha Próx. Ejecución", Today);
        Informe.SetRange(Ejecutandose, false);

        If Informe.FindSet() then begin
            repeat
                HoraInicio := (Informe."Starting Time" <= Time);
                HoraFin := ((Informe."Ending Time" = 0T) or (Informe."Ending Time" >= Time));
                If (HoraInicio and HoraFin) Or (ProximaFecha = 0DT) then begin
                    If (Enhora(Informe."Earliest Start Date/Time", CurrentDateTime) or (ProximaFecha = 0DT)) then begin

                        // If Informe."Url Plantilla" = '' Then begin
                        //     Informe.CalcFields("Plantilla Excel");
                        //     If Informe."Plantilla Excel".HasValue then
                        //         Intstream := Control.UrlPlantillaInstream(UrlPlantilla, Informe, PlantillaBase64, true);
                        // end;
                        Informe.Ejecutandose := true;
                        Informe."Fecha Últ. Ejecución" := Today;
                        Informe.Modify();
                        Commit();
                        Destinatario.Reset;
                        Destinatario.SetRange("No enviar", false);
                        Destinatario.SetRange("ID", Informe."ID");
                        FiltroDestinatario := CompruebaFiltros(Destinatario);
                        Primero := true;
                        if Destinatario.FindSet() then begin
                            repeat
                                if Primero or FiltroDestinatario Then begin
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
                                                Contratos.ExportExcel(Filtros, Informe.Id, Destinatario, out);
                                            end;
                                        Informes::"Ingresos Anticipados":
                                            begin
                                                Clear(Anticipos);
                                                Anticipos.ExportExcel(Filtros, Informe.Id, Destinatario, out);
                                            end;
                                        informes::"Estadisticas Contabilidad":
                                            begin
                                                Clear(Conta);
                                                Conta.ExportExcel(Filtros, Informe.Id, Destinatario, out, RecReftemp, Primero);
                                            end;
                                        Informes::Tablas:
                                            begin
                                                ExportExcel(Filtros, Informe.Id, Destinatario, out);
                                            end;
                                        Informes::"Web Service":
                                            begin

                                                if ExportExcelWeb(Filtros, Informe.Id, Destinatario, out) = 'Retry' then begin
                                                    //RecReftemp.Close();
                                                    ExportExcelWeb(Filtros, Informe.Id, Destinatario, out);
                                                end;

                                            end;
                                        Informes::"Informes Financieros":
                                            begin
                                                Clear(Panrorama);
                                                Panrorama.ExportExcel(Filtros, Informe.Id, Destinatario, out);

                                            end;
                                        Informes::"Saldo InterEmpresas":
                                            begin
                                                Clear(Saldo);
                                                Saldo.ExportExcel(Filtros, Informe.Id, Destinatario, out);
                                            end;

                                    end;
                                    ficheros.Modify();
                                    Commit();
                                end;
                                if NoEnviar then begin
                                    ficheros.CalcFields(Fichero);
                                    ficheros.Fichero.CreateInStream(Intstream);
                                    DownloadFromStream(Intstream, 'Guardar', 'C:\Temp', 'ALL Files (*.*)|*.*', ficheros."Nombre fichero");
                                    Destinatario.FindLast();
                                end else begin
                                    If Primero or FiltroDestinatario Then begin
                                        ficheros.CalcFields(Fichero);
                                        ficheros.Fichero.CreateInStream(Intstream);
                                        Base64 := Base64Convert.ToBase64(Intstream);
                                        ficheros.Delete();
                                        UrlExcel := DocAttch.FormBase64ToUrl(Base64, Format(Informe.Id) + '.xlsx', Id);
                                    end;
                                    EnviaCorreo(Destinatario."e-mail", ficheros, Informe.Descripcion, destinatario."Nombre Informe", Informe."ID", Destinatario."Nombre Empleado", UrlExcel);
                                end;
                                Primero := false;
                            until Destinatario.Next() = 0;
                        end;
                        IF ProximaFecha <> 0DT tHEN begin
                            ProximaFecha := CurrentDateTime + MillisecondsToAdd;

                            Informe."Earliest Start Date/Time" := Informe.CalcNextRunTimeForRecurringReport(Informe, ProximaFecha);
                            If (Informe."Ending Time" = 0T) Or (Time > Informe."Ending Time") then
                                Informe."Fecha Próx. Ejecución" := DT2Date(Informe."Earliest Start Date/Time");

                        end;
                        Informe.Ejecutandose := false;
                        Informe.Modify;
                    end;
                end;
            until Informe.Next() = 0;
        end;
    end;

    procedure StrReplace(String: Text[250]; FindWhat: Text[250]; ReplaceWith: Text[250]) NewString: Text[250]
    begin
        WHILE STRPOS(String, FindWhat) > 0 DO
            String := DELSTR(String, STRPOS(String, FindWhat)) + ReplaceWith + COPYSTR(String, STRPOS(String, FindWhat) + STRLEN(FindWhat));
        NewString := String;
    end;

    local procedure EnviaCorreo(SalesPersonMail: Text; var ficheros: Record Ficheros; Informe: Text; InformeDestinatario: Text; IdInforme: Integer; NomBreDestinatario: Text; Url: Text)


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
        Saludo: Text;
        Base64: Text;
    begin
        rInf.Get();
        workdescription := Informes.GetDescripcionAmpliada();
        if InformeDestinatario <> '' then
            Informe := ConvertStr(InformeDestinatario, ' ', '_') + '_' + Format(Today(), 0, '<Year4><Month,2><Day,2>');
        if Time < 140000T then
            Saludo := 'Buenos días'
        else if Time < 200000T then
            Saludo := 'Buenas tardes'
        else
            Saludo := 'Buenas noches';
        if workdescription <> '' then begin
            //#Saludo #Nombreempleado, Ajduntamos el siguiente informe: #Nombre informe
            If StrPos(workdescription, '#Nombreempleado') > 0 then
                workdescription := StrReplace(workdescription, '#Nombreempleado', NomBreDestinatario)
            else begin
                BigText := ('Estimado:');
                BigText := BigText + '<br> </br>';
                BigText := BigText + '<br> </br>';
            end;
            If StrPos(workdescription, '#Nombreinforme') > 0 then
                workdescription := StrReplace(workdescription, '#Nombreinforme', InformeDestinatario);
            if StrPos(workdescription, '#Saludo') > 0 then
                workdescription := StrReplace(workdescription, '#Saludo', Saludo);
        end else begin
            BigText := ('Estimado:');
            BigText := BigText + '<br> </br>';
            BigText := BigText + '<br> </br>';
        end;


        //(FORMAT(cr,0,'<CHAR>') + FORMAT(lf,0,'<CHAR>')

        Informes.get(IdInforme);


        if workdescription = '' then
            BigText := BigText + 'Remitimos el siguiente Informe: ' + Informe;


        BigText := BigText + '<br> </br>';
        BigText := BigText + workdescription;
        If Url <> '' Then
            BigText := BigText + '<font Color="blue"><a href="' + url + '">' + 'Pulse el Enlace para acceder al informe' + '</a></font></td>';

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
        REmail.AddAttachment(Funciones.CargaPie(Base64), 'emailfoot.png');
        BigText := BigText + '<img src="data:image/png;base64,' + base64 + '" />';//"emailFoot.png" />';
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
                REmail."Send BCC" := 'andreuserra@malla.es;lllompart@malla.es';
        end;
        REmail.SetBodyText(BigText);
        REmail."From Name" := UserId;
        REmail.Subject := Informe;


        // ficheros.CalcFields(Fichero);
        // ficheros.Fichero.CreateInStream(AttachmentStream);
        // REmail.AddAttachment(AttachmentStream, Informe + '.xlsx');

        // if REmail."From Address" <> '' Then
        //     REmail."Send BCC" := REmail."From Address" else
        //     REmail."Send BCC" := BCC();
        REmail.Send(true, emilesc::Informes);
        //ficheros.delete;

    end;

    procedure ExportExcel(var Filtros: Record "Filtros Informes";
    IdInforme: Integer;
    Var Destinatario: Record "Destinatarios Informes"; var ExcelStream: OutStream)
    var
        TempExcelBuffer: Record "Excel Buffer 2" temporary;
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
        PlantillaBase64: Text;
        Continuar: Boolean;
        FechaTarea: Date;
        Campos: Record "Columnas Informes";
        Formatos: Record "Formato Columnas";
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
        LinkCliente: Record Customer;
        Vinculo: Text;
        TempBlob: Codeunit "Temp Blob";
        Base64Convert: Codeunit "Base64 Convert";
        UrlPlantilla: Text;
    begin

        TempExcelBuffer.Reset();
        TempExcelBuffer.DeleteAll();
        FechaTarea := CalcDate('1S', WorkDate());
        Informes.Get(IdInforme);
        Campos.SetRange(Id, Informes.Id);
        Campos.SetFilter(Table, '<>%1', 0);
        Campos.FindFirst();
        RecReftemp.Open(Campos.Table);
        Campos.Reset();
        Row := 1;
        EnterCell(TempExcelBuffer, Row, 1, StrSubstNo('%1 de %2', Informes.Descripcion, DT2Date(Informes."Earliest Start Date/Time")),
        true, false, false, false, '', TempExcelBuffer."Cell Type"::Text, '', 0, '', '', '');
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

        EnterCell(TempExcelBuffer, Row, 1, FieldRef.Caption, true, false, false, false, '', TempExcelBuffer."Cell Type"::Text, '', 0, '', '', '');
        EnterCell(TempExcelBuffer, Row, 2, Destinatario.Valor, false, false, false, false, '', TempExcelBuffer."Cell Type"::Text, '', 0, '', '', '');
        Row += 1;
        Campos.SetRange(Id, Informes.Id);
        Campos.SetRange(Include, true);
        if Campos.FindSet() then
            repeat
                If Not Formatos.Get(campos.Id, campos.Id_campo, true) then begin
                    Formatos.Init();
                    Formatos.Bold := true;
                end;

                EnterCell(TempExcelBuffer, Row, Campos.Orden, Campos.Titulo, Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline"
                , '', TempExcelBuffer."Cell Type"::Text, formatos.Fuente, Formatos."Tamaño", Formatos."Color Fuente", Formatos."Color Fondo", '');
                if Campos."Ancho Columna" <> 0 then
                    TempExcelBuffer.SetColumnWidth(Campos.LetraColumna(Campos.Orden), Campos."Ancho Columna");
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

                        end;
                        If Not Formatos.Get(campos.Id, campos.Id_campo, false) then begin
                            Formatos.Init();
                            if FieldT = FieldType::Decimal then
                                Formatos."Formato Columna" := '_-* #,##0.00_-;-* #,##0.00_-;_-* "-"_-;_-@_-';
                        end;
                        Vinculo := '';
                        If Formatos."Insertar Vínculo" then begin
                            If RecrefTemp.field(Campos.Campo).Relation = 18 then begin
                                if Not LinkCliente.Get(RecrefTemp.Field(Campos.Campo).Value) then
                                    LinkCliente.Init()
                                else
                                    LinkCliente.SetRange("No.", Linkcliente."No.");
                                Vinculo := GetUrl(ClientType::Web, CompanyName, ObjectType::Page, Page::"Customer Card", LinkCliente);

                            end;
                        end;
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
                                EnterCell(TempExcelBuffer, Row, Campos.Orden, Valor, Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Number, Formatos.Fuente, Formatos."Tamaño", Formatos."Color Fuente", Formatos."Color Fondo", Vinculo);
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
                                EnterCell(TempExcelBuffer, Row, Campos.Orden, Valor, Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Text, Formatos.Fuente, Formatos."Tamaño", Formatos."Color Fuente", Formatos."Color Fondo", Vinculo);
                            FieldT::RecordId:
                                EnterCell(TempExcelBuffer, Row, Campos.Orden, Valor, Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Text, Formatos.Fuente, Formatos."Tamaño", Formatos."Color Fuente", Formatos."Color Fondo", Vinculo);
                            FieldT::Blob:
                                EnterCell(TempExcelBuffer, Row, Campos.Orden, Valor, Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Text, Formatos.Fuente, Formatos."Tamaño", Formatos."Color Fuente", Formatos."Color Fondo", Vinculo);
                            FieldT::Guid:
                                EnterCell(TempExcelBuffer, Row, Campos.Orden, Valor, Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Text, Formatos.Fuente, Formatos."Tamaño", Formatos."Color Fuente", Formatos."Color Fondo", Vinculo);

                        End;


                    until Campos.Next() = 0;

            //end;
            until RecrefTemp.Next() = 0;
        RecrefTemp.Close();
        Informes.CalcFields("Plantilla Excel");
        ExcelFileNameEPR := ConvertStr(Informes.Descripcion, ' ', '_');
        if Destinatario."Nombre Informe" <> '' then
            ExcelFileNameEPR := ConvertStr(Destinatario."Nombre Informe", ' ', '_');
        if (Informes."Plantilla Excel".HasValue) Or (Informes."Url Plantilla" <> '') then begin
            if Informes."Plantilla Excel".HasValue then
                //Informes."Plantilla Excel".CreateInStream(InExcelStream);
                Control.UrlPlantillaInstream(UrlPlantilla, Informes, PlantillaBase64, false);
            if Not Informes."Formato Json" then
                TempExcelBuffer.UpdateBookStream(PlantillaBase64, ConvertStr(Informes.Descripcion, ' ', '_'), true);

        end else begin
            if Informes."Formato Json" then
                PlantillaBase64 := '' else
                TempExcelBuffer.CreateNewBook(ExcelFileNameEPR);
        end;
        If Not Informes."Formato Json" then begin
            TempExcelBuffer.WriteSheet(ConvertStr(Informes.Descripcion, ' ', '_'), CompanyName, UserId, Informes."Orientación");
            TempExcelBuffer.CloseBook();
            TempExcelBuffer.SetFriendlyFilename(StrSubstNo(ExcelFileNameEPR, ConvertStr(Informes.Descripcion, ' ', '_'), CurrentDateTime, UserId));
            TempExcelBuffer.SaveToStream(ExcelStream, true);
        end else begin
            TempExcelBuffer.SetRange("Sheet Name", '');
            TempExcelBuffer.ModifyAll("Sheet Name", ConvertStr(Informes.Descripcion, ' ', '_'));
            TempExcelBuffer.Reset();
            PlantillaBase64 := JsonExcel(TempExcelBuffer, PlantillaBase64, gUrlPlantilla);
            Base64Convert.FromBase64(PlantillaBase64, ExcelStream);
        end;
    end;

    procedure ExportExcelWeb(var Filtros: Record "Filtros Informes";
    IdInforme: Integer;
    Var Destinatario: Record "Destinatarios Informes"; var ExcelStream: OutStream): text
    var
        Enlaces: Record "Enlaces Informes";
        TempExcelBuffer: Record "Excel Buffer 2" temporary;
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
        LinkCliente: Record Customer;
        Vinculo: Text;
        LinkTable: RecordRef;
        LinkKampo: FieldRef;
        TableNo: Integer;
        PageID: Integer;
        PageManagement: Codeunit "Page Management";
        Vendor: Record Vendor;
        RecRefEnlace: RecordRef;
        FieldRefEnlace: FieldRef;
        TempBlob: Codeunit "Temp Blob";
        Base64Convert: Codeunit "Base64 Convert";
        PlantillaBase64: Text;
    begin

        TempExcelBuffer.Reset();
        TempExcelBuffer.DeleteAll();
        FechaTarea := CalcDate('1S', WorkDate());
        Informes.Get(IdInforme);
        Campos.SetRange(Id, Informes.Id);

        Empresas.SetRange(Id, Informes.Id);
        Campos.SetFilter(Table, '<>%1', 0);

        Campos.FindFirst();
        RecReftemp.Close();
        RecReftemp.Open(Campos.Table);
        Campos.SetRange(Table);
        Campos.ModifyAll(Table, RecReftemp.Number);
        Empresas.SetRange(Incluir, true);
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
        Empresas.SetRange(Incluir, true);
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
        Empresas.SetRange(Incluir, true);
        If not Empresas.FindSet() then begin
            Empresas.Init();
            Empresas.Empresa := CompanyName;
        end else begin
            HojasSeparadas := Empresas."Hojas Separadas";
            if HojasSeparadas then begin

                if (Informes."Plantilla Excel".HasValue) Or (Informes."Url Plantilla" <> '') then begin
                    //if Informes."Plantilla Excel".HasValue then
                    //  Informes."Plantilla Excel".CreateInStream(InExcelStream);
                    Control.UrlPlantillaInstream(gUrlPlantilla, Informes, PlantillaBase64, false);
                    if Not Informes."Formato Json" then
                        TempExcelBuffer.UpdateBookStream(PlantillaBase64, ConvertStr(Empresas.HojaExcel, ' ', '_'), true);

                end else begin
                    if Informes."Formato Json" then
                        PlantillaBase64 := '' else
                        TempExcelBuffer.CreateNewBook(ConvertStr(Empresas.HojaExcel, ' ', '_'));
                end;
                if Informes."Formato Json" then begin
                    TempExcelBuffer.SetRange("Sheet Name", '');
                    TempExcelBuffer.ModifyAll("Sheet Name", ConvertStr(Empresas.HojaExcel, ' ', '_'));
                end;


            end;
        end;
        repeat

            if HojasSeparadas then begin
                TempExcelBuffer.SelectOrAddSheet(ConvertStr(Empresas.HojaExcel, ' ', '_'));
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
                Jsontext := RestApi.RestApi(url, ResuestType::GET, '', 'debug', 'Ib6343ds.');
                if Jsontext = 'Retry' then begin
                    exit('Retry');
                    Clear(RestApi);
                    Jsontext := RestApi.RestApi(url, ResuestType::GET, '', 'debug', 'Ib6343ds.');
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
                        EnterCell(TempExcelBuffer, Row, Empresas."Columna Excel", Empresas.Empresa, false, false, false, false, '', TempExcelBuffer."Cell Type"::Text, '', 0, '', '', '');
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
                                        Campos.Funcion::Cliente:
                                            begin
                                                FieldT := FieldType::Text;
                                                TextValue := DevuelveCampo(Campos."Field Name", JsonObj, Fieldt);
                                                cUST.ChangeCompany(Empresas.Empresa);
                                                iF cUST.Get(TextValue) then
                                                    TextValue := cUST.Name;
                                            end;
                                        Campos.Funcion::Proveedor:
                                            begin
                                                FieldT := FieldType::Text;
                                                TextValue := DevuelveCampo(Campos."Field Name", JsonObj, Fieldt);
                                                Vendor.ChangeCompany(Empresas.Empresa);
                                                iF Vendor.Get(TextValue) then
                                                    TextValue := Vendor.Name;
                                            end;
                                        Campos.Funcion::Cadena:
                                            begin
                                                FieldT := FieldType::Text;
                                                TextValue := DevuelveCampo(Campos."Field Name", JsonObj, Fieldt);
                                                cUST.ChangeCompany(Empresas.Empresa);
                                                iF cUST.Get(TextValue) then
                                                    TextValue := cUST."Cod cadena";
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
                                        Campos.Funcion::Enlace:
                                            begin
                                                FieldT := FieldType::Text;
                                                TextValue := DevuelveCampo(Campos."Field Name", JsonObj, Fieldt);
                                                Enlaces.Get(Informes.Id, Campos.Id_campo);
                                                RecRefEnlace.Close();
                                                RecRefEnlace.Open(Enlaces."Tabla");
                                                RecRefEnlace.ChangeCompany(Empresas.Empresa);
                                                FieldRefEnlace := RecRefEnlace.Field(Enlaces."Campo Enlace");
                                                FieldRefEnlace.SetFilter(TextValue);
                                                iF RecRefEnlace.FindFirst() then begin
                                                    FieldRefEnlace := RecRefEnlace.Field(Enlaces."Campo Datos");
                                                    TextValue := FieldRefEnlace.Value;
                                                end else
                                                    TextValue := '';
                                            end;
                                        Campos.Funcion::HiperVinculo:
                                            begin
                                                If Not Formatos.Get(campos.Id, campos.Id_campo, false) then
                                                    Formatos.Init();
                                                TableNo := Formatos."Tabla Hipervínculo";
                                                If TableNo = 0 Then TableNo := RecReftemp.Field(Campos.Campo).Relation;
                                                If TableNo <> 0 Then begin
                                                    LinkTable.Open(TableNo);
                                                    if Empresas.Empresa <> '' Then
                                                        LinkTable.ChangeCompany(Empresas.Empresa);
                                                    LinkKampo := LinkTable.Field(Formatos."Campo Hipervínculo");
                                                    If Formatos."Campo Relación" = 0 Then
                                                        LinkKampo.SetRange(DevuelveCampoSinFormato(Campos."Field Name", JsonObj, Fieldt))
                                                    else
                                                        LinkKampo.SetRange(DevuelveCampoSinFormato(Formatos."Nombre Campo Relación", JsonObj, Fieldt));
                                                    PageID := PageManagement.GetPageID(LinkTable);
                                                    if Empresas.Empresa <> '' then
                                                        Vinculo := GetUrl(ClientType::Web, Empresas.Empresa, ObjectType::Page, PageID, LinkTable, true)
                                                    else
                                                        Vinculo := GetUrl(ClientType::Web, CompanyName, ObjectType::Page, PageID, LinkTable, true);
                                                    LinkTable.Close();
                                                    If StrPos(Vinculo, 'http://NAV-MALLA01:48900') <> 0 then
                                                        Vinculo := 'https://bc220.malla.es/' + CopyStr(Vinculo, 26);
                                                    If StrPos(Vinculo, '/POWERBI/') <> 0 then
                                                        Vinculo := CopyStr(Vinculo, 1, StrPos(Vinculo, '/POWERBI/') - 1) + '/BC220/' + CopyStr(Vinculo, StrPos(Vinculo, '/POWERBI/') + 8);
                                                    TextValue := Vinculo;
                                                    Vinculo := '';
                                                end;
                                            end;
                                    End;

                                end;
                                FieldT := FieldType::Text;

                            end;
                            If Not Formatos.Get(campos.Id, campos.Id_campo, false) then begin
                                Formatos.Init();
                                if FieldT = FieldType::Decimal then
                                    Formatos."Formato Columna" := '_-* #,##0.00_-;-* #,##0.00_-;_-* "-"_-;_-@_-';
                            end;
                            Vinculo := '';
                            If Formatos."Insertar Vínculo" then begin
                                TableNo := Formatos."Tabla Hipervínculo";
                                If TableNo = 0 Then TableNo := RecReftemp.Field(Campos.Campo).Relation;
                                If TableNo <> 0 Then begin
                                    LinkTable.Open(TableNo);
                                    if Empresas.Empresa <> '' Then
                                        LinkTable.ChangeCompany(Empresas.Empresa);
                                    LinkKampo := LinkTable.Field(Formatos."Campo Hipervínculo");
                                    If Formatos."Campo Relación" = 0 Then
                                        LinkKampo.SetRange(DevuelveCampoSinFormato(Campos."Field Name", JsonObj, Fieldt))
                                    else
                                        LinkKampo.SetRange(DevuelveCampoSinFormato(Formatos."Nombre Campo Relación", JsonObj, Fieldt));
                                    PageID := PageManagement.GetPageID(LinkTable);
                                    if Empresas.Empresa <> '' then
                                        Vinculo := GetUrl(ClientType::Web, Empresas.Empresa, ObjectType::Page, PageID, LinkTable, true)
                                    else
                                        Vinculo := GetUrl(ClientType::Web, CompanyName, ObjectType::Page, PageID, LinkTable, true);
                                    LinkTable.Close();
                                end;
                                // If (RecrefTemp.field(Campos.Campo).Name in ['Account No.', 'Source No.']) Or (RecReftemp.Field(Campos.Campo).Relation = 18)
                                //  then begin
                                //     if Empresas.Empresa <> '' Then LinkCliente.ChangeCompany(Empresas.Empresa) else LinkCliente.ChangeCompany(CompanyName);
                                //     if Not LinkCliente.Get(DevuelveCampo(Campos."Field Name", JsonObj, Fieldt)) then
                                //         LinkCliente.Init();
                                //     if Empresas.Empresa <> '' then
                                //         Vinculo := GetUrl(ClientType::Web, Empresas.Empresa, ObjectType::Page, Page::"Customer Card", LinkCliente)
                                //     else
                                //         Vinculo := GetUrl(ClientType::Web, CompanyName, ObjectType::Page, Page::"Customer Card", LinkCliente);
                                // end;
                            End;
                            Case FieldT of
                                FieldT::Date:
                                    begin

                                        //EnterCell(TempExcelBuffer, Row, Campos.Orden, Valor, Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Text, Formatos.Fuente, Formatos."Tamaño", Formatos."Color Fuente", Formatos."Color Fondo");
                                        iF Fecha <> 0D then
                                            EnterCell(TempExcelBuffer, Row, Campos.Orden, CopyStr(TypeHelper.FormatDateWithCurrentCulture(Fecha), 1, 250), Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Date, Formatos.Fuente, Formatos."Tamaño", Formatos."Color Fuente", Formatos."Color Fondo", Vinculo)
                                        else
                                            EnterCell(TempExcelBuffer, Row, Campos.Orden, '', Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Text, Formatos.Fuente, Formatos."Tamaño", Formatos."Color Fuente", Formatos."Color Fondo", Vinculo);
                                    END;
                                FieldT::Time:
                                    EnterCell(TempExcelBuffer, Row, Campos.Orden, TextValue, Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Time, Formatos.Fuente, Formatos."Tamaño", Formatos."Color Fuente", Formatos."Color Fondo", Vinculo);
                                FieldT::Integer:
                                    EnterCell(TempExcelBuffer, Row, Campos.Orden, TextValue, Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Number, Formatos.Fuente, Formatos."Tamaño", Formatos."Color Fuente", Formatos."Color Fondo", Vinculo);
                                FieldT::Decimal:
                                    EnterCell(TempExcelBuffer, Row, Campos.Orden, TextValue, Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Number, Formatos.Fuente, Formatos."Tamaño", Formatos."Color Fuente", Formatos."Color Fondo", Vinculo);
                                FieldT::Option:
                                    EnterCell(TempExcelBuffer, Row, Campos.Orden, TextValue, Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Text, Formatos.Fuente, Formatos."Tamaño", Formatos."Color Fuente", Formatos."Color Fondo", Vinculo);
                                FieldT::Code:
                                    EnterCell(TempExcelBuffer, Row, Campos.Orden, TextValue, Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Text, Formatos.Fuente, Formatos."Tamaño", Formatos."Color Fuente", Formatos."Color Fondo", Vinculo);
                                FieldT::Text:
                                    EnterCell(TempExcelBuffer, Row, Campos.Orden, TextValue, Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Text, Formatos.Fuente, Formatos."Tamaño", Formatos."Color Fuente", Formatos."Color Fondo", Vinculo);
                                FieldT::Boolean:
                                    EnterCell(TempExcelBuffer, Row, Campos.Orden, TextValue, Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Text, Formatos.Fuente, Formatos."Tamaño", Formatos."Color Fuente", Formatos."Color Fondo", Vinculo);
                                FieldT::RecordId:
                                    EnterCell(TempExcelBuffer, Row, Campos.Orden, TextValue, Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Text, Formatos.Fuente, Formatos."Tamaño", Formatos."Color Fuente", Formatos."Color Fondo", Vinculo);
                                FieldT::Blob:
                                    EnterCell(TempExcelBuffer, Row, Campos.Orden, TextValue, Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Text, Formatos.Fuente, Formatos."Tamaño", Formatos."Color Fuente", Formatos."Color Fondo", Vinculo);
                                FieldT::Guid:
                                    EnterCell(TempExcelBuffer, Row, Campos.Orden, TextValue, Formatos.Bold, Formatos.Italic, Formatos.Underline, Formatos."Double Underline", Formatos."Formato Columna", TempExcelBuffer."Cell Type"::Text, Formatos.Fuente, Formatos."Tamaño", Formatos."Color Fuente", Formatos."Color Fondo", Vinculo);

                            End;


                        until Campos.Next() = 0;

                end;
            until Periodos.Next() = 0;
            if HojasSeparadas then begin
                //TempExcelBuffer.SelectOrAddSheet(ConvertStr(Empresas.HojaExcel, ' ', '_'));
                if Informes."Formato Json" then begin
                    TempExcelBuffer.SetRange("Sheet Name", '');
                    TempExcelBuffer.ModifyAll("Sheet Name", ConvertStr(Empresas.HojaExcel, ' ', '_'));
                end else begin
                    TempExcelBuffer.WriteSheet(ConvertStr(Empresas.HojaExcel, ' ', '_'), Empresas.Empresa, UserId, Informes."Orientación");
                    TempExcelBuffer.DeleteAll();
                end;
            end;

        until Empresas.Next() = 0;
        Informes.CalcFields("Plantilla Excel");
        if (Informes."Plantilla Excel".HasValue) Or (Informes."Url Plantilla" <> '') then begin
            Informes."Plantilla Excel".CreateInStream(InExcelStream);
            if not HojasSeparadas then begin
                //if Informes."Plantilla Excel".HasValue then
                //  Informes."Plantilla Excel".CreateInStream(InExcelStream);
                Control.UrlPlantillaInstream(gUrlPlantilla, Informes, PlantillaBase64, false);
                if Not Informes."Formato Json" then
                    TempExcelBuffer.UpdateBookStream(PlantillaBase64, ConvertStr(Informes.Descripcion, ' ', '_'), true);
            end;
        end else
            if not HojasSeparadas then begin
                if Informes."Formato Json" then
                    PlantillaBase64 := '' else
                    TempExcelBuffer.CreateNewBook(ConvertStr(Informes.Descripcion, ' ', '_'));
            end;
        if not HojasSeparadas then begin
            if Informes."Formato Json" then begin
                TempExcelBuffer.SetRange("Sheet Name", '');
                TempExcelBuffer.ModifyAll("Sheet Name", ConvertStr(Informes.Descripcion, ' ', '_'));
            end else
                TempExcelBuffer.WriteSheet(ConvertStr(Informes.Descripcion, ' ', '_'), CompanyName, UserId, Informes."Orientación");
        end;
        if Not Informes."Formato Json" then begin
            TempExcelBuffer.CloseBook();
            TempExcelBuffer.SetFriendlyFilename(StrSubstNo(ExcelFileNameEPR, ConvertStr(Informes.Descripcion, ' ', '_'), CurrentDateTime, UserId));
            TempExcelBuffer.SaveToStream(ExcelStream, true);
            exit('OK');
        end else begin
            TempExcelBuffer.Reset();
            PlantillaBase64 := JsonExcel(TempExcelBuffer, PlantillaBase64,
            gUrlPlantilla);
            Base64Convert.FromBase64(PlantillaBase64, ExcelStream);
            exit('OK');
        end;

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

    Procedure DevuelveCampoSinFormato(Campo: Text; JsonObj: JsonObject; Tipo: FieldType) ValorText: Variant
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
                        exit(Valor);
                    end;

                FieldType::Time:
                    begin
                        Valor := MyFieldRef.AsValue().AsTime();
                        exit(Valor);
                    end;
                FieldType::Integer:
                    begin
                        Valor := MyFieldRef.AsValue().AsInteger();
                        exit(Valor);
                    end;
                FieldType::Decimal:
                    begin
                        Valor := MyFieldRef.AsValue().AsDecimal();
                        exit(Valor);
                    end;
                FieldType::Option:
                    begin
                        Valor := MyFieldRef.AsValue().AsText();
                        exit(Valor);
                    end;
                FieldType::Code:
                    begin
                        Valor := MyFieldRef.AsValue().AsText();
                        exit(Valor);
                    end;
                FieldType::Text:
                    begin
                        Valor := MyFieldRef.AsValue().AsText();
                        exit(Valor);
                    end;
                FieldType::Boolean:
                    begin
                        Valor := MyFieldRef.AsValue().AsBoolean();
                        exit(Valor);
                    end;
                FieldType::RecordId:
                    begin
                        Valor := MyFieldRef.AsValue().AsText();
                        exit(Valor);
                    end;
                FieldType::Blob:
                    begin
                        Valor := MyFieldRef.AsValue().AsText();
                        exit(Valor);
                    end;
                FieldType::Guid:
                    begin
                        Valor := MyFieldRef.AsValue().AsText();
                        exit(Valor);
                    end;

            end;
        end;


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
        a: Integer;
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
        if StrPos(ConvertedName, 'º') <> 0 then begin
            a := StrPos(ConvertedName, 'º');
            ConvertedName := CopyStr(ConvertedName, 1, a - 1) + '_x00BA_' + CopyStr(ConvertedName, a + 1);
        end;
        //ª
        if StrPos(ConvertedName, 'ª') <> 0 then begin
            a := StrPos(ConvertedName, 'ª');
            ConvertedName := CopyStr(ConvertedName, 1, a - 1) + '_x00AA_' + CopyStr(ConvertedName, a + 1);
        end;

        //Ñ
        if StrPos(ConvertedName, 'Ñ') <> 0 then begin
            a := StrPos(ConvertedName, 'Ñ');
            ConvertedName := CopyStr(ConvertedName, 1, a - 1) + '_x00D1_' + CopyStr(ConvertedName, a + 1);
        end;
        //ç
        if StrPos(ConvertedName, 'ç') <> 0 then begin
            a := StrPos(ConvertedName, 'ç');
            ConvertedName := CopyStr(ConvertedName, 1, a - 1) + '_x00E7_' + CopyStr(ConvertedName, a + 1);
        end;
        //Ç
        if StrPos(ConvertedName, 'Ç') <> 0 then begin
            a := StrPos(ConvertedName, 'Ç');
            ConvertedName := CopyStr(ConvertedName, 1, a - 1) + '_x00C7_' + CopyStr(ConvertedName, a + 1);
        end;
        ConvertedName := RemoveTrailingUnderscore(ConvertedName);
    end;

    local procedure RemoveTrailingUnderscore(Input: Text): Text
    begin
        Input := DelChr(Input, '>', '_');
        exit(Input);
    end;

    local procedure CrearCabecera(idInformes: Integer; var TempExcelBuffer: Record "Excel Buffer 2" temporary; var Row: Integer; DesdeFecha: Date; HastaFecha: Date; var FieldRef: FieldRef; TenantRecorId: RecordId; var RecRefTemp: RecordRef)
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
                    //'VERSION(1) SORTING(Tipo,Nº mov.) WHERE(Fecha vencimiento=FILTER(1925-04-11..2024-04-11),Cód. forma pago=FILTER(PAG. FIRMA|PAGARE))'
                    EnterCell(TempExcelBuffer, Row, 1, FieldRef.Caption, true, false, false, false, '', TempExcelBuffer."Cell Type"::Text, '', 0, '', '', '');
                    if DesdeFecha <> 0D then
                        EnterCell(TempExcelBuffer, Row, 2, CopyStr(TypeHelper.FormatDateWithCurrentCulture(DesdeFecha), 1, 250), false, false, false, false, '', TempExcelBuffer."Cell Type"::Text, '', 0, '', '', '');
                    EnterCell(TempExcelBuffer, Row, 3, CopyStr(TypeHelper.FormatDateWithCurrentCulture(HastaFecha), 1, 250), false, false, false, false, '', TempExcelBuffer."Cell Type"::Text, '', 0, '', '', '');
                end else begin
                    FieldRef.SetFilter(Filtros.Valor);
                    EnterCell(TempExcelBuffer, Row, 1, FieldRef.Caption, true, false, false, false, '', TempExcelBuffer."Cell Type"::Text, '', 0, '', '', '');
                    EnterCell(TempExcelBuffer, Row, 2, Filtros.Valor, false, false, false, false, '', TempExcelBuffer."Cell Type"::Text, '', 0, '', '', '');
                end;
                CreaFiltroCampo(TenantRecorId, RecReftemp.Number, Filtros.Campo);
            until Filtros.Next() = 0;
        Row += 1;
        FieldRef := RecReftemp.Field(Destinatario."Campo Destinatario");
        FieldRef.SetFilter(Destinatario.Valor);

        Row += 1;

        EnterCell(TempExcelBuffer, Row, 1, FieldRef.Caption, true, false, false, false, '', TempExcelBuffer."Cell Type"::Text, '', 0, '', '', '');
        EnterCell(TempExcelBuffer, Row, 2, Destinatario.Valor, false, false, false, false, '', TempExcelBuffer."Cell Type"::Text, '', 0, '', '', '');
        Row += 1;
        Campos.SetRange(Id, Informes.Id);
        Campos.SetRange(Include, true);
        if Campos.FindSet() then
            repeat

                EnterCell(TempExcelBuffer, Row, Campos.Orden, Campos.Titulo, true, false, false, false, '', TempExcelBuffer."Cell Type"::Text, '', 0, '', '', '');
                if Campos."Ancho Columna" <> 0 then
                    TempExcelBuffer.SetColumnWidth(Campos.LetraColumna(Campos.Orden), Campos."Ancho Columna");
            until Campos.Next() = 0;
        Empresas.SetRange(Incluir, true);
        If Empresas.FindFirst() then
            If Empresas."Columna Excel" <> 0 then
                EnterCell(TempExcelBuffer, Row, Empresas."Columna Excel", 'Empresa', true, false, false, false, '', TempExcelBuffer."Cell Type"::Text, '', 0, '', '', '');
    end;

    local procedure Enhora(EarliestStartDateTime: DateTime; CurrentDateTime: DateTime): Boolean
    begin
        If EarliestStartDateTime = 0DT then
            exit(true);
        //Permite un margen de 5 minutos de diferencia
        if Abs(CurrentDateTime - EarliestStartDateTime) > 300000 then
            exit(false)
        else
            exit(true);
    end;

    local procedure CompruebaFiltros(var Destinatario: Record "Destinatarios Informes"): Boolean
    begin
        If Destinatario.FindSet() then
            repeat
                if Destinatario."Campo Destinatario" <> 0 then
                    exit(true);
            until Destinatario.Next() = 0;
        exit(false);
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
        gUrlPlantilla: Text;
        TextValue: Text;
        Client: HttpClient;
        RequestHeaders: HttpHeaders;
        RequestContent: HttpContent;
        ResponseMessage: HttpResponseMessage;
        RequestMessage: HttpRequestMessage;
        ResponseText: Text;
        contentHeaders: HttpHeaders;
        MEDIA_TYPE: Label 'application/json';
        Control: Codeunit "ControlInformes";



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

    procedure RestApi(url: Text; RequestType: Option Get,patch,put,post,delete; payload: Text; JsonIdentificador: Text): Text
    var
        Ok: Boolean;
        Respuesta: Text;
        TypeHelper: Codeunit "Type Helper";
    begin
        RequestHeaders := Client.DefaultRequestHeaders();
        //RequestHeaders.Add('Authorization', CreateBasicAuthHeader(Username, Password));

        case RequestType of
            RequestType::Get:
                Client.Get(URL, ResponseMessage);
            RequestType::patch:
                begin
                    RequestContent.WriteFrom(payload);

                    RequestContent.GetHeaders(contentHeaders);
                    contentHeaders.Clear();
                    contentHeaders.Add('Content-Type', 'application/json-patch+json');

                    RequestMessage.Content := RequestContent;

                    RequestMessage.SetRequestUri(URL);
                    RequestMessage.Method := 'PATCH';

                    client.Send(RequestMessage, ResponseMessage);
                end;
            RequestType::post:
                begin
                    //RequestContent.WriteFrom(payload);
                    payload := StrSubstNo(JsonIdentificador + '=%1', TypeHelper.UrlEncode(payload));
                    RequestContent.WriteFrom(payload);
                    RequestContent.GetHeaders(contentHeaders);
                    contentHeaders.Clear();
                    contentHeaders.Add('Content-Type', 'application/x-www-form-urlencoded');

                    Client.Post(URL, RequestContent, ResponseMessage);
                end;
            RequestType::delete:
                begin


                    Client.Delete(URL, ResponseMessage);
                end;
        end;

        ResponseMessage.Content().ReadAs(ResponseText);
        exit(ResponseText);

    end;

    Procedure JsonExcel(var Excel: Record "Excel Buffer 2"; Plantilla: Text; Url: Text): Text
    var
        ExcelBufferDialogMgt: Codeunit "Excel Buffer Dialog Management";
        carga: Codeunit "ControlInformes";
        // '{  "fileName": "C:\\temp\\Prueba.xlsx", "base64": "false", "sheets":';
        //[{    "sheetName": "Prueba",   
        //Data: [      {          "Row_No": 1,          "Column_No": 1,          "XlRowID": "1",          "XlColID": "A",          
        //"Cell_Value_as_Text": "Hola Muchacho",          "Cell_Type": 1,          "NumberFormat": "",          "Bold": true,          
        //"Italic": false,          "Underline": false,          "Double_Underline": false,          "Font_Name": "Arial",          
        //"Font_Size": 12,          "Font_Color": "Black",          "Background_Color": "White",          
        //"Vinculo": "http://192.168.10.226:8080/BC190/?company=SA%20VINYA%20DELS%20MOSCATELLS%2C%20S.L.&node=000138e7-a925-0000-1000-c100836bd2d2&page=50011&dc=0&bookmark=27_JAAAAACLAQAAAAJ7_1MAQQBDADIANAAtADAAMAAwADAAMQ",
        //          "Comment": "",          "Formula": "",          "Formula2": "",          "Formula3": "",          "Formula4": "",          
        //"Cell_Value_as_Blob": "",          "Formato_Columna": "",          "IsMark": false }]   } ] }';
        JsonObj: JsonObject;
        JsonSheet: JsonObject;
        JsonArray: JsonArray;
        JsonDataArray: JsonArray;
        JsonDataObject: JsonObject;
        JsonText: Text;
        RequestType: Option Get,patch,put,post,delete;
        Sheets: Record "Empresas informes" temporary;
        a: integer;
        TempBlob: Codeunit "Temp Blob";
        OUTSTREAM: OutStream;
        InSTREAM: InStream;
        Filename: Text;
        Base64: Text;
        Base64Convert: Codeunit "Base64 Convert";
        DocAttachment: Record "Document Attachment" temporary;
        Id: Array[100] of Integer;
        StreamManager: Codeunit "Stream Management";
        G: Guid;
        ChunkSize: Integer;
        Chunk: Text;
        StartPos: Integer;
        EndPos: Integer;
        i: Integer;
        HttpClient: HttpClient;
        HttpContent: HttpContent;
        HttpRequestMessage: HttpRequestMessage;
        HttpResponseMessage: HttpResponseMessage;
        Text005: Label 'Creando hoja Excel...\\', Comment = '{Locked="Excel"}';
        LastUpdate: Datetime;
        TotalRecNo: Integer;
        RecNo: Integer;
    begin
        ExcelBufferDialogMgt.Open(Text005);
        LastUpdate := CurrentDateTime;
        TotalRecNo := Excel.Count();
        if Excel.FindSet() Then begin
            repeat
                Sheets.SetRange(Empresa, Excel."Sheet Name");
                If Not Sheets.FindSet() then begin
                    a += 1;
                    Sheets.Id := a;
                    Sheets."Empresa" := Excel."Sheet Name";
                    Sheets.Insert();
                end;
            until Excel.Next() = 0;

            if Url <> '' then
                JsonObj.Add('fileName', url)
            else
                JsonObj.Add('fileName', Plantilla);
            if Url <> '' then
                JsonObj.Add('base64', 'url')
            else begin
                if Plantilla <> '' then
                    JsonObj.Add('base64', 'true')
                else
                    JsonObj.Add('base64', 'false');
            end;

        end;
        If Sheets.FindSet() then
            repeat
                RecNo := RecNo + 1;
                if not UpdateProgressDialog(ExcelBufferDialogMgt, LastUpdate, RecNo, TotalRecNo) then;
                Clear(JsonSheet);
                Clear(JsonArray);
                JsonSheet.Add('sheetName', Sheets."Empresa");
                Excel.SetRange("Sheet Name", Sheets."Empresa");
                Clear(JsonDataArray);
                Excel.SetFilter("Cell Value as Text", '<>%1', '');
                If Excel.FindSet() then
                    repeat
                        Clear(JsonDataObject);
                        JsonDataObject.Add('Row_No', Excel."Row No.");
                        JsonDataObject.Add('Column_No', Excel."Column No.");
                        JsonDataObject.Add('XlRowID', Excel."XlRowID");
                        JsonDataObject.Add('XlColID', Excel."XlColID");
                        JsonDataObject.Add('Cell_Value_as_Text', Excel."Cell Value as Text");
                        JsonDataObject.Add('Cell_Type', Excel."Cell Type");
                        If Excel.NumberFormat <> '' then
                            JsonDataObject.Add('NumberFormat', Excel.NumberFormat);
                        If Excel.Bold then
                            JsonDataObject.Add('Bold', Excel.Bold);
                        if Excel.Italic then
                            JsonDataObject.Add('Italic', Excel.Italic);
                        if Excel.Underline then
                            JsonDataObject.Add('Underline', Excel.Underline);
                        if Excel."Double Underline" then
                            JsonDataObject.Add('Double_Underline', Excel."Double Underline");
                        if Excel."Font Name" <> '' then
                            JsonDataObject.Add('Font_Name', Excel."Font Name");
                        if Excel."Font Size" <> 0 then
                            JsonDataObject.Add('Font_Size', Excel."Font Size");
                        if Excel."Font Color" <> '' then
                            JsonDataObject.Add('Font_Color', Excel."Font Color");
                        if Excel."Background Color" <> '' then
                            JsonDataObject.Add('Background_Color', Excel."Background Color");
                        if Excel.Vinculo <> '' then
                            JsonDataObject.Add('Vinculo', Excel.Vinculo);
                        If Excel.Comment <> '' then
                            JsonDataObject.Add('Comment', Excel.Comment);
                        if Excel.Formula <> '' then
                            JsonDataObject.Add('Formula', Excel.Formula);
                        if Excel.Formula2 <> '' then
                            JsonDataObject.Add('Formula2', Excel.Formula2);
                        if Excel.Formula3 <> '' then
                            JsonDataObject.Add('Formula3', Excel.Formula3);
                        if Excel.Formula4 <> '' then
                            JsonDataObject.Add('Formula4', Excel.Formula4);
                        //JsonDataObject.Add('Cell_Value_as_Blob', '');
                        if Excel."Formato Columna" <> '' then
                            JsonDataObject.Add('Formato_Columna', Excel."Formato Columna");
                        //JsonDataObject.Add('IsMark', false);
                        JsonDataArray.Add(JsonDataObject);
                    until Excel.Next() = 0;

                JsonSheet.Add('Data', JsonDataArray);
                JsonArray.Add(JsonSheet);
            until Sheets.Next() = 0;
        JsonObj.Add('sheets', JsonArray);
        JsonObj.WriteTo(JsonText);
        Tempblob.CREATEOUTSTREAM(OUTSTREAM, TextEncoding::UTF8);
        OUTSTREAM.WRITE(JsonText);
        TempBlob.CreateInStream(InSTREAM);

        // Base64 := Base64Convert.ToBase64(InSTREAM);
        // a := StrLen(Base64);
        // Clear(TempBlob);
        // G := CreateGuid();
        // //Enviar B64 en paquetes de 10*2014 a FromBase64URL

        // ChunkSize := 30 * 1024 * 1024; // 10 MB

        // StartPos := 1;
        // a := 0;
        // Clear(JsonObj);
        // clear(JsonArray);
        // clear(JsonSheet);
        // while StartPos <= StrLen(Base64) do begin
        //     a += 1;
        //     If a > 100 then
        //         Error('Error en la carga del fichero');
        //     if (StartPos + ChunkSize - 1) < StrLen(Base64) then
        //         EndPos := StartPos + ChunkSize - 1
        //     else
        //         EndPos := StrLen(Base64);
        //     Chunk := CopyStr(Base64, StartPos, EndPos - StartPos + 1);
        //     JsonText := DocAttachment.FormBase64ToUrl(Chunk, G + '#' + Format(a) + '#' + '.txt', Id[a]);
        //     Clear(JsonSheet);
        //     JsonSheet.Add('JsonText', JsonText);

        //     JsonArray.Add(JsonSheet);
        //     StartPos := EndPos + 1;
        // end;
        // JsonObj.Add('JsonText', JsonArray);
        // JsonObj.WriteTo(JsonText);
        //DownloadFromStream(InSTREAM, '', '', '', Filename);
        JsonText := SendStreamToWebService('http://192.168.10.226:81/MallaWebService.asmx/CreaOactualizaLibroBin', InSTREAM);
        //JsonText := carga.RestApi('http://192.168.10.226:81/MallaWebService.asmx/CreaOactualizaLibro', Requesttype::Post, JsonText, 'JsonText');

        //<?xml version="1.0" encoding="utf-8"?>
        //<string xmlns="http://tempuri.org/">UEsDBBQAAAAIABteeln0Lo9U7gAAALwBAAAPABwAeGwvd29ya2Jvb2sueG1sIKIYACigFAAAAAAAAAAAAAAAAAAAAAAAAAAAALXRUU+DMBAH8K/S3LsrMBhCxpa5meiDkjB9Xko5Rh1tSdspH986zWZ88mVv1/9dmt+18+Uoe/KOxgqtCggnARBUXDdC7Qs4uvbmFpaL+Zh/aHOotT4QP69sPhbQOTfklFreoWR2ogdUvtdqI5nzR7OndjDIGtshOtnTKAhmVDKh4Ou+U2rPFVFMYgEP+o2FQE7RY+M9QEwufFElYcrjOktZFmVxyhF+IOY/EN22guNG86NE5b4lBnvm/NK2E4MFQv9SNuX69en++aXc7la7dXlXrapfsOgMC8IW6ySJgmmYxGyWXQFGL89FLz+x+ARQSwMECgAAAAAAG156WW/aYHYoAQAAKAEAAAsAHABfcmVscy8ucmVscyCiGAUAABoAAAB4bC93b3JrYm9vay54bWwgogYAFAAAAAgAG17pZ9C6PVNuAAAAC8BAAADAAgAeGwvd29ya2Jvb2sueG1sUEsFBgAAAAABAAEANgAAAG4AAAAAAA==</string>
        //JsonText := Copystr(Jsontext, Strpos(JsonText, '<string xmlns="http://tempuri.org/">') + 36);
        //JsonText := Copystr(Jsontext, 1, Strpos(JsonText, '</string>') - 1);
        //Para cada id ejecutar DocAttachment.DeleteId;
        //for i := 1 to a do
        //  DocAttachment.DeleteId(Id[i]);
        ExcelBufferDialogMgt.Close();
        exit(JsonText);
    end;

    procedure SendStreamToWebService(Url: Text; InStream: InStream): Text
    var
        HttpClient: HttpClient;
        HttpRequestMessage: HttpRequestMessage;
        HttpResponseMessage: HttpResponseMessage;
        HttpContent: HttpContent;
        ResponseText: Text;
        Compresion: Codeunit "Data Compression";
        CompressedBlob: Codeunit "Temp Blob";
        TempBlob: Codeunit "Temp Blob";
        CompressedStream: OutStream;
        OutStream: OutStream;
        InCompressedStream: InStream;
        Base64Convert: Codeunit "Base64 Convert";
    begin
        // Crear la solicitud HTTP
        CompressedBlob.CreateOutStream(CompressedStream);
        Compresion.GZipCompress(InSTREAM, CompressedStream);
        CompressedBlob.CreateInStream(InStream);

        // Crear contenido HTTP desde el InStream
        HttpContent.WriteFrom(InStream); // Escribe directamente el contenido del InStream
        HttpRequestMessage.Content := HttpContent;
        HttpClient.Post(Url, HttpContent, HttpResponseMessage);
        // Enviar la solicitud
        // Leer la respuesta del servidor
        if HttpResponseMessage.IsSuccessStatusCode() then begin
            HttpResponseMessage.Content.ReadAs(InCompressedStream);
            TempBlob.CreateOutStream(OutStream);
            Compresion.GZipDecompress(InCompressedStream, OutStream);
            Clear(InStream);
            TempBlob.CreateInStream(InStream);
            //CopyStream(OutStream, InStream);
            ResponseText := Base64Convert.ToBase64(InStream);
        end else
            Error('Error al enviar la solicitud: %1', HttpResponseMessage.HttpStatusCode());

        exit(ResponseText);
    end;

    internal procedure UrlPlantillaInstream(var pUrlPlantilla: Text; var Informes: Record Informes; var PlantillaBase64: Text; Modificar: Boolean);
    var
        InExcelStream: InStream;
        Base64Convert: Codeunit "Base64 Convert";
        DocumentAttachment: Record "Document Attachment";
        Id: Integer;
        OutStream: OutStream;
        Tempplob: Codeunit "Temp Blob";
    begin
        if Informes."Url Plantilla" = '' Then begin
            Informes.CalcFields("Plantilla Excel");
            Informes."Plantilla Excel".CreateInStream(InExcelStream);
            PlantillaBase64 := Base64Convert.ToBase64(InExcelStream);
            pUrlPlantilla := DocumentAttachment.FormBase64ToUrl(PlantillaBase64, 'Plantilla' + Format(Informes.Id) + '.xlsx', Id);
            If Modificar then begin
                Informes."Url Plantilla" := pUrlPlantilla;
                //Informes.Modify();
            end;

        end else
            pUrlPlantilla := Informes."Url Plantilla";
        PlantillaBase64 := DocumentAttachment.ToBase64StringOcr(pUrlPlantilla);


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


}
