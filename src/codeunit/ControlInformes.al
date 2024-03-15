Codeunit 7001130 ControlInformes
{
    trigger OnRun()
    var

    begin

        imprimirInformes(0, WorkDate());
    end;


    procedure imprimirInformes(IdInforme: Integer; ProximaFecha: Date)
    var
        Informe: Record "Informes";
        Destinatario: Record "Destinatarios Informes";
        Filtros: Record "Filtros Informes";
        Contratos: Page "Lista Contratos x Empresa";
        Out: OutStream;
        ficheros: Record Ficheros;
        Secuencia: Integer;
    begin
        If IdInforme <> 0 tHEN
            Informe.SetRange("ID", IdInforme);
        If ProximaFecha <> 0D tHEN
            Informe.SetRange("Próxima Fecha", WorkDate());
        If Informe.FindFirst() then begin
            repeat
                Destinatario.Reset;
                Destinatario.SetRange("ID", Informe."ID");
                if Destinatario.FindSet() then begin
                    repeat
                        Filtros.Reset;
                        Filtros.SetRange("ID", Informe."ID");
                        if Filtros.FindSet() then begin
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
                                        Contratos.ExportExcel(Filtros, Destinatario, out);
                                    end;
                            end;
                            ficheros.Modify();
                            Commit();
                            EnviaCorreoComercial(Destinatario."e-mail", ficheros, Informe.Descripcion);
                        end;
                    until Destinatario.Next() = 0;
                end;
                IF ProximaFecha <> 0D tHEN begin
                    Informe."Próxima Fecha" := CalcDate(Informe.Periodicidad, Informe."Próxima Fecha");
                    Informe.Modify;
                end;

            until Informe.Next() = 0;
        end;
    end;

    local procedure EnviaCorreoComercial(SalesPersonMail: Text; var ficheros: Record Ficheros; Informe: Text)


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


    begin
        rInf.Get();
        BigText := ('Estimado:');

        //(FORMAT(cr,0,'<CHAR>') + FORMAT(lf,0,'<CHAR>')
        BigText := BigText + '<br> </br>';
        BigText := BigText + '<br> </br>';
        BigText := BigText + 'Adjunto Informe: ' + Informe;
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
        BigText := BigText + ('Dpto. Medios');
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
        REmail."Send BCC" := 'andreuserra@malla.es';
        REmail.SetBodyText(BigText);
        REmail."From Name" := UserId;


        ficheros.CalcFields(Fichero);
        ficheros.Fichero.CreateInStream(AttachmentStream);
        REmail.AddAttachment(AttachmentStream, Informe + '.xlsx');

        // if REmail."From Address" <> '' Then
        //     REmail."Send BCC" := REmail."From Address" else
        //     REmail."Send BCC" := BCC();
        REmail.Send(true, emilesc::Default);
        ficheros.delete;

    end;
}
