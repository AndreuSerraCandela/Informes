page 7001196 "Destinatarios Informes"
{
    PageType = ListPart;
    SourceTable = "Destinatarios Informes";
    layout
    {
        area(content)
        {
            repeater(Repeater)
            {
                field("Id"; Rec."Id")
                {
                    ApplicationArea = All;
                }
                field("Empleado"; Rec."Empleado")
                {
                    ApplicationArea = All;
                }
                field("e-mail"; Rec."e-mail")
                {
                    ApplicationArea = All;
                }
                field("Nombre Informe"; Rec."Nombre Informe")
                {
                    ApplicationArea = All;
                }

                field("Campo Destinatario"; Rec."Campo Destinatario")
                {


                }
                field(Nombre; NombreCampo())
                {
                    ApplicationArea = All;
                }
                field(Valor; Rec.Valor)
                {

                }
                field("No enviar"; Rec."No enviar")
                {
                    ApplicationArea = All;
                }

            }
        }
    }
    procedure NombreCampo(): Text
    var
        Columna: Record "Columnas Informes";
        Informes: Record "Informes";
        Campos: Record Field;
    begin
        if Rec."Campo Destinatario" = 0 then
            exit('');
        if Informes.Get(Rec.Id) then begin
            Columna.SetRange(Columna."Id", Informes."Id");
            If not Columna.FindFirst() then begin
                Informes.InitDefaultColumns();
                Informes.GetColumns(Informes);
                Commit();
                Columna.FindFirst();
            end;
            Campos.SetRange(tableno, Columna.Table);
            Campos.SetRange("No.", Rec."Campo Destinatario");
            If Campos.FindFirst then
                exit(Campos."Field Caption");
        end;
    end;

}
page 7001189 "Lista Destinatarios Informes"
{
    PageType = List;
    SourceTable = "Destinatarios Informes";
    SourceTableTemporary = true;
    layout
    {
        area(content)
        {
            repeater(Repeater)
            {

                field("Id"; Rec."Id")
                {
                    ApplicationArea = All;
                }
                field("Empleado"; Rec."Empleado")
                {
                    ApplicationArea = All;
                }
                field("e-mail"; Rec."e-mail")
                {
                    ApplicationArea = All;
                }
                field("Nombre Informe"; Rec."Nombre Informe")
                {
                    ApplicationArea = All;
                }

                field("Campo Destinatario"; Rec."Campo Destinatario")
                {


                }
                field(Nombre; NombreCampo())
                {
                    ApplicationArea = All;
                }
                field(Valor; Rec.Valor)
                {

                }
                field("No enviar"; Rec."No enviar")
                {
                    ApplicationArea = All;
                }


            }
        }
    }

    trigger OnAfterGetRecord()
    begin
        CampoInformes := Rec."Campo Destinatario";
    end;

    trigger OnAfterGetCurrRecord()
    begin
        CampoInformes := Rec."Campo Destinatario";
    end;

    procedure NombreCampo(): Text
    var
        Columna: Record "Columnas Informes";
        Informes: Record "Informes";
        Campos: Record Field;
    begin
        if Rec."Campo Destinatario" = 0 then
            exit('');
        if Informes.Get(Rec.Id) then begin
            Columna.SetRange(Columna."Id", Informes."Id");
            If not Columna.FindFirst() then begin
                Informes.InitDefaultColumns();
                Informes.GetColumns(Informes);
                Commit();
                Columna.FindFirst();
            end;
            Campos.SetRange(tableno, Columna.Table);
            Campos.SetRange("No.", Rec."Campo Destinatario");
            If Campos.FindFirst then
                exit(Campos."Field Caption");
        end;
    end;

    Procedure CargaTabla(PTableId: Integer)
    begin
        TableId := PTableId;
    end;



    procedure CargarDatos(Var Destinatarios: Record "Destinatarios Informes" temporary)
    begin
        If Destinatarios.FindFirst() then
            repeat
                Rec := Destinatarios;
                Rec.Insert();
            until Destinatarios.Next() = 0;
    end;

    procedure DesCargarDatos(Var Destinatarios: Record "Destinatarios Informes" temporary)
    begin
        Destinatarios.DeleteAll();
        If Rec.FindFirst() then
            repeat
                Destinatarios := Rec;
                Destinatarios.Insert();
            until Rec.Next() = 0;
    end;

    var
        TableId: Integer;
        CampoInformes: Integer;

}