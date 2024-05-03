page 7001197 "Filtros Informes"
{

    PageType = ListPart;
    SourceTable = "Filtros Informes";
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
                field("Campo"; Rec."Campo")
                {
                    ApplicationArea = All;
                }
                field("Nombre Campo"; NombreCampo())
                {
                    ApplicationArea = All;
                }
                field("Valor"; Rec."Valor")
                {
                    ApplicationArea = All;
                }
                field("Desde"; Rec."Desde")
                {
                    ApplicationArea = All;
                }
                field("Hasta"; Rec."Hasta")
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
        if Rec.Campo = 0 then
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
            Campos.SetRange("No.", Rec.Campo);
            If Campos.FindFirst then
                exit(Campos."Field Caption");
        end;
    end;
}
page 7001190 "Lista Filtros Informes"
{

    PageType = List;
    SourceTable = "Filtros Informes";
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
                field("Campo"; CampoInformes)
                {
                    ApplicationArea = All;
                    trigger OnLookup(var Text: Text): Boolean
                    var
                        Campos: Record Field;
                    begin
                        Campos.SetRange(Campos.TableNo, TableId);
                        If Page.Runmodal(9806, Campos) = Action::LookupOK then begin
                            Rec.Campo := Campos."No.";
                            CampoInformes := Rec.Campo;
                            exit(true);
                        end;
                        exit(false);
                    end;

                }
                field("Nombre Campo"; NombreCampo())
                {
                    ApplicationArea = All;
                }
                field("Valor"; Rec."Valor")
                {
                    ApplicationArea = All;
                }
                field("Desde"; Rec."Desde")
                {
                    ApplicationArea = All;
                }
                field("Hasta"; Rec."Hasta")
                {
                    ApplicationArea = All;
                }
            }
        }
    }
    trigger OnAfterGetRecord()
    begin
        CampoInformes := Rec.Campo;
    end;

    trigger OnAfterGetCurrRecord()
    begin
        CampoInformes := Rec.Campo;
    end;

    procedure NombreCampo(): Text
    var

        Campos: Record Field;
    begin
        if Rec.Campo = 0 then
            exit('');

        Campos.SetRange(tableno, TableId);
        Campos.SetRange("No.", Rec.Campo);
        If Campos.FindFirst then
            exit(Campos."Field Caption");

    end;

    Procedure CargaTabla(PTableId: Integer)
    begin
        TableId := PTableId;
    end;



    procedure CargarDatos(Var Filtros: Record "Filtros Informes" temporary)
    begin
        If Filtros.FindFirst() then
            repeat
                Rec := Filtros;
                Rec.Insert();
            until Filtros.Next() = 0;
    end;

    procedure DesCargarDatos(Var Filtros: Record "Filtros Informes" temporary)
    begin
        Filtros.DeleteAll();
        If Rec.FindFirst() then
            repeat
                Filtros := Rec;
                Filtros.Insert();
            until Rec.Next() = 0;
    end;

    var
        TableId: Integer;
        CampoInformes: Integer;

}

