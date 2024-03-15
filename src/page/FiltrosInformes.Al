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
        Campos: Record Field;
        Informes: Record "Informes";
    begin
        if Informes.Get(Rec.Id) then begin
            Campos.SetRange(Campos.TableNo, Informes."Tabla filtros");
            Campos.SetRange("No.", Rec.Campo);
            If Campos.FindFirst then
                exit(Campos."Field Caption");
        end;
    end;
}
