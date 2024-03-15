page 7001193 "Campos Informes"
{

    PageType = ListPart;
    SourceTable = "Campos Informes";
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
                field(Orden; Rec.Orden)
                {
                    ApplicationArea = All;
                }
                field("Campo"; Rec."Campo")
                {
                    ApplicationArea = All;
                }
                field(Funcion; Rec.Funcion)
                {
                    ApplicationArea = All;
                }
                field("Nombre Campo"; NombreCampo())
                {
                    ApplicationArea = All;
                }
                field(Título; Rec.Titulo)
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
        if Rec.Campo = 0 then
            exit('');
        if Informes.Get(Rec.Id) then begin
            Campos.SetRange(Campos.TableNo, Informes."Tabla filtros");
            Campos.SetRange("No.", Rec.Campo);
            If Campos.FindFirst then
                exit(Campos."Field Caption");
        end;
    end;
}
