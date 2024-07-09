page 7001186 "Años Informes"
{

    PageType = ListPart;
    SourceTable = "Periodos Informes";
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
                field("Periodo"; Rec."Periodo")
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

                field("Desde"; Rec."Desde")
                {
                    ApplicationArea = All;
                }
                field("Hasta"; Rec."Hasta")
                {
                    ApplicationArea = All;
                }
                field(Semana; Rec.Semana)
                {
                    ApplicationArea = All;
                }
                field("Campo 2"; Rec."Campo2")
                {
                    ApplicationArea = All;
                }
                field("Nombre Campo 2"; NombreCampo2())
                {
                    ApplicationArea = All;
                }

                field("Desde 2"; Rec."Desde2")
                {
                    ApplicationArea = All;
                }
                field("Hasta 2"; Rec."Hasta2")
                {
                    ApplicationArea = All;
                }
                field("Semana 2"; Rec.Semana2)
                {
                    ApplicationArea = All;
                }
            }
        }
    }
    trigger OnAfterGetRecord()
    var
        Informes: Record "Informes";

    begin
        if Informes.Get(Rec.Id) then
            If Informes.Informe = Informes.Informe::"Contratos x Empresa" Then CurrPage.Editable := false;
    end;

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

    procedure NombreCampo2(): Text
    var
        Columna: Record "Columnas Informes";
        Informes: Record "Informes";
        Campos: Record Field;
    begin
        if Rec.Campo2 = 0 then
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
            Campos.SetRange("No.", Rec.Campo2);
            If Campos.FindFirst then
                exit(Campos."Field Caption");
        end;
    end;
}
