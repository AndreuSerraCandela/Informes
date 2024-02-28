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

}
