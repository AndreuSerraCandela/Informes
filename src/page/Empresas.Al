page 7001187 "Empresas Informes"
{

    PageType = ListPart;
    SourceTable = "Empresas Informes";
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
                field("Empresa"; Rec.Empresa)
                {
                    ApplicationArea = All;
                }
                field("Hoja Excel"; Rec."HojaExcel")
                {
                    ApplicationArea = All;
                }
                field("Hojas Separadas"; Rec."Hojas Separadas")
                {
                    ApplicationArea = All;
                }
                field("Columna Excel"; Rec."Columna Excel")
                {
                    ApplicationArea = All;
                }

            }
        }
    }

}


