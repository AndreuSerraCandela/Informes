
page 7001198 "Informes"
{
    PageType = List;
    SourceTable = Informes;
    UsageCategory = Lists;
    CardPageId = "Informes Card";
    ApplicationArea = All;
    layout
    {
        area(content)
        {
            repeater(Informes)
            {
                field(Id; Rec.Id)
                {
                    ApplicationArea = All;
                }
                field(Descripcion; Rec.Descripcion)
                {
                    ApplicationArea = All;
                }
                field(Periodicidad; Rec.Periodicidad)
                {
                    ApplicationArea = All;
                }
                field("Fecha Primera Ejecución"; Rec."Fecha Primera Ejecución")
                {
                    ApplicationArea = All;
                }
                field(Informe; Rec.Informe)
                {
                    ApplicationArea = All;
                }
                field("Tabla filtros"; Rec."Tabla filtros")
                {
                    ApplicationArea = All;
                }
            }
        }
    }
}
