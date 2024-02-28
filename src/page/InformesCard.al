
page 7001195 "Informes Card"
{
    PageType = Card;
    SourceTable = Informes;
    layout
    {
        area(content)
        {
            group(General)
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
            part(Destinatarios; "Destinatarios Informes")
            {
                ApplicationArea = All;
                SubPageLink = Id = fIELD(Id);
            }
            part(Filtros; "Filtros Informes")
            {
                ApplicationArea = All;
                SubPageLink = Id = fIELD(Id);
            }

        }
    }
}
