//Crear page Enlaces
page 7001201 "Enlaces"
{
    PageType = List;
    SourceTable = "Enlaces Informes";
    Caption = 'Enlaces Informes';
    layout
    {
        area(content)
        {
            repeater(Detalle)
            {
                field("Id"; Rec.Id)
                {
                    ApplicationArea = All;
                    Visible = false;
                }
                field("Campo Relación"; Rec."Campo Relación")
                {
                    Visible = false;
                    ApplicationArea = All;
                }
                field("Nombre Campo Relación"; Rec."Nombre Campo Relación")
                {
                    Visible = false;
                    ApplicationArea = All;
                }
                field("Tabla"; Rec."Tabla")
                {
                    ApplicationArea = All;
                }
                field("Campo Enlace"; Rec."Campo Enlace")
                {
                    ApplicationArea = All;
                }
                field("Field Name Enlace"; Rec."Field Name Enlace")
                {
                    ApplicationArea = All;
                }
                field("Campo Datos"; Rec."Campo Datos")
                {
                    ApplicationArea = All;
                }
                field("Field Name Datos"; Rec."Field Name Datos")
                {
                    ApplicationArea = All;
                }
            }
        }
    }

}