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
            }
        }
    }
}