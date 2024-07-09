page 7001200 "Formato Columnas"
{
    Caption = 'Formato Informes';
    //DeleteAllowed = false;
    //InsertAllowed = false;
    PageType = List;
    SourceTable = "formato Columnas";
    InsertAllowed = false;
    DeleteAllowed = false;

    layout
    {
        area(content)
        {
            repeater(detalle)
            {
                field(Cabecera; Rec.Cabecera)
                {
                    ApplicationArea = All;
                    ToolTip = 'Especifica si el formato se refiere a la cabecera del informe.';


                }

                field(Orden; Rec.Orden)
                {
                    ApplicationArea = All;
                    Editable = false;
                }
                field("Columna"; Rec.Letra)
                {
                    ApplicationArea = All;
                    Editable = false;
                }
                field(Formato; Rec."Formato Columna")
                {
                    ApplicationArea = All;
                }

                field("Fuente"; Rec.Fuente)
                {
                    ApplicationArea = All;

                }
                field(Bold; Rec.Bold)
                {
                    ApplicationArea = All;

                }
                field(Italic; Rec.Italic)
                {
                    ApplicationArea = All;

                }
                field(Underline; Rec.Underline)
                {
                    ApplicationArea = All;

                }
                field("Double Underline"; Rec."Double Underline")
                {
                    ApplicationArea = All;

                }

                field("Tamaño"; Rec."Tamaño")
                {
                    ApplicationArea = All;

                }
                field("Color Fuente"; Rec."Color Fuente")
                {
                    ApplicationArea = All;
                    trigger OnDrillDown()
                    var
                        Colores: Record "Colores";
                        PageColores: Page "Colores";
                    begin
                        Clear(PageColores);
                        PageColores.RunModal();
                        PageColores.GetRecord(Colores);
                        Rec."Color Fuente" := Colores."Color Excel";
                    end;

                }
                field("Color Fondo"; Rec."Color Fondo")
                {
                    ApplicationArea = All;
                    trigger OnDrillDown()
                    var
                        Colores: Record "Colores";
                        PagColores: Page "Colores";
                    begin
                        Clear(PagColores);
                        PagColores.RunModal();
                        PagColores.GetRecord(Colores);
                        Rec."Color Fondo" := Colores."Color Excel";
                    end;

                }

            }
        }
    }

    actions
    {
        area(processing)
        {
        }
    }

    var
        RecRef: RecordRef;
        IsModified: Boolean;




}
