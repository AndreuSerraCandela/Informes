page 7001192 "Columnas Informes"
{
    Caption = 'Columnas Informes';
    //DeleteAllowed = false;
    InsertAllowed = false;
    PageType = ListPart;
    SourceTable = "Columnas Informes";


    layout
    {
        area(content)
        {
            repeater(detalle)
            {
                field(Incluir; Rec.Include)
                {
                    ApplicationArea = Basic, Suite;
                    ToolTip = 'Specifies the field name that is selected in the data set.';

                    trigger OnValidate()
                    begin
                        // if CalledForExcelExport then
                        //     CheckFieldFilter();
                        IsModified := true;
                    end;
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
                field("Field Name"; Rec."Field Name")
                {
                    ApplicationArea = Basic, Suite;
                    Caption = 'Nombre del campo';
                    Editable = false;
                    ToolTip = 'Specifies the field names in a data set.';
                }


                field("Título"; Rec.Titulo)
                {
                    ApplicationArea = Basic, Suite;
                    ToolTip = 'Specifies the Field Captions in a data set.';
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
