
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

                field("Próxima Fecha"; Rec."Próxima Fecha")
                {
                    ApplicationArea = All;
                }
                field("Tabla filtros"; Rec."Tabla filtros")
                {
                    ApplicationArea = All;
                    trigger OnValidate()
                    var
                        AllObjWithCaption: record AllObjWithCaption;
                    begin
                        AllObjWithCaption.GET(AllObjWithCaption."Object Type"::Table, Rec."Tabla filtros");
                        "Nombre Tabla" := AllObjWithCaption."Object Caption";
                        CurrPage.UPDATE(false);
                    end;
                }
                field("Nombre"; "Nombre Tabla")
                {
                    ApplicationArea = All;
                    trigger OnAssistEdit()
                    var
                        AllObjWithCaption: record AllObjWithCaption;
                    begin
                        AllObjWithCaption.SetRange(AllObjWithCaption."Object Type", AllObjWithCaption."Object Type"::Table);
                        If Page.RunModal(0, AllObjWithCaption) = Action::LookupOK Then
                            Rec."Tabla filtros" := AllObjWithCaption."Object ID";
                        "Nombre Tabla" := AllObjWithCaption."Object Caption";
                        CurrPage.UPDATE(false);
                    end;
                }
                field("Tipo Objeto"; Rec."Tipo Objeto")
                {
                    ApplicationArea = All;
                    // Editable = OtrosInformes;
                }
                field("Id Objeto"; Rec."Id Objeto")
                {
                    ApplicationArea = All;
                    //  Editable = OtrosInformes;
                    trigger OnValidate()
                    var
                        AllObjWithCaption: record AllObjWithCaption;
                    begin
                        AllObjWithCaption.GET(Rec."Tipo Objeto", Rec."Id Objeto");
                        "Nombre Query" := AllObjWithCaption."Object Caption";
                        CurrPage.UPDATE(false);
                    end;
                }
                field("Nombre Objeto"; "Nombre Query")
                {
                    ApplicationArea = All;
                    //   Enabled = OtrosInformes;
                    trigger OnAssistEdit()
                    var
                        AllObjWithCaption: record AllObjWithCaption;
                    begin
                        AllObjWithCaption.SetRange(AllObjWithCaption."Object Type", Rec."Tipo Objeto");
                        If Page.RunModal(0, AllObjWithCaption) = Action::LookupOK Then
                            Rec."Id Objeto" := AllObjWithCaption."Object ID";
                        "Nombre Query" := AllObjWithCaption."Object Caption";
                        CurrPage.UPDATE(false);
                    end;
                }
                field("Crear Tarea"; Rec."Crear Tarea")
                {
                    ApplicationArea = All;
                }
                field("Descripcion Tarea"; Rec."Descripcion Tarea")
                {
                    ApplicationArea = All;
                }
            }
        }
    }
    trigger OnAfterGetRecord()

    var
        "Nombre Tabla": Text;
        AllObjWithCaption: record AllObjWithCaption;
    begin
        If AllObjWithCaption.GET(AllObjWithCaption."Object Type"::Table, Rec."Tabla filtros") then
            "Nombre Tabla" := AllObjWithCaption."Object Caption";
        if AllObjWithCaption.GET(Rec."Tipo Objeto", Rec."Id Objeto") then
            "Nombre Query" := AllObjWithCaption."Object Caption";
        OtrosInformes := rec.Informe.AsInteger() > 1;
    end;

    var
        "Nombre Tabla": Text;

        "Nombre Query": Text;
        OtrosInformes: Boolean;
}
