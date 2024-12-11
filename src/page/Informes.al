
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
                // field("Date Formula"; DTF)
                // {
                //     ApplicationArea = All;
                //     trigger OnValidate()
                //     begin
                //         CalcularDate();
                //     end;
                // }
                // field(Fecha; FechaFormula)
                // {
                //     ApplicationArea = All;
                //     trigger OnValidate()
                //     begin
                //         CalcularDate();
                //     end;
                // }
                // field(Resultado; Resultado)
                // {
                //     ApplicationArea = All;
                // }


                field("Fecha/Hora Inicio"; Rec."Earliest Start Date/Time")
                {
                    ApplicationArea = Basic, Suite;
                    ToolTip = 'Specifies the earliest date and time when the job queue entry should be run. The format for the date and time must be month/day/year hour:minute, and then AM or PM. For example, 3/10/2021 12:00 AM.';
                }
                field("Fecha expiracion"; Rec."Expiration Date/Time")
                {
                    ApplicationArea = Basic, Suite;
                    Importance = Additional;
                    ToolTip = 'Specifies the date and time when the job queue entry is to expire, after which the job queue entry will not be run.  The format for the date and time must be month/day/year hour:minute, and then AM or PM. For example, 3/10/2021 12:00 AM.';
                }

                field("Lunes"; Rec."Run on Mondays")
                {
                    ApplicationArea = Basic, Suite;
                    ToolTip = 'Specifies that the job queue entry runs on Mondays.';
                }
                field("Martes"; Rec."Run on Tuesdays")
                {
                    ApplicationArea = Basic, Suite;
                    ToolTip = 'Specifies that the job queue entry runs on Tuesdays.';
                }
                field("Miércoles"; Rec."Run on Wednesdays")
                {
                    ApplicationArea = Basic, Suite;
                    ToolTip = 'Specifies that the job queue entry runs on Wednesdays.';
                }
                field("Jueves"; Rec."Run on Thursdays")
                {
                    ApplicationArea = Basic, Suite;
                    ToolTip = 'Specifies that the job queue entry runs on Thursdays.';
                }
                field("Viernes"; Rec."Run on Fridays")
                {
                    ApplicationArea = Basic, Suite;
                    ToolTip = 'Specifies that the job queue entry runs on Fridays.';
                }
                field("Sábados"; Rec."Run on Saturdays")
                {
                    ApplicationArea = Basic, Suite;
                    ToolTip = 'Specifies that the job queue entry runs on Saturdays.';
                }
                field("Domingos"; Rec."Run on Sundays")
                {
                    ApplicationArea = Basic, Suite;
                    ToolTip = 'Specifies that the job queue entry runs on Sundays.';
                }
                field("Periodicidad"; Rec.Periodicidad)
                {
                    ApplicationArea = Basic, Suite;
                    Caption = 'Periodicidad';



                }
                field("Hora Inicio"; Rec."Starting Time")
                {
                    ApplicationArea = Basic, Suite;
                    Importance = Promoted;
                    ToolTip = 'Specifies the earliest time of the day that the recurring job queue entry is to be run.';
                }
                field("Hora Fin"; Rec."Ending Time")
                {
                    ApplicationArea = Basic, Suite;
                    Importance = Promoted;
                    ToolTip = 'Specifies the latest time of the day that the recurring job queue entry is to be run.';
                }
                field("Minutos entre ejecuciones"; Rec."No. of Minutes between Runs")
                {
                    ApplicationArea = Basic, Suite;
                    Importance = Promoted;
                    ToolTip = 'Specifies the minimum number of minutes that are to elapse between runs of a job queue entry. The value cannot be less than one minute. This field only has meaning if the job queue entry is set to be a recurring job. If you use a no. of minutes between runs, the date formula setting is cleared.';
                }

                // field("Tabla filtros"; Rec."Tabla filtros")
                // {
                //     ApplicationArea = All;
                //     trigger OnValidate()
                //     var
                //         AllObjWithCaption: record AllObjWithCaption;
                //     begin
                //         AllObjWithCaption.GET(AllObjWithCaption."Object Type"::Table, Rec."Tabla filtros");
                //         "Nombre Tabla" := AllObjWithCaption."Object Caption";
                //         CurrPage.UPDATE(false);
                //     end;
                // }
                // field("Nombre"; "Nombre Tabla")
                // {
                //     ApplicationArea = All;
                //     trigger OnAssistEdit()
                //     var
                //         AllObjWithCaption: record AllObjWithCaption;
                //     begin
                //         AllObjWithCaption.SetRange(AllObjWithCaption."Object Type", AllObjWithCaption."Object Type"::Table);
                //         If Page.RunModal(0, AllObjWithCaption) = Action::LookupOK Then
                //             Rec."Tabla filtros" := AllObjWithCaption."Object ID";
                //         "Nombre Tabla" := AllObjWithCaption."Object Caption";
                //         CurrPage.UPDATE(false);
                //     end;
                // }
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
                field("Formato Json"; Rec."Formato Json")
                {
                    ApplicationArea = All;
                }
            }
        }

    }
    actions
    {
        area(processing)
        {
            action("Crear un Nuevo Informe")
            {
                ApplicationArea = All;
                Image = StepInto;
                trigger OnAction()
                var

                begin
                    page.RunModal(Page::"Informes Setup Wizard");
                end;
            }
        }
        area(Promoted)
        {
            actionref("Crea_ref"; "Crear un Nuevo Informe") { }
        }
    }
    trigger OnAfterGetRecord()

    var
        "Nombre Tabla": Text;
        AllObjWithCaption: record AllObjWithCaption;
        Columnas: record "Columnas Informes";
    begin
        // If AllObjWithCaption.GET(AllObjWithCaption."Object Type"::Table, Rec."Tabla filtros") then
        //     "Nombre Tabla" := AllObjWithCaption."Object Caption";
        if AllObjWithCaption.GET(Rec."Tipo Objeto", Rec."Id Objeto") then
            "Nombre Query" := AllObjWithCaption."Object Caption";
        OtrosInformes := rec.Informe.AsInteger() > 1;
        Columnas.SetRange(Id, Rec."Id");
        If Columnas.FindFirst() Then
            if AllObjWithCaption.GET(AllObjWithCaption."Object Type"::Table, columnas.Table) then
                "Nombre Tabla" := AllObjWithCaption."Object Caption";
    end;

    var
        "Nombre Tabla": Text;

        "Nombre Query": Text;
        OtrosInformes: Boolean;
        DTF: DateFormula;
        DT: DateFormula;
        FechaFormula: Date;
        Resultado: Date;

    procedure CalcularDate()
    begin
        if DTF <> DT then begin
            If FechaFormula <> 0D then
                Resultado := CalcDate(DTF, FechaFormula);
            CurrPage.UPDATE(false);
        end;
    end;
}
