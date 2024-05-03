
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
                    Editable = false;
                }
                field(Descripcion; Rec.Descripcion)
                {
                    ApplicationArea = All;
                }
                group("Work Description")
                {
                    Caption = 'Descripción Informe';
                    field(WorkDescription; WorkDescription)
                    {
                        ApplicationArea = all;
                        MultiLine = true;
                        ShowCaption = false;
                        ToolTip = 'Especifique la descripción del informe.';

                        trigger OnValidate()
                        begin
                            Rec.SetDescripcionAmpliada(WorkDescription);
                        end;
                    }
                }
                group(Recurrencia)
                {
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
                }
                field(Informe; Rec.Informe)
                {
                    ApplicationArea = All;
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
                field("Nombre"; "Nombre Tabla")
                {
                    ApplicationArea = All;
                    Editable = false;

                }
                field("Tipo Objeto"; Rec."Tipo Objeto")
                {
                    ApplicationArea = All;
                    //  Editable = OtrosInformes;
                }
                field("Id Objeto"; Rec."Id Objeto")
                {
                    ApplicationArea = All;
                    //Editable = OtrosInformes;
                    trigger OnValidate()
                    var
                        AllObjWithCaption: record AllObjWithCaption;
                        Columnas: Record "Columnas Informes";
                        CreateOrCopy: Option "Create a new data set","Create a copy of an existing data set","Edit an existing data set";
                        "source Service Name": Text;
                        "Destination Service Name": Text;
                    begin
                        AllObjWithCaption.GET(Rec."Tipo Objeto", Rec."Id Objeto");
                        "Nombre Query" := AllObjWithCaption."Object Caption";
                        Columnas.SetRange(Id, Rec."Id");
                        if Columnas.Count = 0 Then begin
                            If Rec."Id Objeto" <> 0 Then
                                Rec.InitColumns(Rec."Tipo Objeto", Rec."Id Objeto", CreateOrCopy::"Create a new data set", "Source Service Name", "Destination Service Name");
                            Rec.GetColumns(Rec);
                            CurrPage.UPDATE(false);
                            If Columnas.FindFirst() Then
                                if AllObjWithCaption.GET(AllObjWithCaption."Object Type"::Table, columnas.Table) then
                                    "Nombre Tabla" := AllObjWithCaption."Object Caption";
                        end;
                    end;
                }
                field("Nombre Objeto"; "Nombre Query")
                {
                    ApplicationArea = All;
                    Editable = false;
                    //Enabled = OtrosInformes;

                }
                field("Crear Tarea"; Rec."Crear Tarea")
                {
                    ApplicationArea = All;
                }
                field("Descripcion Tarea"; Rec."Descripcion Tarea")
                {
                    ApplicationArea = All;
                }
                field("Campo Tarea"; CampoTarea)
                {
                    ApplicationArea = All;
                    trigger OnLookup(var Text: Text): Boolean
                    var
                        Campos: Record Field;
                    begin
                        Campos.SetRange(Campos.TableNo, TableId);
                        If Page.Runmodal(9806, Campos) = Action::LookupOK then begin
                            Rec."Campo Tarea" := Campos."No.";
                            CampoTarea := Rec."Campo Tarea";
                            exit(true);
                        end;
                        exit(false);
                    end;

                }
                field("Nombre Campo"; NombreCampo())
                {
                    ApplicationArea = All;
                }
                field("Crear Web Service"; Rec."Crear Web Service")
                {
                    ApplicationArea = All;
                }
            }
            part(Destinatarios; "Destinatarios Informes")
            {
                ApplicationArea = All;
                SubPageLink = Id = fIELD(Id);

            }
            // part(Campos; "Campos Informes")
            // {
            //     ApplicationArea = All;
            //     SubPageLink = Id = fIELD(Id);
            // }
            part(Columnas; "Columnas Informes")
            {
                ApplicationArea = All;
                SubPageLink = Id = fIELD(Id);
                UpdatePropagation = Both;
            }
            part(Filtros; "Filtros Informes")
            {
                ApplicationArea = All;
                SubPageLink = Id = fIELD(Id);
            }
            part(Empresas; "Empresas Informes")
            {
                ApplicationArea = All;
                SubPageLink = Id = fIELD(Id);
            }
            part(Años; "Años Informes")
            {
                ApplicationArea = All;
                SubPageLink = Id = fIELD(Id);
            }

        }
    }
    // añadir botón para imprimir informes
    actions
    {
        area(Processing)
        {
            action(Print)
            {
                ApplicationArea = All;
                Image = Print;
                Caption = 'Imprimir';
                trigger OnAction()
                var
                    Informes: Codeunit ControlInformes;
                begin
                    Informes.imprimirInformes(Rec.Id, 0DT, false);// Código para imprimir informe
                end;
            }
            action(Guardar)
            {
                ApplicationArea = All;
                Image = Save;
                Caption = 'Guardar';
                trigger OnAction()
                var
                    Informes: Codeunit ControlInformes;
                begin
                    Informes.imprimirInformes(Rec.Id, 0DT, true);// Código para imprimir informe
                end;
            }

            action("Importar Plantilla")
            {
                ApplicationArea = All;
                Image = Excel;
                trigger OnAction()
                var
                    NVInStream: InStream;
                    OutStr: OutStream;
                    TempBlob: Codeunit "Temp Blob";
                    Base64: Codeunit "Base64 Convert";
                    Base64Txt: Text;
                    RecRf: RecordRef;
                    Plantilla: Text;
                begin
                    UPLOADINTOSTREAM('Import', '', ' Excel Files (*.xls)|*.xls;*.xlsx', Plantilla, NVInStream);
                    Base64Txt := Base64.ToBase64(NVInStream);
                    TempBlob.CreateOutStream(OutStr);
                    Base64.FromBase64(Base64Txt, OutStr);
                    RecRf.Get(Rec.RecordId);
                    TempBlob.ToRecordRef(RecRf, Rec.FieldNo("Plantilla Excel"));
                    // Rec."Certificado firma Efactura".CreateOutStream(OutStr);
                    // CopyStream(OutStr, NVInStream);
                    RecRf.Modify();
                    Rec.Get(Rec."ID");
                    Rec.CalcFields("Plantilla Excel");
                    if not rec."Plantilla Excel".HasValue Then Error('No se ha importado la plantilla excel');
                end;
            }
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
            actionref(Imprimir; Print) { }
            actionref(Importar; "Importar Plantilla") { }
            actionref("Crea_ref"; "Crear un Nuevo Informe") { }
        }
    }
    trigger OnAfterGetRecord()

    var
        "Nombre Tabla": Text;
        AllObjWithCaption: record AllObjWithCaption;
        CreateOrCopy: Option "Create a new data set","Create a copy of an existing data set","Edit an existing data set";
        "Source Service Name": Text;
        "Destination Service Name": Text;
        Columnas: Record "Columnas Informes";
    begin
        // If AllObjWithCaption.GET(AllObjWithCaption."Object Type"::Table, Rec."Tabla filtros") then
        //     "Nombre Tabla" := AllObjWithCaption."Object Caption";
        if AllObjWithCaption.GET(Rec."Tipo Objeto", Rec."Id Objeto") then
            "Nombre Query" := AllObjWithCaption."Object Caption";
        OtrosInformes := rec.Informe.AsInteger() > 1;
        Columnas.SetRange(Id, Rec."Id");
        if Columnas.Count = 0 Then begin
            If Rec."Id Objeto" <> 0 Then
                Rec.InitColumns(Rec."Tipo Objeto", Rec."Id Objeto", CreateOrCopy::"Create a new data set", "Source Service Name", "Destination Service Name");
            Rec.GetColumns(Rec);
            CurrPage.UPDATE(false);
        end;
        If Columnas.FindFirst() Then
            if AllObjWithCaption.GET(AllObjWithCaption."Object Type"::Table, columnas.Table) then
                "Nombre Tabla" := AllObjWithCaption."Object Caption";

        CampoTarea := Rec."Campo Tarea";
        WorkDescription := Rec.GetDescripcionAmpliada();
    end;

    trigger OnAfterGetCurrRecord()
    begin
        CampoTarea := Rec."Campo Tarea";
    end;


    var
        WorkDescription: Text;
        CampoTarea: Integer;
        "Nombre Tabla": Text;

        "Nombre Query": Text;
        OtrosInformes: Boolean;
        TableId: Integer;

    Procedure CargaTabla(PTableId: Integer)
    begin
        TableId := PTableId;
    end;

    procedure NombreCampo(): Text
    var

        Campos: Record Field;
    begin
        if Rec."Campo Tarea" = 0 then
            exit('');

        Campos.SetRange(tableno, TableId);
        Campos.SetRange("No.", Rec."Campo Tarea");
        If Campos.FindFirst then
            exit(Campos."Field Caption");

    end;


}
