page 7001193 "Informes Setup Wizard"
{
    Caption = 'Informes Setup';
    PageType = NavigatePage;
    SourceTable = "Informes";
    SourceTableTemporary = true;
    UsageCategory = Administration;
    AdditionalSearchTerms = 'Setup up reporting data for your own reports';

    layout
    {
        area(content)
        {
            group(Control17)
            {
                Editable = false;
                ShowCaption = false;
                Visible = TopBannerVisible AND NOT (CurrentPage = 5);
                field("MediaResourcesStandard.""Media Reference"""; MediaResourcesStandard."Media Reference")
                {
                    ApplicationArea = Basic, Suite;
                    Editable = false;
                    ShowCaption = false;
                }
            }
            group(Control19)
            {
                Editable = false;
                ShowCaption = false;
                Visible = TopBannerVisible AND (CurrentPage = 5);
                field("MediaResourcesDone.""Media Reference"""; MediaResourcesDone."Media Reference")
                {
                    ApplicationArea = Basic, Suite;
                    Editable = false;
                    ShowCaption = false;
                }
            }
            group(Step1)
            {
                Visible = CurrentPage = 1;
                group("Para1.1")
                {
                    Caption = 'Bienvenido a la configuración de Informes';
                    label("Para1.1.1_aslabel")
                    {
                        ApplicationArea = Basic, Suite;
                        Caption = 'Puede crear conjuntos de datos que puede usar para crear informes en Excel';
                    }
                    label("Whitespace1.1.1")
                    {
                        ApplicationArea = Basic, Suite;
                        Caption = '';
                    }
                    label("Para1.1.2")
                    {
                        ApplicationArea = Basic, Suite;
                        Caption = 'Esta guia está especialmente diseñada para algunas herramientas como Excel, puede seleccionar columnas o establecer filtros desde esta guía de configuración asistida. Después de completar la configuración asistida, puede usar plantillas Excel para crear o modificar informes que utilicen el nuevo sinforme, seleccionando columnas, formatos, tablas dinámicas y estableciendo los filtros según sea necesario.';
                    }
                }
                group("Para1.2")
                {
                    Caption = '¡Vamos!';

                    label("Para1.2.1_aslabel")
                    {
                        ApplicationArea = Basic, Suite;
                        Caption = 'Para comenzar, seleccione una de las siguientes opciones:';
                    }
                }
            }
            group(Step2)
            {
                Caption = '';
                Visible = CurrentPage = 2;
                group("Para2.1")
                {
                    Caption = '¿Qué le gustaría hacer?';
                    field(ActionType; ActionType)
                    {
                        ApplicationArea = Basic, Suite;
                        OptionCaption = 'Crear un nuevo informe,Crear una copia de un informe existente,Editar un informe existente';
                        ShowCaption = false;

                        trigger OnValidate()
                        begin
                            Rec."ID" := 0;
                            ClearTables();
                            ClearObjectType();
                            ClearName();
                        end;
                    }
                }
            }
            group(Step3)
            {
                InstructionalText = '';
                Visible = CurrentPage = 3;
                group("Para3.1")
                {
                    Caption = 'Elija el origen de los datos para su informe';
                    InstructionalText = '';
                    group(Control28)
                    {
                        ShowCaption = false;
                        Visible = ActionType > 0;
                        field(NameLookup; ServiceNameLookup)
                        {
                            ApplicationArea = Basic, Suite;
                            AssistEdit = false;
                            Caption = 'Nombre';
                            Lookup = true;

                            trigger OnLookup(var Text: Text): Boolean
                            var
                                TenantWebService: Record Informes;
                                TenantWebServicesLookup: Page Informes;
                            begin
                                TenantWebServicesLookup.SetTableView(TenantWebService);
                                TenantWebServicesLookup.LookupMode := true;
                                if TenantWebServicesLookup.RunModal() = ACTION::LookupOK then begin
                                    TenantWebServicesLookup.GetRecord(TenantWebService);
                                    ObjectIdLookup := TenantWebService."Id Objeto";
                                    ServiceNameLookup := TenantwebService.Descripcion;
                                    ObjectTypeLookup := TenantWebService."Tipo Objeto";

                                end;
                                Rec.Descripcion := ServiceNameLookup;
                                Rec."Tipo Objeto" := ObjectTypeLookup;
                                Rec."Id Objeto" := ObjectIdLookup;
                                ClearTables();
                            end;

                            trigger OnValidate()
                            var
                                TenantWebService: Record Informes;
                            begin
                                TenantWebService.SetRange(Descripcion, ServiceNameLookup);
                                if TenantWebService.FindFirst() then begin
                                    ObjectIdLookup := TenantWebService."Id Objeto";
                                    ServiceNameLookup := TenantwebService.Descripcion;
                                    ObjectTypeLookup := TenantWebService."Tipo Objeto";
                                    Rec.Descripcion := ServiceNameLookup;
                                    Rec."Tipo Objeto" := ObjectTypeLookup;
                                    Rec."Id Objeto" := ObjectIdLookup;
                                end else
                                    Error(UseLookupErr);
                            end;
                        }
                    }
                    group(Control32)
                    {
                        ShowCaption = false;
                        Visible = ActionType < 2;
                        field(ServiceNameEdit; ServiceNameEdit)
                        {
                            ApplicationArea = Basic, Suite;
                            Caption = 'Descripción';
                            ExtendedDatatype = None;
                            ToolTip = 'Escriba un nombre para el informe, sin espacios';

                            trigger OnValidate()
                            var

                            begin
                                Rec.Descripcion := ServiceNameEdit;

                            end;
                        }
                    }
                    field(ObjectTypeLookup; ObjectTypeLookup)
                    {
                        ApplicationArea = Basic, Suite;
                        Caption = 'Tipo de objeto origen';
                        Enabled = ActionType = 0;
                        OptionCaption = ',,,,,,,,Page,Query';

                        trigger OnValidate()
                        begin
                            rec."Tipo Objeto" := ObjectTypeLookup;
                            if rec."Tipo Objeto" <> xrec."Tipo Objeto" then begin
                                ClearTables();
                                Rec."Id Objeto" := 0;
                            end;
                        end;
                    }
                    field("Object ID"; Rec."Id Objeto")
                    {
                        ApplicationArea = Basic, Suite;
                        Caption = 'Id de origen de datos';
                        Enabled = ActionType = 0;
                        Lookup = true;

                        trigger OnLookup(var Text: Text): Boolean
                        var
                            AllObjWithCaption: Record AllObjWithCaption;
                            AllObjectsWithCaption: Page "All Objects with Caption";
                        begin
                            if ObjectTypeLookup = Rec."Tipo Objeto"::Page then begin
                                AllObjWithCaption.SetRange("Object Type", AllObjWithCaption."Object Type"::Page);
                                AllObjWithCaption.SetRange("Object Subtype", 'List')
                            end else
                                if ObjectTypeLookup = Rec."Tipo Objeto"::Query then
                                    AllObjWithCaption.SetRange("Object Type", AllObjWithCaption."Object Type"::Query);

                            AllObjectsWithCaption.SetTableView(AllObjWithCaption);

                            AllObjectsWithCaption.LookupMode := true;
                            if AllObjectsWithCaption.RunModal() = ACTION::LookupOK then begin
                                AllObjectsWithCaption.GetRecord(AllObjWithCaption);
                                if not ((AllObjWithCaption."Object Type" = AllObjWithCaption."Object Type"::Page) or
                                        (AllObjWithCaption."Object Type" = AllObjWithCaption."Object Type"::Query))
                                then
                                    Error(InvalidObjectTypeErr);
                                if (AllObjWithCaption."Object Type" = AllObjWithCaption."Object Type"::Page) and
                                   (AllObjWithCaption."Object Subtype" <> 'List')
                                then
                                    Error(InvalidPageTypeErr);
                                Rec."Id Objeto" := AllObjWithCaption."Object ID";
                                ObjectTypeLookup := AllObjWithCaption."Object Type";

                                if Rec."Id Objeto" <> xRec."Id Objeto" then
                                    ClearTables();
                                Rec.InitColumns(ObjectTypeLookup, Rec."Id Objeto", ActionType, ServiceNameLookup, ServiceNameEdit);
                                Rec.GetColumnsTemp(TempTenantWebServiceColumns);
                                CurrPage.ODataColSubForm.Page.CargarDatos(TempTenantWebServiceColumns);
                            end;
                        end;

                        trigger OnValidate()
                        var
                            AllObjWithCaption: Record AllObjWithCaption;
                        begin
                            if ObjectTypeLookup = ObjectTypeLookup::Page then begin
                                AllObjWithCaption.SetRange("Object Type", AllObjWithCaption."Object Type"::Page);
                                AllObjWithCaption.SetRange("Object ID", Rec."Id Objeto");
                                if AllObjWithCaption.FindFirst() then
                                    if AllObjWithCaption."Object Subtype" <> 'List' then
                                        Error(InvalidPageTypeErr);
                            end;

                            if Rec."Id Objeto" <> xRec."Id Objeto" then
                                ClearTables();
                            ObjectTypeLookup := AllObjWithCaption."Object Type";

                            if Rec."Id Objeto" <> xRec."Id Objeto" then
                                ClearTables();
                            Rec.InitColumns(ObjectTypeLookup, Rec."Id Objeto", ActionType, ServiceNameLookup, ServiceNameEdit);
                            Rec.GetColumnsTemp(TempTenantWebServiceColumns);
                            CurrPage.ODataColSubForm.Page.CargarDatos(TempTenantWebServiceColumns);
                        end;
                    }
                    field(ObjectName; DisplayObjectName())
                    {
                        ApplicationArea = Basic, Suite;
                        Caption = 'Nombre del objeto';
                        Enabled = false;
                    }
                    group(Cuando)
                    {
                        field("Que día Empezamos y, a que hora"; Rec."Earliest Start Date/Time")
                        {
                            ApplicationArea = Basic, Suite;
                            ToolTip = 'Specifies the earliest date and time when the job queue entry should be run. The format for the date and time must be month/day/year hour:minute, and then AM or PM. For example, 3/10/2021 12:00 AM.';
                        }
                        field("Que dia Terminamos; En blanco, sine die"; Rec."Expiration Date/Time")
                        {
                            ApplicationArea = Basic, Suite;
                            Importance = Additional;
                            ToolTip = 'Specifies the date and time when the job queue entry is to expire, after which the job queue entry will not be run.  The format for the date and time must be month/day/year hour:minute, and then AM or PM. For example, 3/10/2021 12:00 AM.';
                        }

                        field("Lo necesitamos todos los lunes"; Rec."Run on Mondays")
                        {
                            ApplicationArea = Basic, Suite;
                            ToolTip = 'Specifies that the job queue entry runs on Mondays.';
                        }
                        field("y/o los martes"; Rec."Run on Tuesdays")
                        {
                            ApplicationArea = Basic, Suite;
                            ToolTip = 'Specifies that the job queue entry runs on Tuesdays.';
                        }
                        field("y/o los miércoles"; Rec."Run on Wednesdays")
                        {
                            ApplicationArea = Basic, Suite;
                            ToolTip = 'Specifies that the job queue entry runs on Wednesdays.';
                        }
                        field("y/o los jueves"; Rec."Run on Thursdays")
                        {
                            ApplicationArea = Basic, Suite;
                            ToolTip = 'Specifies that the job queue entry runs on Thursdays.';
                        }
                        field("los viernes"; Rec."Run on Fridays")
                        {
                            ApplicationArea = Basic, Suite;
                            ToolTip = 'Specifies that the job queue entry runs on Fridays.';
                        }
                        field("Sábados"; Rec."Run on Saturdays")
                        {
                            ApplicationArea = Basic, Suite;
                            ToolTip = 'Specifies that the job queue entry runs on Saturdays.';
                        }
                        field("¿Domingos?"; Rec."Run on Sundays")
                        {
                            ApplicationArea = Basic, Suite;
                            ToolTip = 'Specifies that the job queue entry runs on Sundays.';
                        }
                        field("¿Repetimos de alguna otra forma?"; Rec.Periodicidad)
                        {
                            ApplicationArea = Basic, Suite;
                            Caption = 'Periodicidad';
                            Enabled = ActionType = 0;


                        }
                        field("¿A partir que hora lo necesita?"; Rec."Starting Time")
                        {
                            ApplicationArea = Basic, Suite;
                            Importance = Promoted;
                            ToolTip = 'Specifies the earliest time of the day that the recurring job queue entry is to be run.';
                        }
                        field("¿Hasta que hora?"; Rec."Ending Time")
                        {
                            ApplicationArea = Basic, Suite;
                            Importance = Promoted;
                            ToolTip = 'Specifies the latest time of the day that the recurring job queue entry is to be run.';
                        }
                        field("¿Cada Cuantos Minutos"; Rec."No. of Minutes between Runs")
                        {
                            ApplicationArea = Basic, Suite;
                            Importance = Promoted;
                            ToolTip = 'Specifies the minimum number of minutes that are to elapse between runs of a job queue entry. The value cannot be less than one minute. This field only has meaning if the job queue entry is set to be a recurring job. If you use a no. of minutes between runs, the date formula setting is cleared.';
                        }
                    }
                }
            }
            group(Step4)
            {
                Caption = '';
                Visible = CurrentPage = 4;
                group("Para4.1")
                {
                    Caption = 'Seleccione las columnas que desea incluir en su informe';
                    InstructionalText = 'Seleccione las columnas que desea incluir en su informe. Puede agregar filtros más adelante.';
                    part(ODataColSubForm; "Columnas Informes Temp")
                    {


                        ApplicationArea = Basic, Suite;
                        Caption = ' ';
                    }
                }
            }
            group(Step5)
            {
                Visible = CurrentPage = 5;
                group("Para5.1")
                {
                    Caption = 'Éxito!';

                    label("Para5.1.1_aslabel")
                    {
                        ApplicationArea = Basic, Suite;
                        Caption = 'El informe se ha configurado correctamente.';
                    }

                }
            }
        }
    }

    actions
    {
        area(processing)
        {
            action(AddFiltersAction)
            {
                ApplicationArea = Basic, Suite;
                Caption = '&Añadir filtros';
                Enabled = CurrentPage = 4;
                Image = "Filter";
                InFooterBar = true;
                Visible = true;

                trigger OnAction()
                var
                    ListaFiltros: Page "Lista Filtros Informes";
                    ColumnaTemp: Record "Columnas Informes" temporary;
                begin
                    Clear(ChangeFields);
                    ChangeFields := CurrPage.ODataColSubForm.PAGE.IncludeIsChanged();
                    if ChangeFields then begin
                        Clear(TempTenantWebServiceFilter);
                        TempTenantWebServiceFilter.DeleteAll();
                    end;
                    ListaFiltros.CargarDatos(TempTenantWebServiceFilter);
                    if TempTenantWebServiceColumns.table = 0 Then begin
                        Rec.InitDefaultColumns();
                        Rec.GetColumnsTemp(TempTenantWebServiceColumns);

                    end;
                    ListaFiltros.CargaTabla(TempTenantWebServiceColumns.Table);
                    ListaFiltros.RunModal();
                    ListaFiltros.DesCargarDatos(TempTenantWebServiceFilter);
                    CurrentPage := 5;

                end;
            }
            action(AddDestinatariosAction)
            {
                ApplicationArea = Basic, Suite;
                Caption = '&Añadir Destinatarios';
                Enabled = CurrentPage = 5;
                Image = PersonInCharge;
                InFooterBar = true;
                Visible = true;

                trigger OnAction()
                var
                    ListaDestinatarios: Page "Lista Destinatarios Informes";
                begin
                    Clear(ChangeFields);
                    ChangeFields := CurrPage.ODataColSubForm.PAGE.IncludeIsChanged();
                    if ChangeFields then begin
                        Clear(TempDestinatarios);
                        TempDestinatarios.DeleteAll();
                    end;
                    ListaDestinatarios.CargarDatos(TempDestinatarios);
                    ListaDestinatarios.CargaTabla(TempTenantWebServiceColumns.Table);
                    ListaDestinatarios.RunModal();
                    ListaDestinatarios.DesCargarDatos(TempDestinatarios);



                end;
            }
            action(BackAction)
            {
                ApplicationArea = Basic, Suite;
                Caption = '&Anterior';
                Enabled = (CurrentPage > 1) AND (CurrentPage < 5);
                Image = PreviousRecord;
                InFooterBar = true;

                trigger OnAction()
                begin
                    CurrentPage := CurrentPage - 1;
                    CurrPage.Update();
                end;
            }
            action(NextAction)
            {
                ApplicationArea = Basic, Suite;
                Caption = '&Siguiente';
                Enabled = (CurrentPage >= 1) AND (CurrentPage < 4);
                Image = NextRecord;
                InFooterBar = true;

                trigger OnAction()
                begin
                    case CurrentPage of
                        1, 2:
                            CurrentPage := CurrentPage + 1;
                        3:
                            begin
                                if (ActionType = ActionType::"Create a new data set") and (ServiceNameEdit = '') then
                                    Error(MissingServiceNameErr);
                                if (ActionType = ActionType::"Create a copy of an existing data set") and
                                   ((ServiceNameEdit = '') or (ServiceNameLookup = ''))
                                then
                                    Error(MissingServiceNameErr);
                                if (ActionType = ActionType::"Edit an existing data set") and (ServiceNameLookup = '') then
                                    Error(MissingServiceNameErr);
                                if Rec."Id Objeto" = 0 then
                                    Error(MissingObjectIDErr);

                                CurrPage.Update();
                                Rec.InitColumns(ObjectTypeLookup, Rec."Id Objeto", ActionType, ServiceNameLookup, ServiceNameEdit);
                                CurrentPage := CurrentPage + 1;
                            end;
                        4:
                            begin
                                if TempTenantWebServiceColumns.FindFirst() then
                                    CurrentPage := CurrentPage + 1
                                else
                                    Error(PublishWithoutFieldsErr);
                            end;
                    end;

                    CurrPage.Update(false);
                end;
            }
            // action(PublishAction)
            // {
            //     ApplicationArea = Basic, Suite;
            //     Caption = '&Publish';
            //     Enabled = CurrentPage = 4;
            //     Image = Post;
            //     InFooterBar = true;
            //     Promoted = true;

            //     trigger OnAction()
            //     var
            //         GuidedExperience: Codeunit "Guided Experience";
            //     begin
            //         if TempTenantWebServiceColumns.FindFirst() then
            //             TempTenantWebServiceColumns.DeleteAll();
            //         Commit();
            //         Page.RunModal(Page::"Lista Filtros Informes", TempTenantWebServiceFilter);
            //         CopyTempTableToConcreteTable();
            //         GuidedExperience.CompleteAssistedSetup(ObjectType::Page, Page::"Informes Setup Wizard");
            //         PublishFlag := true;
            //         CurrentPage := CurrentPage + 1;
            //         CurrPage.Update(false);
            //     end;
            // }
            action(FinishAction)
            {
                ApplicationArea = Basic, Suite;
                Caption = '&Finalizar';
                Enabled = CurrentPage = 5;
                Image = Approve;
                InFooterBar = true;

                trigger OnAction()
                begin
                    CopyTempTableToConcreteTable();

                    CurrPage.Close();
                end;
            }
            action(CreateExcelWorkBook)
            {
                ApplicationArea = Basic, Suite;
                Caption = '&Plantilla Excel';
                Enabled = (CurrentPage = 5);
                InFooterBar = true;


                trigger OnAction()
                var
                    NVInStream: InStream;
                    OutStr: OutStream;
                    TempBlob: Codeunit "Temp Blob";
                    Base64: Codeunit "Base64 Convert";
                    Base64Txt: Text;
                    RecRf: RecordRef;
                    Plantilla: Text;
                    DocAttachment: Record "Document Attachment";
                    Ids: integer;
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
                    Rec."Url Plantilla" := DocAttachment.FormBase64ToUrl(Base64Txt, 'Plantilla' + Format(Rec.Id) + '.xlsx', Ids);
                    Rec.Modify();
                end;

            }

        }
#pragma warning disable AL0788 // TODO: - Corregir en un futuro
        area(Promoted)
#pragma warning restore AL0788 // TODO: - Corregir en un futuro
        {
            actionref(AddFiltersAction_Ref; AddFiltersAction) { }
            actionref(BackAction_Ref; BackAction) { }
            actionref(NextAction_Ref; NextAction) { }
            actionref(FinishAction_Ref; FinishAction) { }
            actionref(CreateExcelWorkBook_Ref; CreateExcelWorkBook) { }
        }
    }

    trigger OnInit()
    begin
        CheckPermissions();
        LoadTopBanners();
        CurrentPage := 1;
        ObjectTypeLookup := Rec."Tipo Objeto"::Page;
        Rec."Tipo Objeto" := Rec."Tipo Objeto"::Page;
        EditInExcelVisible();
    end;

    trigger OnOpenPage()
    begin
        Rec.Insert();
    end;

    trigger OnQueryClosePage(CloseAction: Action): Boolean
    begin
        if not PublishFlag then
            if CloseAction = ACTION::OK then
                if not Confirm(NAVNotSetUpQst, false) then
                    exit(false);
    end;

    var
        MediaRepositoryStandard: Record "Media Repository";
        MediaRepositoryDone: Record "Media Repository";
        MediaResourcesStandard: Record "Media Resources";
        MediaResourcesDone: Record "Media Resources";
        TempTenantWebServiceColumns: Record "Columnas Informes" temporary;
        TempTenantWebServiceFilter: Record "Filtros Informes" temporary;
        TempDestinatarios: Record "Destinatarios Informes" temporary;
        EditinExcelFilters: Codeunit "Edit in Excel Filters";
        ClientTypeManagement: Codeunit "Client Type Management";
        CurrentPage: Integer;
        PublishFlag: Boolean;
        TopBannerVisible: Boolean;
        MissingServiceNameErr: Label 'Por favor, introduzca un nombre para el informe.';
        MissingObjectIDErr: Label 'Seleccione un objeto de origen de datos.';
        DuplicateServiceNameErr: Label 'Este nombre ya está en uso. Por favor, introduzca un nombre diferente.';
        WebServiceNameNotValidErr: Label 'El nombre del informe no es válido. Por favor, introduzca un nombre sin espacios.';
        ChangeFields: Boolean;
        ActionType: Option "Create a new data set","Create a copy of an existing data set","Edit an existing data set";
        ObjectIdLookup: Integer;
        ServiceNameLookup: Text;
        ServiceNameEdit: Text[240];
        ObjectTypeLookup: Option ,,,,,,,,"Page","Query";
        UseLookupErr: Label 'Use el botón de búsqueda para seleccionar un registro.';
        oDataUrl: Text;
        MissingFieldsErr: Label 'Please select field(s) before adding a filter.';
        InvalidPageTypeErr: Label 'La página seleccionada no es una lista. Por favor, seleccione una lista.';
        InvalidObjectTypeErr: Label 'Solo se pueden seleccionar páginas o consultas.';
        NAVNotSetUpQst: Label 'El asistente de configuración no se ha completado. ¿Está seguro de que desea cerrar la página?';
        PublishWithoutFieldsErr: Label 'Por favor, seleccione al menos una columna para publicar.';
        PermissionsErr: Label 'Usted no tiene permisos para realizar esta acción.';
        ServiceNotFoundErr: Label 'El informe no se ha encontrado.';
        ExcelVisible: Boolean;

    local procedure UpdateEditInExcelFilters(AllObjWithCaption: Record AllObjWithCaption; FilterText: Text)
    var
        ODataUtility: Codeunit ODataUtility;
        RecordRef: RecordRef;
        FieldRef: FieldRef;
        FieldIndex: Integer;
        TotalFields: Integer;
        FilterValueEnglish: Text;
        FilterEDMValue: Text;
        FilterEDMType: Enum "Edit in Excel Edm Type";
        FieldName: Text;
        FilterType: Enum "Edit in Excel Filter Type";
        PreviousGlobalLanguage: Integer;
        EnglishLanguage: Integer;
    begin
        clear(EditinExcelFilters);
        RecordRef.Open(AllObjWithCaption."Object ID");
        RecordRef.SetView(FilterText);
        TotalFields := RecordRef.FIELDCOUNT;

        for FieldIndex := 1 to TotalFields do begin
            FieldRef := RecordRef.FieldIndex(FieldIndex);

            PreviousGlobalLanguage := GlobalLanguage();
            // Retrieve filters in English-US for ease of processing
            EnglishLanguage := 1033;
            GlobalLanguage(EnglishLanguage);
            FilterValueEnglish := FieldRef.GetFilter();
            GlobalLanguage(PreviousGlobalLanguage);

            if FilterValueEnglish <> '' then
                if IsFilterRange(FieldRef) then
                    if IsFilterRangeSingleValue(FieldRef) then begin
                        FieldName := ODataUtility.ExternalizeName(FieldRef.Name);
                        FilterEDMType := ConvertFieldTypeToEdmType(FieldRef.Type);
                        FilterType := Enum::"Edit in Excel Filter Type"::Equal;
                        FilterEDMValue := ConvertToEDMValue(FilterValueEnglish, FilterEDMType);
                        EditinExcelFilters.AddField(FieldName, FilterType, FilterEDMValue, FilterEDMType);
                    end
        end
    end;

    local procedure ConvertToEDMValue(FilterValue: Text; EDMType: Enum "Edit in Excel Edm Type"): Text
    var
        FilterEDMValue: Text;
    begin
        FilterEDMValue := FilterValue;
        case EDMType of
            Enum::"Edit in Excel Edm Type"::"Edm.Boolean":
                if FilterValue = 'Yes' then
                    FilterEDMValue := 'true'
                else
                    FilterEDMValue := 'false';
        end;
        exit(FilterEDMValue)
    end;

    [TryFunction]
    local procedure IsFilterRange(FieldRef: FieldRef)
    var
        TempRange: Text;
    begin
        TempRange := FieldRef.GetRangeMax();
    end;

    local procedure IsFilterRangeSingleValue(FieldRef: FieldRef): Boolean
    begin
        exit(FieldRef.GetRangeMin() = FieldRef.GetRangeMax())
    end;

    procedure ConvertFieldTypeToEdmType(FieldType: FieldType): Enum "Edit in Excel Edm Type";
    var
        EdmType: Enum "Edit in Excel Edm Type";
    begin
        case FieldType of
            FieldType::Text, FieldType::Code, FieldType::Guid, FieldType::Option:
                EdmType := Enum::"Edit in Excel Edm Type"::"Edm.String";
            FieldType::Integer:
                EdmType := Enum::"Edit in Excel Edm Type"::"Edm.Int32";
            FieldType::BigInteger:
                EdmType := Enum::"Edit in Excel Edm Type"::"Edm.Int64";
            FieldType::Decimal:
                EdmType := Enum::"Edit in Excel Edm Type"::"Edm.Decimal";
            FieldType::DateTime, FieldType::Date:
                EdmType := Enum::"Edit in Excel Edm Type"::"Edm.DateTimeOffset";
            FieldType::Boolean:
                EdmType := Enum::"Edit in Excel Edm Type"::"Edm.Boolean";
            else
                EdmType := Enum::"Edit in Excel Edm Type"::"Edm.String";
        end;

        exit(EdmType);
    end;

    local procedure LoadTopBanners()
    begin
        if MediaRepositoryStandard.Get('AssistedSetup-NoText-400px.png', Format(ClientTypeManagement.GetCurrentClientType())) and
           MediaRepositoryDone.Get('AssistedSetupDone-NoText-400px.png', Format(ClientTypeManagement.GetCurrentClientType()))
        then
            if MediaResourcesStandard.Get(MediaRepositoryStandard."Media Resources Ref") and
               MediaResourcesDone.Get(MediaRepositoryDone."Media Resources Ref")
            then
                TopBannerVisible := MediaResourcesDone."Media Reference".HasValue;
    end;

    local procedure CopyTempTableToConcreteTable()
    var
        TenantWebServiceColumns: Record "columnas informes";
        TenantWebServiceFilter: Record "filtros informes";
        TenantWebService: Record "Informes";
        TenantWebServiceOData: Record "Destinatarios Informes";
        ODataUtility: Codeunit ODataUtility;
        SelectText: Text;
        ODataV3FilterText: Text;
        ODataV4FilterText: Text;
    begin
        TenantWebService.Init();
        TenantWebService := Rec;
        TenantWebService.Validate("Tipo Objeto", ObjectTypeLookup);
        TenantWebService.Validate("Id Objeto", Rec."Id Objeto");
        if (ActionType = ActionType::"Create a new data set") or
           (ActionType = ActionType::"Create a copy of an existing data set")
        then begin
            TenantWebService.Validate(Descripcion, ServiceNameEdit);
            TenantWebService.Insert(true)
        end else begin
            TenantWebService.Validate(Descripcion, ServiceNameLookup);
            TenantWebService.Modify(true);
        end;

        if TempTenantWebServiceColumns.FindFirst() then begin
            if ActionType = ActionType::"Edit an existing data set" then begin
                TenantWebServiceColumns.Init();
                TenantWebServiceColumns.SetRange(ID, TenantWebService.Id);
                TenantWebServiceColumns.DeleteAll();
            end;

            repeat
                TenantWebServiceColumns.TransferFields(TempTenantWebServiceColumns, true);
                TenantWebServiceColumns.ID := TenantWebService.Id;
                TenantWebServiceColumns.Insert(true);
            until TempTenantWebServiceColumns.Next() = 0;
        end;
        if TempDestinatarios.FindFirst() then begin
            if ActionType = ActionType::"Edit an existing data set" then begin
                TenantWebServiceOData.Init();
                TenantWebServiceOData.SetRange(ID, TenantWebService.Id);
                TenantWebServiceOData.DeleteAll();
            end;

            repeat
                TenantWebServiceOData.TransferFields(TempDestinatarios, true);
                TenantWebServiceOData.ID := TenantWebService.Id;
                TenantWebServiceOData.Insert(true);
            until TenantWebServiceOData.Next() = 0;
        end;

        if TempTenantWebServiceFilter.Find('-') then begin
            if ActionType = ActionType::"Edit an existing data set" then begin
                TenantWebServiceFilter.Init();
                TenantWebServiceFilter.SetRange(ID, TenantWebService.Id);
                TenantWebServiceFilter.DeleteAll();
            end;
            repeat
                TenantWebServiceFilter.TransferFields(TempTenantWebServiceFilter, true);
                TenantWebServiceFilter.ID := TenantWebService.Id;

                TenantWebServiceFilter.Insert(true);
            until TempTenantWebServiceFilter.Next() = 0;
            PublishFlag := true;
        end;
        Commit();
    end;



    local procedure ClearTables()
    begin
        CurrPage.ODataColSubForm.PAGE.DeleteColumns();
        TempTenantWebServiceColumns.DeleteAll();
        TempTenantWebServiceFilter.DeleteAll();
    end;

    local procedure ClearObjectType()
    begin
        Rec."Tipo Objeto" := Rec."Tipo Objeto"::Page;
        ObjectTypeLookup := ObjectTypeLookup::Page;
    end;

    local procedure ClearName()
    begin
        Rec.Descripcion := '';
        ServiceNameEdit := '';
        Clear(ServiceNameLookup);
    end;

    local procedure DisplayObjectName(): Text
    var
        AllObjWithCaption: Record AllObjWithCaption;
    begin
        AllObjWithCaption.SetRange("Object Type", Rec."Tipo Objeto");
        AllObjWithCaption.SetRange("Object ID", rec."Id Objeto");
        if AllObjWithCaption.FindFirst() then
            exit(AllObjWithCaption."Object Caption");
    end;

    local procedure CheckPermissions()
    var
        TenantWebService: Record "Tenant Web Service";
        TenantWebServiceOData: Record "Tenant Web Service OData";
        TenantWebServiceColumns: Record "Tenant Web Service Columns";
        TenantWebServiceFilter: Record "Tenant Web Service Filter";
    begin
        if not TenantWebService.WritePermission then
            Error(PermissionsErr);

        if not TenantWebService.ReadPermission then
            Error(PermissionsErr);

        if not TempTenantWebServiceColumns.WritePermission then
            Error(PermissionsErr);

        if not TenantWebServiceColumns.ReadPermission then
            Error(PermissionsErr);

        if not TenantWebServiceFilter.WritePermission then
            Error(PermissionsErr);

        if not TenantWebServiceFilter.ReadPermission then
            Error(PermissionsErr);

        if not TenantWebServiceOData.WritePermission then
            Error(PermissionsErr);

        if not TenantWebServiceOData.ReadPermission then
            Error(PermissionsErr);
    end;

    procedure EditInExcelVisible()
    begin
        if ClientTypeManagement.GetCurrentClientType() = CLIENTTYPE::Web then
            ExcelVisible := true;
    end;
}