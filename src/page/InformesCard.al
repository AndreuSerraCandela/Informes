
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
                    begin
                        AllObjWithCaption.GET(Rec."Tipo Objeto", Rec."Id Objeto");
                        "Nombre Query" := AllObjWithCaption."Object Caption";

                    end;
                }
                field("Nombre Objeto"; "Nombre Query")
                {
                    ApplicationArea = All;
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
            }
            part(Destinatarios; "Destinatarios Informes")
            {
                ApplicationArea = All;
                SubPageLink = Id = fIELD(Id);

            }
            part(Campos; "Campos Informes")
            {
                ApplicationArea = All;
                SubPageLink = Id = fIELD(Id);
            }
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
                    Informes.imprimirInformes(Rec.Id, 0D);// Código para imprimir informe
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
        }
        area(Promoted)
        {
            actionref(Imprimir; Print) { }
            actionref(Importar; "Importar Plantilla") { }
        }
    }
    trigger OnAfterGetRecord()

    var
        "Nombre Tabla": Text;
        AllObjWithCaption: record AllObjWithCaption;
        CreateOrCopy: Option "Create a new data set","Create a copy of an existing data set","Edit an existing data set";
        "Source Service Name": Text;
        "Destination Service Name": Text;
    begin
        If AllObjWithCaption.GET(AllObjWithCaption."Object Type"::Table, Rec."Tabla filtros") then
            "Nombre Tabla" := AllObjWithCaption."Object Caption";
        if AllObjWithCaption.GET(Rec."Tipo Objeto", Rec."Id Objeto") then
            "Nombre Query" := AllObjWithCaption."Object Caption";
        OtrosInformes := rec.Informe.AsInteger() > 1;
        If Rec."Id Objeto" <> 0 Then
            InitColumns(Rec."Tipo Objeto", Rec."Id Objeto", CreateOrCopy::"Create a new data set", "Source Service Name", "Destination Service Name");
        GetColumns(Rec);
    end;

    [Scope('OnPrem')]
    procedure InitColumns(ObjectType: Option ,,,,,,,,"Page","Query"; ObjectID: Integer; InActionType: Option "Create a new data set","Create a copy of an existing data set","Edit an existing data set"; InSourceServiceName: Text; DestinationServiceName: Text)
    var
        AllObj: Record AllObj;
        ApplicationObjectMetadata: Record "Application Object Metadata";
        inStream: InStream;
        Columnas: Record "Columnas Informes";
    begin
        Columnas.SetRange(Id, Rec."Id");
        If Columnas.FindFirst() then exit;
        OdataColumnChose.InitColumns(ObjectType, ObjectID, InActionType, InSourceServiceName, DestinationServiceName);

    end;

    procedure GetColumns(var Informe: Record Informes)
    var
        TempTenantWebServiceColumns: Record "Tenant Web Service Columns" temporary;
        Columnas: Record "Columnas Informes";
        Orden: Integer;
        RecRef: RecordRef;
        Fieldref: FieldRef;
        a: Integer;
    begin
        OdataColumnChose.GetColumns(TempTenantWebServiceColumns);
        Columnas.SetRange(Id, Informe."Id");
        Columnas.DeleteAll();
        if TempTenantWebServiceColumns.FindFirst() then
            repeat
                Columnas.Init();
                Columnas.Id := Informe."Id";
                Columnas.Include := TempTenantWebServiceColumns.Include;
                Columnas."Field Name" := TempTenantWebServiceColumns."Field Name";
                If TempTenantWebServiceColumns."Field Caption" = '' then begin
                    RecRef.Open(TempTenantWebServiceColumns."Data Item");
                    if RecRef.FieldExist(TempTenantWebServiceColumns."Field Number") then begin
                        Fieldref := RecRef.FIELD(TempTenantWebServiceColumns."Field Number");
                        Columnas.Titulo := Fieldref.CAPTION;
                    end else
                        Columnas.Titulo := TempTenantWebServiceColumns."Field Name";
                    RecRef.Close();
                end else
                    Columnas.Titulo := TempTenantWebServiceColumns."Field Caption";
                Orden += 1;
                Columnas.Orden := orden;
                Columnas.Campo := TempTenantWebServiceColumns."Field Number";
                while not Columnas.Insert() do begin
                    a += 1;
                    Columnas.Id_campo := a;
                end;

            until TempTenantWebServiceColumns.Next() = 0;
        Reset();
    end;



    var
        "Nombre Tabla": Text;

        "Nombre Query": Text;
        OtrosInformes: Boolean;

        OdataColumnChose: Page "OData Column Choose SubForm";
        TenantColumns: Record "Tenant Web Service Columns";
        SourceObjectType: Option ,,,,,,,,"Page","Query";
        ActionType: Option "Create a new data set","Create a copy of an existing data set","Edit an existing data set";
        SourceServiceName: Text;
        SourceObjectID: Integer;
        IsModified: Boolean;
        CheckFieldErr: Label 'You cannot exclude field from selection because of applied filter for it.';
        AskYourSystemAdministratorToSetupErr: Label 'Cannot complete this task. Ask your administrator for assistance.';
        CalledForExcelExport: Boolean;

}
