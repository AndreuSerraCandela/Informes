//Crear una tabla para guardar informes
table 7001240 Informes
{
    fields
    {
        field(1; Id; Integer)
        {
            caption = 'Id';
            DataClassification = ToBeClassified;
            AutoIncrement = true;
        }
        field(2; Descripcion; Text[250])
        {
            caption = 'Descripcion';
            DataClassification = ToBeClassified;
        }
        field(3; Periodicidad; DateFormula)
        {
            caption = 'Periodicidad';
            DataClassification = ToBeClassified;
            trigger OnValidate()
            begin
                Clear("No. of Minutes between Runs");
                ClearRunOnWeekdays();
                SetRecurringField();
                if IsNextRunDateFormulaSet() and ("Earliest Start Date/Time" = 0DT) then
                    "Earliest Start Date/Time" := CalcNextRunTimeForRecurringReport(Rec, CurrentDateTime);

            end;
        }
        field(44; "Fecha Primera Ejecución"; Date)
        {
            ObsoleteState = Removed;
            caption = 'Fecha 1ª Ejecución';
            DataClassification = ToBeClassified;
            trigger OnValidate()
            begin
                SetRecurringField();

            end;
        }
        field(4; "Earliest Start Date/Time"; DateTime)
        {
            Caption = 'Fecha próxima Ejecución';

            trigger OnLookup()
            begin
                Validate("Earliest Start Date/Time", LookupDateTime("Earliest Start Date/Time", 0DT, "Expiration Date/Time"));

            end;

            trigger OnValidate()
            begin
                CheckStartAndExpirationDateTime();
                "Fecha Próx. Ejecución" := DT2Date("Earliest Start Date/Time");
            end;
        }
        field(8; "Expiration Date/Time"; DateTime)
        {
            Caption = 'Fecha/Hora de Expiración';

            trigger OnLookup()
            begin
                Validate("Expiration Date/Time", LookupDateTime("Expiration Date/Time", "Earliest Start Date/Time", 0DT));
            end;

            trigger OnValidate()
            begin
                CheckStartAndExpirationDateTime();
            end;
        }
        field(5; Informe; Enum Informes)
        {
            caption = 'Informe';
            DataClassification = ToBeClassified;
            // trigger OnValidate()
            // begin
            //     Case Informe of
            //         Informes::"Contratos x Empresa":
            //             Validate("Tabla filtros", 36);
            //         Informes::"Estadisticas Contabilidad":
            //             Validate("Tabla filtros", 17);
            //     End;
            // end;
        }
        field(6; "Tabla filtros"; Integer)
        {
            ObsoleteState = Removed;
            caption = 'Tabla para filtros';
            TableRelation = AllObjWithCaption."Object ID" WHERE("Object Type" = CONST(Table));
        }
        field(7; "Próxima Fecha"; Date)
        {
            ObsoleteState = Removed;
            caption = 'Fecha próxima Ejecución';
            DataClassification = ToBeClassified;
        }
        field(9; "Plantilla Excel"; Blob)
        {
            caption = 'Plantilla Excel';
            DataClassification = ToBeClassified;
        }
        field(10; "Crear Tarea"; Boolean)
        {
            caption = 'Crear Tarea';
            DataClassification = ToBeClassified;
        }
        field(11; "Descripcion Tarea"; Text[250])
        {

        }
        field(12; "Campo Tarea"; Integer)
        {
            caption = 'Informe';
            DataClassification = ToBeClassified;
        }
        field(13; "Id Objeto"; Integer)
        {
            TableRelation = AllObjWithCaption."Object ID" WHERE("Object Type" = field("Tipo Objeto"));
        }
        field(14; "Tipo Objeto"; Option)
        {
            OptionMembers = ,,,,,,,,"Page","Query";
            OptionCaption = ',,,,,,,,Página,Consulta';
            caption = 'Tipo Objeto';
            DataClassification = ToBeClassified;
        }
        field(15; "Crear Web Service"; Boolean)
        {
            caption = 'Crear Web Service';
            DataClassification = ToBeClassified;
            trigger OnValidate()
            var
                Creaweb: Codeunit "Web Service Management";

            begin
                if "Crear Web Service" then
                    Creaweb.CreateWebService(Rec."Tipo Objeto", Rec."Id Objeto", Rec.Descripcion, true);
                Rec.Informe := Informes::"Web Service";
            end;
        }
        field(19; "Run on Mondays"; Boolean)
        {
            Caption = 'Ejecutar los Lunes';

            trigger OnValidate()
            begin
                Clear(Periodicidad);
                SetRecurringField();

            end;
        }
        field(20; "Run on Tuesdays"; Boolean)
        {
            Caption = 'Ejecutar los Martes';

            trigger OnValidate()
            begin
                Clear(Periodicidad);
                SetRecurringField();
            end;
        }
        field(21; "Run on Wednesdays"; Boolean)
        {
            Caption = 'Ejecutar los Miércoles';

            trigger OnValidate()
            begin
                Clear(Periodicidad);
                SetRecurringField();
            end;
        }
        field(22; "Run on Thursdays"; Boolean)
        {
            Caption = 'Ejecutar los Jueves';

            trigger OnValidate()
            begin
                Clear(Periodicidad);
                SetRecurringField();
            end;
        }
        field(23; "Run on Fridays"; Boolean)
        {
            Caption = 'Ejecutar los Viernes';

            trigger OnValidate()
            begin
                Clear(Periodicidad);
                SetRecurringField();
            end;
        }
        field(24; "Run on Saturdays"; Boolean)
        {
            Caption = 'Ejecutar los Sábados';

            trigger OnValidate()
            begin
                Clear(Periodicidad);
                SetRecurringField();
            end;
        }
        field(25; "Run on Sundays"; Boolean)
        {
            Caption = 'Ejecutar los Domingos';

            trigger OnValidate()
            begin
                Clear(Periodicidad);
                SetRecurringField();
            end;
        }
        field(26; "Starting Time"; Time)
        {
            Caption = 'Hora de Inicio';

            trigger OnValidate()
            begin
                if "Starting Time" = 0T then
                    "Reference Starting Time" := 0DT
                else
                    "Reference Starting Time" := CreateDateTime(DMY2Date(1, 1, 2000), "Starting Time");
            end;
        }
        field(27; "Ending Time"; Time)
        {
            Caption = 'Hora de Fin';


        }
        field(28; "Reference Starting Time"; DateTime)
        {
            Caption = 'Hora de Inicio de Referencia';
            Editable = false;

            trigger OnValidate()
            begin
                "Starting Time" := DT2Time("Reference Starting Time");
            end;
        }
        field(18; "No. of Minutes between Runs"; Integer)
        {
            Caption = 'Minutos entre ejecuciones';

            trigger OnValidate()
            begin
                If ("No. of Minutes between Runs" < 15) and ("No. of Minutes between Runs" <> 0) then
                    Error('El tiempo entre ejecuciones debe ser mayor a 15 minutos');
                Clear(Periodicidad);
                SetRecurringField();
                SetMinimumNumberOfMinutesBetweenRuns();
            end;
        }
        field(200; "Descripción Ampliada"; BLOB)
        {

        }
        field(29; Bcc; Text[250])
        {
            Caption = 'CCO';
            DataClassification = ToBeClassified;
        }
        field(36; "Solo Bcc"; Boolean)
        {
            Caption = 'Solo CCO';
            DataClassification = ToBeClassified;
        }
        field(30; Ejecutandose; Boolean)
        {

        }
        field(31; Orientación; Enum Orientacion)
        {
            Caption = 'Orientación';
            DataClassification = ToBeClassified;
        }
        field(32; "Fecha Últ. Ejecución"; Date)
        {
            Caption = 'Fecha Última Ejecución';
            DataClassification = ToBeClassified;
        }
        field(33; "Fecha Próx. Ejecución"; Date)
        {
            Caption = 'Fecha Último Ejecución';
            DataClassification = ToBeClassified;
        }
        field(34; "Formato Json"; Boolean)
        {
            Caption = 'Formato Json';
            DataClassification = ToBeClassified;
        }
        field(35; "Url Plantilla"; Text[250])
        {

            DataClassification = ToBeClassified;
        }



    }
    keys
    {
        key(PK; Id)
        {
            Clustered = true;
        }
    }
    [Scope('OnPrem')]
    procedure InitColumns(ObjectType: Option ,,,,,,,,"Page","Query"; ObjectID: Integer; InActionType: Option "Create a new data set","Create a copy of an existing data set","Edit an existing data set"; InSourceServiceName: Text; DestinationServiceName: Text)
    var
        AllObj: Record AllObj;
        ApplicationObjectMetadata: Record "Application Object Metadata";
        inStream: InStream;
        Columnas: Record "Columnas Informes";
    begin
        Columnas.SetRange(Id, Rec."Id");
        //If Columnas.FindFirst() then exit;
        OdataColumnChose.InitColumns(ObjectType, ObjectID, InActionType, InSourceServiceName, DestinationServiceName);

    end;

    [Scope('OnPrem')]
    procedure InitDefaultColumns()
    var
        CreateOrCopy: Option "Create a new data set","Create a copy of an existing data set","Edit an existing data set";
        "Source Service Name": Text;
        "Destination Service Name": Text;

    begin
        InitColumns(Rec."Tipo Objeto", Rec."Id Objeto", CreateOrCopy::"Create a new data set", "Source Service Name", "Destination Service Name");

    end;


    procedure GetColumnsTemp(var Columnas: Record "Columnas Informes" temporary)
    var
        TempTenantWebServiceColumns: Record "Tenant Web Service Columns" temporary;

        Field: Record Field;
        Orden: Integer;
        RecRef: RecordRef;
        Fieldref: FieldRef;
        a: Integer;
        Existe: Boolean;
        Tabla: Integer;
        TablaCamposPagina: record "Page Table Field";
    begin
        Columnas.DeleteAll();
        OdataColumnChose.GetColumns(TempTenantWebServiceColumns);
        if TempTenantWebServiceColumns.FindFirst() then
            repeat
                Columnas.Init();
                Columnas.Include := TempTenantWebServiceColumns.Include;
                Columnas."Field Name" := TempTenantWebServiceColumns."Field Name";
                Columnas.Table := TempTenantWebServiceColumns."Data Item";
                If TempTenantWebServiceColumns."Field Caption" = '' then begin
                    RecRef.Open(TempTenantWebServiceColumns."Data Item");
                    Tabla := TempTenantWebServiceColumns."Data Item";
                    Existe := RecRef.FieldExist(TempTenantWebServiceColumns."Field Number");
                    if Existe then begin
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
                if Existe then begin
                    while not Columnas.Insert() do begin
                        a += 1;
                        Columnas.Id_campo := a;
                    end;
                end;

            until TempTenantWebServiceColumns.Next() = 0;
        If Rec."Tipo Objeto" = Rec."Tipo Objeto"::Page Then begin
            TablaCamposPagina.SetRange("Page ID", Rec."Id Objeto");
            If TablaCamposPagina.FindFirst() then
                repeat
                    Columnas.SetRange(Campo, TablaCamposPagina."Field ID");
                    if not Columnas.FindFirst() then begin
                        Columnas.Init();
                        Columnas.Campo := TablaCamposPagina."Field ID";
                        Columnas.Titulo := TablaCamposPagina.Caption;
                        //TODO: Verificar si es campo de tabla
                        //if TablaCamposPagina.FieldKind=TablaCamposPagina.FieldKind::TableField then begin
#pragma warning disable AL0432
                        if TablaCamposPagina.IsTableField then begin
#pragma warning restore AL0432
                            Field.Get(TempTenantWebServiceColumns."Data Item", TablaCamposPagina."Field ID");
                            Columnas."Field Name" := Field.FieldName;
                        end;
                        while not Columnas.Insert() do begin
                            a += 1;
                            Columnas.Id_campo := a;
                        end;
                    end;

                until TablaCamposPagina.Next() = 0;
            Field.SetRange(tableno, Tabla);
            If Field.FindFirst() then
                repeat
                    Columnas.SetRange(Campo, Field."No.");
                    if not Columnas.FindFirst() then begin
                        Columnas.Init();
                        Columnas.Campo := Field."No.";
                        Columnas.Titulo := Field."Field Caption";
                        Columnas."Field Name" := Field."FieldName";
                        while not Columnas.Insert() do begin
                            a += 1;
                            Columnas.Id_campo := a;
                        end;
                    end;

                until Field.Next() = 0;
        end;
        Columnas.Reset();
    end;


    procedure GetColumns(var Informe: Record Informes)
    var
        TempTenantWebServiceColumns: Record "Tenant Web Service Columns" temporary;
        Columnas: Record "Columnas Informes";
        Field: Record Field;
        Orden: Integer;
        RecRef: RecordRef;
        Fieldref: FieldRef;
        a: Integer;
        Tabla: Integer;
        Existe: Boolean;
        TablaCamposPagina: record "Page Table Field";
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
                Columnas.Table := TempTenantWebServiceColumns."Data Item";
                if Columnas.Table <> 0 then
                    Tabla := Columnas.Table;
                If TempTenantWebServiceColumns."Field Caption" = '' then begin
                    RecRef.Open(TempTenantWebServiceColumns."Data Item");
                    Existe := RecRef.FieldExist(TempTenantWebServiceColumns."Field Number");
                    if Existe then begin
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
                //if Existe then begin
                while not Columnas.Insert() do begin
                    a += 1;
                    Columnas.Id_campo := a;
                end;


            until TempTenantWebServiceColumns.Next() = 0;
        If Informe."Tipo Objeto" = Informe."Tipo Objeto"::Page Then begin
            TablaCamposPagina.SetRange("Page ID", Informe."Id Objeto");
            If TablaCamposPagina.FindFirst() then
                repeat
                    Columnas.SetRange(Id, Informe."Id");
                    Columnas.SetRange(Campo, TablaCamposPagina."Field ID");
                    if not Columnas.FindFirst() then begin
                        Columnas.Init();
                        Columnas.Id := Informe."Id";
                        Columnas.Campo := TablaCamposPagina."Field ID";
                        Columnas.Titulo := TablaCamposPagina.Caption;
                        //TODO: Verificar si es campo de tabla
                        //if TablaCamposPagina.FieldKind=TablaCamposPagina.FieldKind::TableField then begin
#pragma warning disable AL0432
                        if TablaCamposPagina.IsTableField then begin
#pragma warning restore AL0432
                            Field.Get(TempTenantWebServiceColumns."Data Item", TablaCamposPagina."Field ID");
                            Columnas."Field Name" := Field.FieldName;
                        end;
                        while not Columnas.Insert() do begin
                            a += 1;
                            Columnas.Id_campo := a;
                        end;
                    end;

                until TablaCamposPagina.Next() = 0;


        end;

    end;

    local procedure SetRecurringField()
    var

    begin
        If "Earliest Start Date/Time" = 0DT then exit;
        //procedure SetRecurringField()
        SetMinimumNumberOfMinutesBetweenRuns();
    end;

    local procedure SetMinimumNumberOfMinutesBetweenRuns()
    begin
        if not IsNextRunDateFormulaSet() and (Rec."No. of Minutes between Runs" = 0) then
            Rec."No. of Minutes between Runs" := 1440; // Default to one day
                                                       // "Próxima Fecha" := CalcDate(Periodicidad, "Earliest Start Date/Time");
    end;

    procedure IsNextRunDateFormulaSet(): Boolean
    begin
        exit(Format("Periodicidad") <> '');
    end;

    local procedure ClearRunOnWeekdays()
    begin
        "Run on Fridays" := false;
        "Run on Mondays" := false;
        "Run on Saturdays" := false;
        "Run on Sundays" := false;
        "Run on Thursdays" := false;
        "Run on Tuesdays" := false;
        "Run on Wednesdays" := false;
    end;

    procedure CalcNextRunTimeForRecurringReport(var JobQueueEntry: Record "Informes"; StartingDateTime: DateTime) NewRunDateTime: DateTime
    var
        RunOnDate: array[7] of Boolean;
        StartingWeekDay: Integer;
        NoOfExtraDays: Integer;
        NoOfDays: Integer;
        Found: Boolean;
        ACro: Boolean;
    begin
        with JobQueueEntry do begin
            //TestField("Recurring Job");
            RunOnDate[1] := "Run on Mondays";
            RunOnDate[2] := "Run on Tuesdays";
            RunOnDate[3] := "Run on Wednesdays";
            RunOnDate[4] := "Run on Thursdays";
            RunOnDate[5] := "Run on Fridays";
            RunOnDate[6] := "Run on Saturdays";
            RunOnDate[7] := "Run on Sundays";
            NewRunDateTime := StartingDateTime;
            if "Ending Time" = 0T then begin
                ACro := true;
                "Ending Time" := "Starting Time";
            end;

            NoOfDays := 0;
            if ("Ending Time" <> 0T) and (NewRunDateTime >= GetEndingDateTime(NewRunDateTime)) then begin
                NewRunDateTime := GetStartingDateTime(NewRunDateTime);
                NoOfDays := NoOfDays + 1;
            end;

            StartingWeekDay := Date2DWY(DT2Date(StartingDateTime), 1);
            Found := RunOnDate[(StartingWeekDay - 1 + NoOfDays) mod 7 + 1];
            NoOfExtraDays := 0;
            while not Found and (NoOfExtraDays < 7) do begin
                NoOfExtraDays := NoOfExtraDays + 1;
                NoOfDays := NoOfDays + 1;
                Found := RunOnDate[(StartingWeekDay - 1 + NoOfDays) mod 7 + 1];
            end;

            if ("Starting Time" <> 0T) and (NewRunDateTime < GetStartingDateTime(NewRunDateTime)) then
                NewRunDateTime := GetStartingDateTime(NewRunDateTime);

            if (NoOfDays > 0) and (NewRunDateTime > GetStartingDateTime(NewRunDateTime)) then
                NewRunDateTime := GetStartingDateTime(NewRunDateTime);

            if ("Starting Time" = 0T) and (NoOfExtraDays > 0) and ("No. of Minutes between Runs" <> 0) then
                NewRunDateTime := CreateDateTime(DT2Date(NewRunDateTime), 0T);

            if Found then
                NewRunDateTime := CreateDateTime(DT2Date(NewRunDateTime) + NoOfDays, DT2Time(NewRunDateTime));
            If ACro then
                "Ending Time" := 0T;
        end;


    end;



    procedure AddMinutesToDateTime(SourceDateTime: DateTime; NoOfMinutes: Integer) NewDateTime: DateTime
    var
        MillisecondsToAdd: BigInteger;
    begin
        MillisecondsToAdd := NoOfMinutes;
        MillisecondsToAdd := MillisecondsToAdd * 60000;
        NewDateTime := SourceDateTime + MillisecondsToAdd;
    end;

    local procedure CalcRunTimeForRecurringReport(var Informes: Record Informes; StartingDateTime: DateTime) NewRunDateTime: DateTime
    var
        RunOnDate: array[7] of Boolean;
        StartingWeekDay: Integer;
        NoOfExtraDays: Integer;
        NoOfDays: Integer;
        Found: Boolean;
    begin
        with Informes do begin
            RunOnDate[1] := "Run on Mondays";
            RunOnDate[2] := "Run on Tuesdays";
            RunOnDate[3] := "Run on Wednesdays";
            RunOnDate[4] := "Run on Thursdays";
            RunOnDate[5] := "Run on Fridays";
            RunOnDate[6] := "Run on Saturdays";
            RunOnDate[7] := "Run on Sundays";
            NewRunDateTime := StartingDateTime;
            NoOfDays := 0;
            if ("Ending Time" <> 0T) and (NewRunDateTime > GetEndingDateTime(NewRunDateTime)) then begin
                NewRunDateTime := GetStartingDateTime(NewRunDateTime);
                NoOfDays := NoOfDays + 1;
            end;

            StartingWeekDay := Date2DWY(DT2Date(StartingDateTime), 1);
            Found := RunOnDate[(StartingWeekDay - 1 + NoOfDays) mod 7 + 1];
            NoOfExtraDays := 0;
            while not Found and (NoOfExtraDays < 7) do begin
                NoOfExtraDays := NoOfExtraDays + 1;
                NoOfDays := NoOfDays + 1;
                Found := RunOnDate[(StartingWeekDay - 1 + NoOfDays) mod 7 + 1];
            end;

            if ("Starting Time" <> 0T) and (NewRunDateTime < GetStartingDateTime(NewRunDateTime)) then
                NewRunDateTime := GetStartingDateTime(NewRunDateTime);

            if (NoOfDays > 0) and (NewRunDateTime > GetStartingDateTime(NewRunDateTime)) then
                NewRunDateTime := GetStartingDateTime(NewRunDateTime);

            if ("Starting Time" = 0T) and (NoOfExtraDays > 0) and ("No. of Minutes between Runs" <> 0) then
                NewRunDateTime := CreateDateTime(DT2Date(NewRunDateTime), 0T);

            if Found then
                NewRunDateTime := CreateDateTime(DT2Date(NewRunDateTime) + NoOfDays, DT2Time(NewRunDateTime));
        end;


    end;

    procedure GetStartingDateTime(Date: DateTime): DateTime
    begin
        if "Reference Starting Time" = 0DT then
            Validate("Starting Time");
        exit(CreateDateTime(DT2Date(Date), DT2Time("Reference Starting Time")));
    end;

    procedure GetEndingDateTime(Date: DateTime): DateTime
    begin
        if "Reference Starting Time" = 0DT then
            Validate("Starting Time");
        if "Ending Time" = 0T then
            exit(CreateDateTime(DT2Date(Date), 0T));
        if "Starting Time" = 0T then
            exit(CreateDateTime(DT2Date(Date), "Ending Time"));
        if "Starting Time" < "Ending Time" then
            exit(CreateDateTime(DT2Date(Date), "Ending Time"));
        if "Starting Time" = "Ending Time" then
            exit(CreateDateTime(DT2Date(Date), "Ending Time"));
        exit(CreateDateTime(DT2Date(Date) + 1, "Ending Time"));
    end;

    local procedure LookupDateTime(InitDateTime: DateTime; EarliestDateTime: DateTime; LatestDateTime: DateTime): DateTime
    var
        DateTimeDialog: Page "Date-Time Dialog";
        NewDateTime: DateTime;
    begin
        NewDateTime := InitDateTime;
        if InitDateTime < EarliestDateTime then
            InitDateTime := EarliestDateTime;
        if (LatestDateTime <> 0DT) and (InitDateTime > LatestDateTime) then
            InitDateTime := LatestDateTime;

        DateTimeDialog.SetDateTime(RoundDateTime(InitDateTime, 1000));

        if DateTimeDialog.RunModal() = ACTION::OK then
            NewDateTime := DateTimeDialog.GetDateTime();
        exit(NewDateTime);
    end;

    local procedure CheckStartAndExpirationDateTime()
    begin
        if IsExpired("Earliest Start Date/Time") then
            Error(ExpiresBeforeStartErr, FieldCaption("Expiration Date/Time"), FieldCaption("Earliest Start Date/Time"));
    end;

    procedure IsExpired(AtDateTime: DateTime): Boolean
    begin
        exit((AtDateTime <> 0DT) and ("Expiration Date/Time" <> 0DT) and ("Expiration Date/Time" < AtDateTime));
    end;

    procedure SetDescripcionAmpliada(NewDescripcionAmpliada: Text)
    var
        OutStream: OutStream;
    begin
        Clear("Descripción Ampliada");
        "Descripción Ampliada".CreateOutStream(OutStream, TEXTENCODING::UTF8);
        OutStream.WriteText(NewDescripcionAmpliada);
        Modify();
    end;

    procedure GetDescripcionAmpliada() DescripcionAmliada: Text
    var
        TypeHelper: Codeunit "Type Helper";
        InStream: InStream;
    begin
        CalcFields("Descripción Ampliada");
        "Descripción Ampliada".CreateInStream(InStream, TEXTENCODING::UTF8);
        exit(TypeHelper.TryReadAsTextWithSepAndFieldErrMsg(InStream, TypeHelper.LFSeparator(), FieldName("Descripción Ampliada")));
    end;

    var
        ExpiresBeforeStartErr: Label '%1 debe ser posterior a %2.', Comment = '%1 = Expiration Date, %2=Start date';
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
//crear page para informes