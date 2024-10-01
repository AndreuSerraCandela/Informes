//Crear una tabla para guardar informes
table 7001244 "Columnas Informes"
{
    LookupPageId = "Lista Columnas";
    DrillDownPageId = "Columnas Informes";
    fields
    {
        field(1; Id; Integer)
        {
            caption = 'Id';
            DataClassification = ToBeClassified;
        }
        field(2; Campo; Integer)
        {
            caption = 'Campo';
            trigger OnLookup()
            var
                Informes: Record "Informes";
                Campos: Record Field;
                Columna: Record "Columnas Informes";
                Columnat: Record "Columnas Informes" temporary;
                Preguntar: Boolean;
                ControlInformes: Codeunit ControlInformes;
            begin
                if Funcion = Funcion::Columna Then exit;
                Informes.Get(Id);
                Preguntar := Informes.Informe = Informes.Informe::"Web Service";
                If Preguntar Then Preguntar := Confirm('¿Desea seleccionar un campo del web dervice?');
                if Preguntar = False then begin
                    Columna.SetRange(Columna."Id", "Id");
                    If not Columna.FindFirst() then begin
                        Informes.InitDefaultColumns();
                        Informes.GetColumns(Informes);
                        Commit();
                        Columna.FindFirst();
                    end;
                    Campos.SetRange(Campos.TableNo, Columna.Table);
                    If Page.Runmodal(9806, Campos) = Action::LookupOK then
                        Campo := Campos."No.";
                end else begin
                    InitDefaultColumns(Informes);
                    GetColumnsTemp(Columnat);
                    Commit();
                    If Page.Runmodal(Page::"Lista Columnas", Columnat) = Action::LookupOK then begin
                        Campo := Columnat.Campo;
                        Titulo := Columnat.Titulo;
                        "Field Name" := Columnat."Field Name";
                    end;
                end;
                If Informes.Get(Id) then begin
                    If Informes.Informe = Informes.Informe::"Web Service" then begin
                        "Field Name" := ControlInformes.ExternalizeName("Field Name");
                    end;
                end;
            end;

            trigger OnValidate()
            var
                Informes: Record "Informes";
                Campos: Record Field;
                Columna: Record "Columnas Informes";
                ControlInformes: Codeunit "ControlInformes";

            begin
                If Funcion = Funcion::Columna Then exit;
                Columna.SetRange(Columna."Id", "Id");
                If not Columna.FindFirst() then begin
                    Informes.InitDefaultColumns();
                    Informes.GetColumns(Informes);
                    Commit();
                    Columna.FindFirst();
                end;
                Campos.SetRange(Campos.TableNo, Columna.Table);
                Campos.SetRange(Campos."No.", Campo);
                Campos.FindFirst();
                Titulo := Campos."Field Caption";
                "Field Name" := Campos.FieldName;
                If Informes.Get(Id) then begin
                    If Informes.Informe = Informes.Informe::"Web Service" then begin
                        "Field Name" := ControlInformes.ExternalizeName("Field Name");
                    end;
                end;
            end;


        }
        field(3; Titulo; Text[250])
        {
            caption = 'Titulo';
            DataClassification = ToBeClassified;
            trigger OnValidate()
            var
                Self: Record "Columnas Informes";
                a: Integer;
            begin
                if Titulo <> '' then begin
                    if Orden = 0 then begin
                        Self.SetRange(Self."Id", "Id");
                        Self.SetRange(Include, true);
                        if Self.Findset then
                            repeat
                                if Self.Orden > a then a := Self.Orden;
                            until Self.Next = 0;
                        Validate(Orden, a + 1);
                    end;
                    Include := true;

                end;
            end;
        }
        field(4; Orden; Integer)
        {
            caption = 'Orden';
            DataClassification = ToBeClassified;
            trigger OnValidate()
            var
                Self: Record "Columnas Informes";
            begin
                Letra := '';
                Mostrar := Mostrar::"No";
                if Orden = 0 then
                    Include := false
                else begin
                    Include := true;
                    Mostrar := Mostrar::" ";
                    Letra := LetraColumna(Orden);
                    Self.SetRange(Self."Id", "Id");
                    Self.SetRange(Self.Orden, Orden);
                    Self.SetFilter(Id_campo, '<>%1', Id_campo);
                    if Self.FindFirst() then
                        error('El orden ya existe en la lista de columnas');
                end;
            end;
        }
        field(5; Funcion; Enum Funciones)
        {
            caption = 'Funcion';
            DataClassification = ToBeClassified;

        }
        field(12; Letra; Code[2])
        {
            caption = 'Letra';
            DataClassification = ToBeClassified;
            trigger OnValidate()
            var
                x: Integer;
                i: Integer;
                C: Integer;
            begin
                Letra := UpperCase(Letra);
                for i := 1 to strlen(Letra) do begin
                    x := Letra[i];
                    if (x < 65) or (x > 90) then
                        error('Solo se permiten letras mayúsculas de la A a la Z');
                    C := C + x - 64 + (25 * (i - 1));
                end;
                Validate(Orden, C);

            end;
        }
        field(6; Include; Boolean)
        {
            caption = 'Incluir';
            DataClassification = ToBeClassified;
            trigger OnValidate()
            var
                Self: Record "Columnas Informes";
                a: Integer;
            begin
                if not Include then
                    Validate(Orden, 0)
                else begin
                    if Orden = 0 then begin
                        Self.SetRange(Self."Id", "Id");
                        Self.SetRange(Include, true);
                        if Self.Findset then
                            repeat
                                if Self.Orden > a then a := Self.Orden;
                            until Self.Next = 0;
                        Validate(Orden, a + 1);
                    end;
                end;

            end;
        }
        field(7; "Field Name"; Text[250])
        {
            caption = 'Field Name';
            DataClassification = ToBeClassified;
        }
        field(8; "Id_campo"; Integer)
        {
            caption = 'Id_campo';
            AutoIncrement = true;
        }
        field(9; Table; Integer)
        {
            caption = 'Table';
            DataClassification = ToBeClassified;
        }
        field(10; "Ancho Columna"; Integer)
        {
            caption = 'Ancho Columna';
            DataClassification = ToBeClassified;
        }
        field(11; "Formato Columna"; Text[250])
        {
            caption = 'Formato Columna';
            DataClassification = ToBeClassified;
        }
        field(13; Mostrar; Option)
        {
            OptionMembers = " ","No";
        }






    }
    keys
    {
        key(PK; Id, Id_Campo)
        {
            Clustered = true;
        }
        key(Mostrar; Mostrar, Orden)
        {

        }


    }
    procedure GetColumnsTemp(var Columnas: Record "Columnas Informes" temporary)
    var
        TempTenantWebServiceColumns: Record "Tenant Web Service Columns" temporary;

        Field: Record Field;
        Orden: Integer;
        RecRef: RecordRef;
        Fieldref: FieldRef;
        a: Integer;
        Existe: Boolean;
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
            //end;

            until TempTenantWebServiceColumns.Next() = 0;


    end;

    procedure InitDefaultColumns(Informe: Record "Informes")
    var
        CreateOrCopy: Option "Create a new data set","Create a copy of an existing data set","Edit an existing data set";
        "Source Service Name": Text;
        "Destination Service Name": Text;

    begin
        InitColumns(Informe."Tipo Objeto", Informe."Id Objeto", CreateOrCopy::"Create a new data set", "Source Service Name", "Destination Service Name");

    end;

    procedure InitColumns(ObjectType: Option ,,,,,,,,"Page","Query"; ObjectID: Integer; InActionType: Option "Create a new data set","Create a copy of an existing data set","Edit an existing data set"; InSourceServiceName: Text; DestinationServiceName: Text)
    var
        AllObj: Record AllObj;
        ApplicationObjectMetadata: Record "Application Object Metadata";
        inStream: InStream;

    begin


        OdataColumnChose.InitColumns(ObjectType, ObjectID, InActionType, InSourceServiceName, DestinationServiceName);

    end;

    Procedure LetraColumna(Columna: integer) xlColID: Text
    var
        x: Integer;
        i: Integer;
        y: Integer;
        c: Char;
        t: Text[30];
    begin
        xlColID := '';
        x := Columna;
        while x > 26 do begin
            y := x mod 26;
            if y = 0 then
                y := 26;
            c := 64 + y;
            i := i + 1;
            t[i] := c;
            x := (x - y) div 26;
        end;
        if x > 0 then begin
            c := 64 + x;
            i := i + 1;
            t[i] := c;
        end;
        for x := 1 to i do
            xlColID[x] := t[1 + i - x];
        exit(xlColID);
    end;


    var
        OdataColumnChose: Page "OData Column Choose SubForm";
}
