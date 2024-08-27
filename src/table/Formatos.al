//Page para filttros informes
table 7001246 "Formato Columnas"
{
    LookupPageId = "Formato Columnas";
    DrillDownPageId = "Formato Columnas";
    fields
    {
        field(1; Id; Integer)
        {
            caption = 'Id';
            DataClassification = ToBeClassified;
        }
        field(2; Orden; Integer)
        {
            caption = 'Id';
            DataClassification = ToBeClassified;
        }

        field(22; Letra; Code[2])
        {
            caption = 'Letra';
            DataClassification = ToBeClassified;
            trigger OnValidate()
            begin
                Letra := LetraColumna(Orden);
            end;
        }
        field(3; Cabecera; Boolean)
        {
            caption = 'Cabecera';
            DataClassification = ToBeClassified;

        }

        field(4; "Id_campo"; Integer)
        {
            caption = 'Id_campo';
            AutoIncrement = true;
        }
        field(5; "Bold"; Boolean)
        {
            caption = 'Negrita';
            DataClassification = ToBeClassified;
        }
        field(6; "Italic"; Boolean)
        {
            caption = 'Cursiva';
            DataClassification = ToBeClassified;
        }
        field(7; "Underline"; Boolean)
        {
            caption = 'Subrayado';
            DataClassification = ToBeClassified;
        }
        field(8; "Double Underline"; Boolean)
        {
            caption = 'Doble Subrayado';
            DataClassification = ToBeClassified;
        }


        field(11; "Formato Columna"; Text[250])
        {
            caption = 'Formato Columna';
            DataClassification = ToBeClassified;
        }
        field(12; Fuente; Text[250])
        {
            caption = 'Fuente';
            DataClassification = ToBeClassified;
        }
        field(13; Tamaño; Integer)
        {
            caption = 'Tamaño';
            DataClassification = ToBeClassified;
        }
        field(14; "Color Fuente"; Text[30])
        {
            caption = 'Color';
            DataClassification = ToBeClassified;

        }
        field(15; "Color Fondo"; Text[30])
        {
            caption = 'Color Fondo';
            DataClassification = ToBeClassified;

        }
        field(16; "Insertar Vínculo"; Boolean)
        {

            DataClassification = ToBeClassified;
        }
        field(17; "Tabla Hipervínculo"; Integer)
        {
            TableRelation = AllObjWithCaption."Object ID" where("Object Type" = CONST(Table));
        }
        field(18; "Campo Hipervínculo"; Integer)
        {
            trigger OnLookup()
            var
                Campos: Record Field;

            begin
                Campos.SetRange(Campos.TableNo, "Tabla Hipervínculo");
                If Page.Runmodal(9806, Campos) = Action::LookupOK then
                    "Campo Hipervínculo" := Campos."No.";


            end;
        }
        field(19; "Campo Relación"; Integer)
        {
            trigger OnLookup()
            var
                Informes: Record "Informes";
                Campos: Record Field;
                Columna: Record "Columnas Informes";
            begin
                if Informes.Get(Id) then begin
                    Columna.SetRange(Columna."Id", Informes."Id");
                    If not Columna.FindFirst() then begin
                        Informes.InitDefaultColumns();
                        Informes.GetColumns(Informes);
                        Commit();
                        Columna.FindFirst();
                    end;
                    Campos.SetRange(Campos.TableNo, Columna.Table);
                    If Page.Runmodal(9806, Campos) = Action::LookupOK then begin
                        "Campo Relación" := Campos."No.";
                        "Nombre Campo Relación" := Campos.FieldName;
                    end;


                end;
            end;
        }
        field(20; "Nombre Campo Relación"; Text[250])
        {
            caption = 'Nombre Campo Relación';
            DataClassification = ToBeClassified;
        }
    }
    keys
    {
        key(PK; Id, Id_Campo, Cabecera)
        {
            Clustered = true;
        }

    }
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
}