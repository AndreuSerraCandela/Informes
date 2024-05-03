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

            begin

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
                Titulo := Campos."Field Caption";
                "Field Name" := Campos.FieldName;
            end;

            trigger OnValidate()
            var
                Informes: Record "Informes";
                Campos: Record Field;
                Columna: Record "Columnas Informes";

            begin

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
            end;


        }
        field(3; Titulo; Text[250])
        {
            caption = 'Titulo';
            DataClassification = ToBeClassified;
            trigger OnValidate()
            begin
                if Titulo <> '' then
                    Include := true;
            end;
        }
        field(4; Orden; Integer)
        {
            caption = 'Orden';
            DataClassification = ToBeClassified;
            trigger OnValidate()
            begin
                if Orden = 0 then
                    Include := false
                else
                    Include := true;
            end;
        }
        field(5; Funcion; Enum Funciones)
        {
            caption = 'Funcion';
            DataClassification = ToBeClassified;

        }
        field(6; Include; Boolean)
        {
            caption = 'Incluir';
            DataClassification = ToBeClassified;
            trigger OnValidate()
            begin
                if not Include then
                    Orden := 0;
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





    }
    keys
    {
        key(PK; Id, Id_Campo)
        {
            Clustered = true;
        }

    }
}
//Page para filttros informes


