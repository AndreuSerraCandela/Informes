//Crear una tabla para guardar informes
table 7001244 "Columnas Informes"
{
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
            begin
                if Informes.Get(Id) then begin
                    Campos.SetRange(Campos.TableNo, Informes."Tabla filtros");
                    If Page.Runmodal(9806, Campos) = Action::LookupOK then
                        Campo := Campos."No.";
                    Titulo := Campos."Field Caption";
                end;
            end;

        }
        field(3; Titulo; Text[250])
        {
            caption = 'Titulo';
            DataClassification = ToBeClassified;
        }
        field(4; Orden; Integer)
        {
            caption = 'Orden';
            DataClassification = ToBeClassified;
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


