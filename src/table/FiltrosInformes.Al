//Crear una tabla para guardar informes
table 7001242 "Filtros Informes"
{
    fields
    {
        field(1; Id; Integer)
        {
            AutoIncrement = true;
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
                if Informes.Get(Id) then
                    Campos.SetRange(Campos.TableNo, Informes."Tabla filtros");
            end;

        }
        field(3; Valor; Text[250])
        {
            caption = 'Valor Filtro';
            DataClassification = ToBeClassified;
        }
        field(4; Desde; DateFormula)
        {
            caption = 'Desde (fecha ejecución)';
            DataClassification = ToBeClassified;
        }
        field(5; "Hasta"; DateFormula)
        {
            caption = 'Hasta (fecha ejecución)';
            DataClassification = ToBeClassified;
        }

    }
    keys
    {
        key(PK; Id, Campo)
        {
            Clustered = true;
        }

    }
}
//Page para filttros informes


