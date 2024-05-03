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
                    If Page.Runmodal(9806, Campos) = Action::LookupOK then
                        Campo := Campos."No.";

                end;
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


