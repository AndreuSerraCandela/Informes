//Crear una tabla para guardar informes
table 7001242 "Filtros Informes"
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
                Columna: Record "Columnas Informes";
                ColumnaTemp: Record "Columnas Informes" temporary;
            begin
                if Informes.Get(Id) then begin
                    Columna.SetRange(Columna."Id", Informes."Id");
                    If not Columna.FindFirst() then begin
                        Informes.InitDefaultColumns();
                        Informes.GetColumns(Informes);

                        Commit();
                        Columna.FindFirst();
                    end else
                        Informes.InitDefaultColumns();

                    Informes.GetColumnsTemp(ColumnaTemp);
                    If Page.Runmodal(Page::"Lista Columnas", ColumnaTemp) = Action::LookupOK then
                        Campo := ColumnaTemp.Campo;

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
        field(6; "Fecha Cierre"; Boolean)
        { }

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


