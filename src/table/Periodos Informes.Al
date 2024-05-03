//Crear una tabla para guardar informes
table 7001243 "Periodos Informes"
{
    //ObsoleteState = Removed;
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
                    Titulo := Campos."Field Caption";
                end;
            end;

        }
        field(3; Titulo; Text[250])
        {
            caption = 'Titulo';
            DataClassification = ToBeClassified;
        }

        field(7; "Field Name"; Text[250])
        {
            caption = 'Field Name';
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
        field(6; Periodo; Text[30])
        {
            caption = 'Periodo';
            DataClassification = ToBeClassified;
        }
        field(8; Semana; Boolean)
        {
            caption = 'Semana';
            DataClassification = ToBeClassified;
        }
        field(12; Campo2; Integer)
        {
            caption = 'Campo2';
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
                        Campo2 := Campos."No.";
                    Titulo2 := Campos."Field Caption";
                end;
            end;

        }
        field(13; Titulo2; Text[250])
        {
            caption = 'Titulo2';
            DataClassification = ToBeClassified;
        }

        field(17; "Field Name2"; Text[250])
        {
            caption = 'Field Name2';
            DataClassification = ToBeClassified;
        }
        field(14; Desde2; DateFormula)
        {
            caption = 'Desde (fecha ejecución)1';
            DataClassification = ToBeClassified;
        }
        field(15; "Hasta2"; DateFormula)
        {
            caption = 'Hasta (fecha ejecución)1';
            DataClassification = ToBeClassified;
        }
        field(18; Semana2; Boolean)
        {
            caption = 'Semana';
            DataClassification = ToBeClassified;
        }





    }
    keys
    {
        key(PK; Id, Periodo)
        {
            Clustered = true;
        }


    }
}
//Page para filttros informes


