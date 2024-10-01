//Crear una tabla para guardar informes
table 7001237 "Enlaces Informes"
{
    LookupPageId = "Enlaces";
    DrillDownPageId = "Enlaces";
    fields
    {
        field(1; Id; Integer)
        {
            caption = 'Id';
            DataClassification = ToBeClassified;
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
        field(2; "Tabla"; Integer)
        {
            caption = 'Tablafiltros';
            TableRelation = AllObjWithCaption."Object ID" WHERE("Object Type" = CONST(Table));
        }
        field(3; "Campo Enlace"; Integer)
        {
            caption = 'Campo Enlace';
            trigger OnLookup()
            var
                Campos: Record Field;

            begin
                Campos.SetRange(Campos.TableNo, Tabla);
                If Page.Runmodal(9806, Campos) = Action::LookupOK then
                    "Campo Enlace" := Campos."No.";
                "Field Name Enlace" := Campos.FieldName;
            end;

            trigger OnValidate()
            var
                Campos: Record Field;

            begin
                Campos.SetRange(Campos.TableNo, Tabla);
                Campos.SetRange(Campos."No.", "Campo Enlace");
                Campos.FindFirst();
                "Field Name Enlace" := Campos.FieldName;

            end;


        }


        field(7; "Field Name Enlace"; Text[250])
        {
            caption = 'Nombre Campo Enlace';
            DataClassification = ToBeClassified;
        }
        field(8; "Campo Datos"; Integer)
        {
            caption = 'Campo Datos';
            trigger OnLookup()
            var
                Campos: Record Field;

            begin
                Campos.SetRange(Campos.TableNo, Tabla);
                If Page.Runmodal(9806, Campos) = Action::LookupOK then
                    "Campo Datos" := Campos."No.";
                "Field Name Datos" := Campos.FieldName;
            end;

            trigger OnValidate()
            var
                Campos: Record Field;

            begin
                Campos.SetRange(Campos.TableNo, Tabla);
                Campos.SetRange(Campos."No.", "Campo Enlace");
                Campos.FindFirst();
                "Field Name Datos" := Campos.FieldName;

            end;


        }


        field(9; "Field Name Datos"; Text[250])
        {
            caption = 'Nombre Campo Datos';
            DataClassification = ToBeClassified;
        }








    }
    keys
    {
        key(PK; Id, "Campo Relación")
        {
            Clustered = true;
        }



    }

}
