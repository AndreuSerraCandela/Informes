//Crear una tabla para guardar informes
table 7001241 "Destinatarios Informes"
{
    fields
    {
        field(1; Id; Integer)
        {
            caption = 'Id';
            DataClassification = ToBeClassified;
            TableRelation = "Informes"."Id";
        }
        field(2; Empleado; Code[20])
        {
            caption = 'Empleado';
            DataClassification = ToBeClassified;
            TableRelation = Employee."No.";
            trigger OnValidate()
            var
                Employee: Record Employee;
            begin
                if Employee.Get(Empleado)
                 then
                    "e-mail" := Employee."Company E-Mail";
                If "e-mail" = '' Then Message('El empleado no tiene correo electrónico');
            end;
        }


        field(3; "e-mail"; Text[250])
        {
            caption = 'Correo Electrónico';
            DataClassification = ToBeClassified;
        }
        field(4; "Campo Destinatario"; Integer)
        {
            caption = 'Campo Destinatario';
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
                        "Campo Destinatario" := Campos."No.";

                end;
            end;
        }
        field(5; Valor; Text[250])
        {
            caption = 'Valor Filtro';
            DataClassification = ToBeClassified;
        }
        field(6; "Nombre Informe"; text[250])
        {
            caption = 'Nombre Informe';
            DataClassification = ToBeClassified;
        }



    }
    keys
    {
        key(PK; Id, Empleado)
        {
            Clustered = true;
        }

    }
}
