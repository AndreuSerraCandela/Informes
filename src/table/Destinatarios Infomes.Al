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
                    "e-mail" := Employee."E-Mail";
            end;
        }


        field(3; "e-mail"; Text[80])
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
            begin
                if Informes.Get(Id) then begin
                    Campos.SetRange(Campos.TableNo, Informes."Tabla filtros");
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



    }
    keys
    {
        key(PK; Id, Empleado)
        {
            Clustered = true;
        }

    }
}
