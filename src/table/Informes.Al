//Crear una tabla para guardar informes
table 7001240 Informes
{
    fields
    {
        field(1; Id; Integer)
        {
            caption = 'Id';
            DataClassification = ToBeClassified;
            TableRelation = "Informes"."Id";
        }
        field(2; Descripcion; Text[250])
        {
            caption = 'Descripcion';
            DataClassification = ToBeClassified;
        }
        field(3; Periodicidad; DateFormula)
        {
            caption = 'Periodicidad';
            DataClassification = ToBeClassified;
        }
        field(4; "Fecha Primera Ejecución"; Date)
        {
            caption = 'Fecha 1ª Ejecución';
            DataClassification = ToBeClassified;
        }
        field(5; Informe; Integer)
        {
            caption = 'Tabla';
            DataClassification = ToBeClassified;
            TableRelation = AllObjWithCaption."Object ID" WHERE("Object Type" = filter(Report | Page));
            trigger OnValidate()
            begin
                Case Informe of

                End;
            end;
        }
        field(6; "Tabla filtros"; Integer)
        {
            caption = 'Tabla para filtros';
            TableRelation = AllObjWithCaption."Object ID" WHERE("Object Type" = CONST(Table));
        }

    }
    keys
    {
        key(PK; Id)
        {
            Clustered = true;
        }
    }
}
//crear page para informes