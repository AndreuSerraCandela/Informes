//Crear una tabla para guardar informes
table 7001240 Informes
{
    fields
    {
        field(1; Id; Integer)
        {
            caption = 'Id';
            DataClassification = ToBeClassified;
            AutoIncrement = true;
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
            trigger OnValidate()
            begin
                If "Fecha Primera Ejecución" <> 0D then
                    "Próxima Fecha" := CalcDate(Periodicidad, "Fecha Primera Ejecución");
            end;
        }
        field(4; "Fecha Primera Ejecución"; Date)
        {
            caption = 'Fecha 1ª Ejecución';
            DataClassification = ToBeClassified;
            trigger OnValidate()
            begin
                "Próxima Fecha" := CalcDate(Periodicidad, "Fecha Primera Ejecución");
            end;
        }
        field(5; Informe; Enum Informes)
        {
            caption = 'Informe';
            DataClassification = ToBeClassified;
            trigger OnValidate()
            begin
                Case Informe of
                    Informes::"Contratos x Empresa":
                        Validate("Tabla filtros", 36);
                    Informes::"Estadisticas Contabilidad":
                        Validate("Tabla filtros", 17);
                End;
            end;
        }
        field(6; "Tabla filtros"; Integer)
        {
            caption = 'Tabla para filtros';
            TableRelation = AllObjWithCaption."Object ID" WHERE("Object Type" = CONST(Table));
        }
        field(7; "Próxima Fecha"; Date)
        {
            caption = 'Fecha 1ª Ejecución';
            DataClassification = ToBeClassified;
        }
        field(9; "Plantilla Excel"; Blob)
        {
            caption = 'Plantilla Excel';
            DataClassification = ToBeClassified;
        }
        field(10; "Crear Tarea"; Boolean)
        {
            caption = 'Crear Tarea';
            DataClassification = ToBeClassified;
        }
        field(11; "Descripcion Tarea"; Text[250])
        {

        }
        field(13; "Id Objeto"; Integer)
        {
            TableRelation = AllObjWithCaption."Object ID" WHERE("Object Type" = field("Tipo Objeto"));
        }
        field(14; "Tipo Objeto"; Option)
        {
            OptionMembers = ,,,,,,,,"Page","Query";
            OptionCaption = ',,,,,,,,Página,Consulta';
            caption = 'Tipo Objeto';
            DataClassification = ToBeClassified;
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