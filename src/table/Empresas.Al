//Crear una tabla para guardar informes
table 7001245 "Empresas Informes"
{
    //ObsoleteState = Removed;
    fields
    {
        field(1; Id; Integer)
        {
            caption = 'Id';
            DataClassification = ToBeClassified;
        }
        field(2; Empresa; Text[30])
        {
            TableRelation = Company;

        }
        field(3; HojaExcel; Text[100])
        {
            DataClassification = ToBeClassified;
        }
        field(4; "Hojas Separadas"; Boolean)
        {
            DataClassification = ToBeClassified;
        }
        field(5; "Columna Excel"; Integer)
        {
            DataClassification = ToBeClassified;
        }
        field(6; "Incluir"; Boolean)
        {
            DataClassification = ToBeClassified;
        }
        






    }
    keys
    {
        key(PK; Id, Empresa)
        {
            Clustered = true;
        }


    }
}
//Page para filttros informes


