page 7001192 "Columnas Informes"
{
    Caption = 'Columnas Informes';
    //DeleteAllowed = false;
    //InsertAllowed = false;
    PageType = ListPart;
    SourceTable = "Columnas Informes";
    SourceTableView = sorting(Mostrar, Orden) order(ascending);



    layout
    {
        area(content)
        {
            repeater(detalle)
            {



                field("Letra"; Rec.Letra)
                {
                    ApplicationArea = All;
                    Width = 5;
                }
                field(Funcion; Rec.Funcion)
                {
                    ApplicationArea = All;
                    trigger OnDrillDown()
                    var
                        Enlaces: Record "Enlaces Informes";
                    begin
                        If Rec.Funcion = Funciones::Enlace then begin
                            If not Enlaces.Get(Rec.Id, Rec.Id_campo) then begin
                                Enlaces.init();
                                Enlaces.Id := Rec.Id;
                                Enlaces."Campo Relación" := Rec.Id_campo;
                                Enlaces."Nombre Campo Relación" := Rec."Field Name";
                                Enlaces.Insert();
                                Commit();
                            end;
                            Enlaces.SetRange("Campo Relación", Rec.Id_campo);
                            Enlaces.SetRange(Id, Rec.Id);
                            Page.RunModal(Page::"Enlaces", Enlaces);
                        end;
                    end;
                }
                field("Campo"; Rec."Campo")
                {
                    ApplicationArea = All;
                }




                field("Título"; Rec.Titulo)
                {
                    ApplicationArea = Basic, Suite;
                    ToolTip = 'Specifies the Field Captions in a data set.';
                    Width = 80;
                }
                field("Ancho Columna"; Rec."Ancho Columna")
                {
                    ApplicationArea = All;
                }

                field("Formato"; Rec."Formato Columna")
                {
                    ApplicationArea = All;
                    Width = 150;
                    Editable = false;
                    trigger OnDrillDown()
                    var
                        Formatos: Record "Formato Columnas";
                        C: Label '"';

                    begin
                        Formatos.setrange(id, Rec.Id);
                        Formatos.setrange(Id_campo, Rec.Id_campo);
                        if Not Formatos.findfirst() then begin
                            Formatos.init();
                            Formatos.Id := Rec.Id;
                            Formatos.Id_campo := Rec.Id_campo;
                            Formatos.Cabecera := true;
                            Formatos.Letra := Rec.Letra;
                            Formatos.Insert();
                            Formatos.init();
                            Formatos.Id := Rec.Id;
                            Formatos.Id_campo := Rec.Id_campo;
                            Formatos.Cabecera := false;
                            Formatos.Letra := Rec.Letra;
                            Formatos.Insert();
                        end;
                        Commit();

                        Page.RunModal(Page::"Formato Columnas", Formatos);
                        Commit();
                        Formatos.SetRange(Cabecera, false);
                        If Formatos.FindFirst() then begin
                            Rec."Formato Columna" := '';
                            if Formatos."Formato Columna" <> '' then
                                Rec."Formato Columna" := 'Formato={' + Formatos."Formato Columna" + '}';
                            If Formatos.Fuente <> '' then
                                if Rec."Formato Columna" <> '' then
                                    Rec."Formato Columna" := CopyStr(Rec."Formato Columna" + ';Fuente={' + Formatos.Fuente, 1, 250) + '}'
                                else
                                    Rec."Formato Columna" := 'Fuente={' + Formatos.Fuente + '}';
                            If Formatos."Color Fuente" <> '' then
                                if Rec."Formato Columna" <> '' then
                                    Rec."Formato Columna" := CopyStr(Rec."Formato Columna" + ';Color Fuente={' + Formatos."Color Fuente", 1, 250) + '}'
                                else
                                    Rec."Formato Columna" := CopyStr('Color Fuente={' + Formatos."Color Fuente", 1, 250) + '}';
                            If Formatos.Tamaño <> 0 then
                                if Rec."Formato Columna" <> '' then
                                    Rec."Formato Columna" := CopyStr(Rec."Formato Columna" + ';Tamaño={' + Format(Formatos.Tamaño), 1, 250) + '}'
                                else
                                    Rec."Formato Columna" := CopyStr('Tamaño={' + Format(Formatos.Tamaño), 1, 250) + '}';
                            If Formatos.Bold then
                                if Rec."Formato Columna" <> '' then
                                    Rec."Formato Columna" := CopyStr(Rec."Formato Columna" + ';Negrita={Sí', 1, 250) + '}'
                                else
                                    Rec."Formato Columna" := CopyStr('Negrita={Sí', 1, 250) + '}';
                            If Formatos.Italic then
                                if Rec."Formato Columna" <> '' then
                                    Rec."Formato Columna" := CopyStr(Rec."Formato Columna" + ';Cursiva={Sí', 1, 250) + '}'
                                else
                                    Rec."Formato Columna" := CopyStr('Cursiva={Sí', 1, 250) + '}';
                            If Formatos.Underline then
                                if Rec."Formato Columna" <> '' then
                                    Rec."Formato Columna" := CopyStr(Rec."Formato Columna" + ';Subrayado={Sí', 1, 250) + '}'
                                else
                                    Rec."Formato Columna" := CopyStr('Subrayado={Sí', 1, 250) + '}';
                            If Formatos."Double Underline" then
                                if Rec."Formato Columna" <> '' then
                                    Rec."Formato Columna" := CopyStr(Rec."Formato Columna" + ';Doble Subrayado={Sí', 1, 250) + '}'
                                else
                                    Rec."Formato Columna" := CopyStr('Doble Subrayado={Sí', 1, 250) + '}';
                            If Formatos."Color Fondo" <> '' then
                                if Rec."Formato Columna" <> '' then
                                    Rec."Formato Columna" := CopyStr(Rec."Formato Columna" + ';Color Fondo={' + Formatos."Color Fondo", 1, 250) + '}'
                                else
                                    Rec."Formato Columna" := CopyStr('Color Fondo={' + Formatos."Color Fondo", 1, 250) + '}';
                            If StrLen(Rec."Formato Columna") = 250 then
                                Rec."Formato Columna" := CopyStr(Rec."Formato Columna", 1, 247) + '...';

                        end;



                    end;
                }
                field("Field Name"; Rec."Field Name")
                {
                    ApplicationArea = Basic, Suite;
                    Caption = 'Nombre del campo';
                    Editable = false;
                    ToolTip = 'Specifies the field names in a data set.';
                }
                field(Orden; Rec.Orden)
                {
                    ApplicationArea = All;
                }
                field(Incluir; Rec.Include)
                {
                    ApplicationArea = Basic, Suite;
                    ToolTip = 'Specifies the field name that is selected in the data set.';

                    trigger OnValidate()
                    begin
                        // if CalledForExcelExport then
                        //     CheckFieldFilter();
                        IsModified := true;
                    end;
                }

            }
        }
    }

    actions
    {
        area(Processing)
        {
            action("Solo Seleccionados")
            {
                ApplicationArea = All;
                Image = Filter;
                trigger OnAction()
                begin
                    Rec.setrange(Include, true);
                end;
            }
            action("Recalcular Mostrar")
            {
                ApplicationArea = All;
                Image = Filter;
                trigger OnAction()
                var
                    Self: Record "Columnas Informes";
                begin
                    If Self.FindFirst() then
                        repeat
                            if Self.Include then
                                Self.Mostrar := Self.Mostrar::" "
                            else
                                Self.Mostrar := Self.Mostrar::"No";

                            Self.Modify();
                        until Self.Next() = 0;
                end;
            }
            action(Todos)
            {
                ApplicationArea = All;
                Image = ClearFilter;
                trigger OnAction()
                begin
                    Rec.setrange(Include);
                end;
            }
        }

    }
    procedure DeleteColumns()
    begin
        Rec.DeleteAll();
    end;

    procedure IncludeIsChanged(): Boolean
    var
        LocalDirty: Boolean;
    begin
        LocalDirty := IsModified;
        Clear(IsModified);
        exit(LocalDirty);
    end;

    trigger OnAfterGetRecord()
    begin
        If (Rec.Letra = '') and (Rec.Orden <> 0) then
            Rec.Letra := Rec.LetraColumna(Rec.Orden);
    end;

    var
        RecRef: RecordRef;
        IsModified: Boolean;




}
