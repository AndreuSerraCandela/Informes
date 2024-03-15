//Pagina con Tabla temporar de proveedores y saldo x empresa y dos años
page 7001199 "Proveedores Empresa"
{
    PageType = List;
    ApplicationArea = All;
    UsageCategory = Lists;
    SourceTable = Vendor;
    SourceTableTemporary = true;
    layout
    {
        area(content)
        {
            group(Filtros)
            {
                field("Año 1"; "Año 1")
                {
                    ApplicationArea = All;
                }
                field("Año 2"; "Año 2")
                {
                    ApplicationArea = All;
                }
                field("Incluir Proveedores Bloqueados"; "IncluirProveedoresBloqueados")
                {
                    ApplicationArea = All;

                }
                field("Incluir empresas del grupo"; "IncluirEmpresasDelGrupo")
                {
                    ApplicationArea = All;
                }
                field("Incluir Proveedores sin Cif"; "IncluirProveedoresSinCif")
                {
                    ApplicationArea = All;
                }
            }

            repeater(Proveedores)
            {
                field(Cif; Rec."VAT Registration No.")
                {
                    ApplicationArea = All;
                }
                field(Nombre; Rec.Name)
                {
                    ApplicationArea = All;
                }
                field(Empresa; PEmpresa())
                {
                    ApplicationArea = All;
                }
                field(Saldo; "PSaldo"('AC'))
                {
                    CaptionClass = Cap(1);
                    DecimalPlaces = 2 : 2;
                    ApplicationArea = All;


                }
                field(SaldoAnterior; "PSaldo"('AA'))
                {
                    CaptionClass = Cap(2);
                    DecimalPlaces = 2 : 2;
                    ApplicationArea = All;

                }
            }
        }
    }
    actions
    {
        area(processing)
        {
            action(Actualizar)
            {
                ApplicationArea = All;
                Image = RefreshLines;
                trigger OnAction()
                var
                    rEmp: Record "Company";
                    Vendor: Record Vendor;
                    DetLedgerEntry: Record "Detailed Vendor Ledg. Entry";
                    Control: Codeunit ControlProcesos;
                    rInf: Record "Company Information";
                    Ventana: Dialog;
                    a: Integer;
                begin
                    Rec.DeleteAll();
                    rEmp.SetRange("Evaluation Company", false);
                    Ventana.OPEN('Procesando Proveedor ############1# de ############2#\' +
                                'Empresa ############################################3#\');
                    if rEmp.FINDFIRST THEN
                        REPEAT
                            if Control.Permiso_Empresas(rEmp.Name) then begin
                                Ventana.Update(3, rEmp.Name);
                                a := 0;
                                rInf.ChangeCompany(rEmp."Name");
                                rinf.Get();
                                if rInf."Clave Recursos" <> '' then begin
                                    Vendor.ChangeCompany(rEmp."Name");
                                    If not IncluirProveedoresBloqueados then
                                        Vendor.SETRANGE("Blocked", Vendor."Blocked"::" ");
                                    If not IncluirProveedoresSinCif then
                                        Vendor.SETFILTER("VAT Registration No.", '<>%1', '');
                                    Ventana.Update(2, Vendor.COUNT);
                                    if Vendor.FindFirst() then begin
                                        REPEAT
                                            a += 1;
                                            Ventana.Update(1, a);
                                            Rec := Vendor;
                                            Rec."No." := rInf."Clave Recursos" + Vendor."No.";
                                            Rec.Address := rEmp.Name;
                                            DetLedgerEntry.ChangeCompany(rEmp."Name");
                                            DetLedgerEntry.SETRANGE("Vendor No.", Vendor."No.");
                                            DetLedgerEntry.SETRANGE("Posting Date", 0D, DMY2Date(31, 12, "Año 1"));
                                            DetLedgerEntry.SETRANGE("Excluded from calculation", false);
                                            DetLedgerEntry.CALCSUMS("Amount");
                                            Rec."Budgeted Amount" := -DetLedgerEntry."Amount";
                                            DetLedgerEntry.SETRANGE("Vendor No.", Vendor."No.");
                                            DetLedgerEntry.SETRANGE("Posting Date", 0D, DMY2Date(31, 12, "Año 2"));
                                            DetLedgerEntry.SETRANGE("Excluded from calculation", false);
                                            DetLedgerEntry.CALCSUMS("Amount");
                                            Rec."Prepayment %" := -DetLedgerEntry."Amount";
                                            if (IncluirEmpresasDelGrupo) or (Vendor."IC Partner Code" = '') then
                                                Rec.Insert();

                                        until Vendor.NEXT = 0;
                                    end;
                                end;
                            end;
                        until rEmp.NEXT = 0;
                    Ventana.Close();
                end;

            }
        }
        area(Promoted)
        {
            actionref(Actualizar_ref; Actualizar) { }
        }
    }

    PROCEDURE Cap(num: Integer): Text[30];
    BEGIN
        IF num = 1 THEN
            EXIT('3,' + Format("Año 1"))
        ELSE
            EXIT('3,' + Format("Año 2"));


    END;

    procedure PEmpresa(): Text[30];
    begin
        Exit(Rec.Address);
    end;

    procedure PSaldo(AA: Text): Decimal;
    begin
        If AA = 'AA' then
            Exit(Rec."Prepayment %")
        else
            Exit(Rec."Budgeted Amount");
    end;

    var
        "Año 1": Integer;
        "Año 2": Integer;
        "IncluirProveedoresBloqueados": Boolean;
        "IncluirEmpresasDelGrupo": Boolean;
        "IncluirProveedoresSinCif": Boolean;
}