//Pagina con Tabla temporar de proveedores y saldo x empresa y dos años
page 7001194 "Clientes Empresa"
{
    PageType = List;
    ApplicationArea = All;
    UsageCategory = Lists;
    SourceTable = Customer;
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
                field("Incluir Clientes Bloqueados"; "IncluirClientesBloqueados")
                {
                    ApplicationArea = All;

                }
                field("Incluir empresas del grupo"; "IncluirEmpresasDelGrupo")
                {
                    ApplicationArea = All;
                }
                field("Incluir clientes sin Cif"; "IncluirClientesSinCif")
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
                    Customer: Record Customer;
                    DetLedgerEntry: Record "Detailed Cust. Ledg. Entry";
                    Control: Codeunit ControlProcesos;
                    rInf: Record "Company Information";
                    a: Integer;
                    Ventana: Dialog;
                begin
                    Rec.DeleteAll();
                    rEmp.SetRange("Evaluation Company", false);
                    Ventana.OPEN('Procesando Clientes ############1# de ############2#\' +
                                'Empresa ############################################3#\');
                    if rEmp.FINDFIRST THEN
                        REPEAT
                            if Control.Permiso_Empresas(rEmp.Name) then begin
                                rInf.ChangeCompany(rEmp."Name");
                                a := 0;
                                Ventana.Update(3, rEmp.Name);
                                rinf.Get();
                                if rInf."Clave Recursos" <> '' then begin
                                    Customer.ChangeCompany(rEmp."Name");
                                    If not IncluirClientesBloqueados then
                                        Customer.SETRANGE("Blocked", Customer."Blocked"::" ");
                                    If not IncluirClientesSinCif then
                                        Customer.SETFILTER("VAT Registration No.", '<>%1', '');
                                    Ventana.Update(2, Customer.COUNT);
                                    if Customer.FindFirst() then begin
                                        REPEAT
                                            a += 1;
                                            Ventana.Update(1, a);
                                            Rec := Customer;
                                            Rec."No." := rInf."Clave Recursos" + Customer."No.";
                                            Rec.Address := rEmp.Name;
                                            DetLedgerEntry.ChangeCompany(rEmp."Name");
                                            DetLedgerEntry.SETRANGE("Customer No.", Customer."No.");
                                            DetLedgerEntry.SETRANGE("Posting Date", 0D, DMY2Date(31, 12, "Año 1"));
                                            DetLedgerEntry.SETRANGE("Excluded from calculation", false);
                                            DetLedgerEntry.CALCSUMS("Amount");
                                            Rec."Budgeted Amount" := DetLedgerEntry."Amount";
                                            DetLedgerEntry.SETRANGE("Customer No.", Customer."No.");
                                            DetLedgerEntry.SETRANGE("Posting Date", 0D, DMY2Date(31, 12, "Año 2"));
                                            DetLedgerEntry.SETRANGE("Excluded from calculation", false);
                                            DetLedgerEntry.CALCSUMS("Amount");
                                            Rec."Prepayment %" := DetLedgerEntry."Amount";
                                            if (IncluirEmpresasDelGrupo) or (Customer."IC Partner Code" = '') then
                                                Rec.Insert();


                                        until Customer.NEXT = 0;
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
        "IncluirClientesBloqueados": Boolean;
        "IncluirEmpresasDelGrupo": Boolean;
        IncluirClientesSinCif: Boolean;
}