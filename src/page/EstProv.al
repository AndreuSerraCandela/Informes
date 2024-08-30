page 7001181 "Lista Proveedores Compra_Venta"
{
    Caption = 'Lista Proveedores Compra_Venta';
    PageType = List;
    ApplicationArea = All;
    Editable = true;
    UsageCategory = Lists;
    CardPageId = "Ficha Proveedor Compra";
    SourceTable = "Proveedores Compra";
    layout
    {
        area(Content)
        {
            group(Fechas)
            {

                field(Desde; Desde)
                {
                    ApplicationArea = all;
                    trigger OnValidate()
                    begin
                        CalculadoVentas := false;
                        if (Hasta <> 0D) And (Desde <> 0D) Then Rec.SetRange("Date Filter", Desde, Hasta);
                    end;
                }
                field(Hasta; Hasta)
                {
                    ApplicationArea = all;
                    trigger OnValidate()
                    begin
                        CalculadoVentas := false;
                        if (Hasta <> 0D) And (Desde <> 0D) Then Rec.SetRange("Date Filter", Desde, Hasta);
                    end;
                }
                field(Tipo; TipoEstadistica)
                {
                    ApplicationArea = All;
                    trigger OnValidate()
                    begin
                        CalculadoVentas := false;

                    end;
                }
            }
            repeater(Detalle)
            {
                Editable = false;


                field(Name; rec.Name)
                {
                    ApplicationArea = All;
                }

                field(ColumnValues1; ColumnValues[1])
                {
                    ApplicationArea = All;
                    //AutoFormatExpression = FormatStr(1);
                    AutoFormatType = 11;
                    BlankZero = true;
                    CaptionClass = '3,' + ColumnCaptions[1];
                    StyleExpr = ColumnStyle1;

                    trigger OnDrillDown()
                    begin
                        DrillDown(1);
                    end;
                }
                field(ColumnValues2; ColumnValues[2])
                {
                    ApplicationArea = All;
                    //AutoFormatExpression = FormatStr(2);
                    AutoFormatType = 11;
                    BlankZero = true;
                    CaptionClass = '3,' + ColumnCaptions[2];
                    StyleExpr = ColumnStyle2;
                    Visible = NoOfColumns >= 2;

                    trigger OnDrillDown()
                    begin
                        DrillDown(2);
                    end;
                }
                field(ColumnValues3; ColumnValues[3])
                {
                    ApplicationArea = All;
                    //AutoFormatExpression = FormatStr(3);
                    AutoFormatType = 11;
                    BlankZero = true;
                    CaptionClass = '3,' + ColumnCaptions[3];
                    StyleExpr = ColumnStyle3;
                    Visible = NoOfColumns >= 3;

                    trigger OnDrillDown()
                    begin
                        DrillDown(3);
                    end;
                }
                field(ColumnValues4; ColumnValues[4])
                {
                    ApplicationArea = All;
                    //AutoFormatExpression = FormatStr(4);
                    AutoFormatType = 11;
                    BlankZero = true;
                    CaptionClass = '3,' + ColumnCaptions[4];
                    StyleExpr = ColumnStyle4;
                    Visible = NoOfColumns >= 4;

                    trigger OnDrillDown()
                    begin
                        DrillDown(4);
                    end;
                }
                field(ColumnValues5; ColumnValues[5])
                {
                    ApplicationArea = All;
                    //AutoFormatExpression = FormatStr(5);
                    AutoFormatType = 11;
                    BlankZero = true;
                    CaptionClass = '3,' + ColumnCaptions[5];
                    StyleExpr = ColumnStyle5;
                    Visible = NoOfColumns >= 5;

                    trigger OnDrillDown()
                    begin
                        DrillDown(5);
                    end;
                }
                field(ColumnValues6; ColumnValues[6])
                {
                    ApplicationArea = All;
                    //AutoFormatExpression = FormatStr(6);
                    AutoFormatType = 11;
                    BlankZero = true;
                    CaptionClass = '3,' + ColumnCaptions[6];
                    StyleExpr = ColumnStyle6;
                    Visible = NoOfColumns >= 6;

                    trigger OnDrillDown()
                    begin
                        DrillDown(6);
                    end;
                }
                field(ColumnValues7; ColumnValues[7])
                {
                    ApplicationArea = All;
                    //AutoFormatExpression = FormatStr(7);
                    AutoFormatType = 11;
                    BlankZero = true;
                    CaptionClass = '3,' + ColumnCaptions[7];
                    StyleExpr = ColumnStyle7;
                    Visible = NoOfColumns >= 7;

                    trigger OnDrillDown()
                    begin
                        DrillDown(7);
                    end;
                }
                field(ColumnValues8; ColumnValues[8])
                {
                    ApplicationArea = All;
                    //AutoFormatExpression = FormatStr(8);
                    AutoFormatType = 11;
                    BlankZero = true;
                    CaptionClass = '3,' + ColumnCaptions[8];
                    StyleExpr = ColumnStyle8;
                    Visible = NoOfColumns >= 8;

                    trigger OnDrillDown()
                    begin
                        DrillDown(8);
                    end;
                }
                field(ColumnValues9; ColumnValues[9])
                {
                    ApplicationArea = All;
                    //AutoFormatExpression = FormatStr(9);
                    AutoFormatType = 11;
                    BlankZero = true;
                    CaptionClass = '3,' + ColumnCaptions[9];
                    StyleExpr = ColumnStyle9;
                    Visible = NoOfColumns >= 9;

                    trigger OnDrillDown()
                    begin
                        DrillDown(9);
                    end;
                }
                field(ColumnValues10; ColumnValues[10])
                {
                    ApplicationArea = All;
                    //AutoFormatExpression = FormatStr(10);
                    AutoFormatType = 11;
                    BlankZero = true;
                    CaptionClass = '3,' + ColumnCaptions[10];
                    StyleExpr = ColumnStyle10;
                    Visible = NoOfColumns >= 10;

                    trigger OnDrillDown()
                    begin
                        DrillDown(10);
                    end;
                }
                field(ColumnValues11; ColumnValues[11])
                {
                    ApplicationArea = All;
                    //AutoFormatExpression = FormatStr(11);
                    AutoFormatType = 11;
                    BlankZero = true;
                    CaptionClass = '3,' + ColumnCaptions[11];
                    StyleExpr = ColumnStyle11;
                    Visible = NoOfColumns >= 11;

                    trigger OnDrillDown()
                    begin
                        DrillDown(11);
                    end;
                }
                field(ColumnValues12; ColumnValues[12])
                {
                    ApplicationArea = All;
                    //AutoFormatExpression = FormatStr(12);
                    AutoFormatType = 11;
                    BlankZero = true;
                    CaptionClass = '3,' + ColumnCaptions[12];
                    StyleExpr = ColumnStyle12;
                    Visible = NoOfColumns >= 12;

                    trigger OnDrillDown()
                    begin
                        DrillDown(12);
                    end;
                }
            }
        }
    }
    actions
    {
        area(Processing)
        {
            action("Ficha")
            {
                ApplicationArea = All;
                Image = List;
                ShortCutKey = 'Mayús+F5';
                //       CaptionML=ESP=Ficha;
                RunObject = Page "Ficha Intercambio";
                RunPageLink = "No." = FIELD("No.");
            }
            action("&Desplegar")
            {
                ApplicationArea = All;
                Image = BOMLevel;
                ShortCutKey = F11;
                //CaptionML=ESP=Desplegar;
                trigger OnAction()
                VAR
                    rEmpresa: Record 2000000006;
                    Customer: Record 18;
                    Vendor: Record 23;
                    RxE: Record "Proveedores x Empresa";
                    Cli: Code[20];
                    Prov: Code[20];

                BEGIN
                    Desplegar(Rec);
                END;
            }
            action("&Ver Provee x Empresa")
            {
                ApplicationArea = All;
                Image = Find;
                //CaptionML=ESP=Ver Clientes/Provee x Empresa;
                trigger OnAction()
                VAR
                    RxE: Record "Proveedores x Empresa";
                BEGIN
                    RxE.SETRANGE(RxE."Código Proveedor", Rec."No.");
                    Page.RUNMODAL(Page::"Proveedores x Empresa", RxE);
                END;
            }
            action("Copiar &Datos del Proveedor")
            {
                ApplicationArea = All;
                ShortCutKey = F6;
                Image = Copy;
                //CaptionML=ESP=Copiar Datos del cliente;
                trigger OnAction()
                VAR
                    Customer: Record 23;
                BEGIN
                    if Page.RUNMODAL(0, Customer) = ACTION::LookupOK THEN BEGIN
                        //"No.":=Customer."No.";
                        if Rec."No." = '' THEN Rec.INSERT(TRUE);
                        //"Search Name":=Customer."Search Name";
                        Rec.Name := Customer.Name;
                        Rec."Name 2" := Customer."Name 2";
                        Rec.Address := Customer.Address;
                        Rec."Address 2" := Customer."Address 2";
                        Rec.City := Customer.City;
                        Rec.Contact := Customer.Contact;
                        Rec."Phone No." := Customer."Phone No.";
                        Rec."Telex No." := Customer."Telex No.";
                        Rec."Our Account No." := Customer."Our Account No.";
                        Rec."Territory Code" := Customer."Territory Code";
                        Rec."Global Dimension 1 Code" := Customer."Global Dimension 1 Code";
                        Rec."Global Dimension 2 Code" := Customer."Global Dimension 2 Code";
                        Rec."Budgeted Amount" := Customer."Budgeted Amount";
                        Rec."Currency Code" := Customer."Currency Code";
                        Rec."Payment Terms Code" := Customer."Payment Terms Code";
                        Rec."Shipment Method Code" := Customer."Shipment Method Code";
                        Rec."Shipping Agent Code" := Customer."Shipping Agent Code";
                        Rec."Country/Region Code" := Customer."Country/Region Code";
                        Rec.Blocked := Customer.Blocked;
                        Rec."Payment Method Code" := Customer."Payment Method Code";
                        Rec."Last Date Modified" := Customer."Last Date Modified";
                        Rec."Fax No." := Customer."Fax No.";
                        Rec."VAT Registration No." := Customer."VAT Registration No.";
                        Rec."Post Code" := Customer."Post Code";
                        Rec.County := Customer.County;
                        Rec."E-Mail" := Customer."E-Mail";
                        Rec."Home Page" := Customer."Home Page";
                        Rec."Primary Contact No." := Customer."Primary Contact No.";

                        Rec.MODIFY;
                    END;
                END;
            }

            action("Ver &Detalle")
            {
                ApplicationArea = All;
                Image = AllLines;
                // PushAction=RunObject;                // CaptionML=ESP=Ver Detalle Todos;
                trigger OnAction()
                var
                    Impo: Decimal;
                begin
                    case TipoEstadistica of
                        TipoEstadistica::Compra:
                            DrillDownTodos(false);//Page.RunModal(Page::"Proveedores x Empresa");
                        TipoEstadistica::Venta:
                            DrillDownVenta('', '', Impo, true);
                        TipoEstadistica::Contabiliad:
                            DrillDownContableTodos();
                        TipoEstadistica::"Contabilidad Pedidos":
                            DrillDownTodos(true);
                    end;

                end;

            }

            action("Grupo Compras")
            {
                ApplicationArea = All;
                Image = Company;
                trigger OnAction()
                var
                    Gr: Record "Grupo de empresas";
                    EmpGr: Record "Empresa grupo";
                begin
                    if not Gr.Get('COMPRAS', 'Grupo Compras') Then begin
                        Gr.Codigo := 'COMPRAS';
                        Gr.Descripcion := 'Grupo Compras';
                        Gr.Insert();
                        Commit();
                    end;
                    EmpGr.SetRange("Cod. grupo", 'COMPRAS');
                    Page.RunModal(0, EmpGr);
                end;
            }
            action("&Calcular")
            {
                ApplicationArea = All;
                Scope = Repeater;
                Image = Calculate;
                ShortCutKey = F9;
                trigger OnAction()
                BEGIN
                    Case TipoEstadistica of
                        TipoEstadistica::Compra:
                            Calcular('', false);
                        TipoEstadistica::Venta:
                            CalcularVenta('');
                        TipoEstadistica::Contabiliad:
                            CalcularContabilidad('');
                        TipoEstadistica::"Contabilidad Pedidos":
                            Calcular('', true);
                    End;

                END;
            }
        }
    }
    var
        T: Boolean;
        CalculadoVentas: Boolean;
        rLinVentaT: Record 37 temporary;
        AL: Boolean;
        "F-T": Boolean;
        "F+A": Boolean;
        MatrixRec: Record "Proveedores x Empresa";
        ColumnOffset: Integer;
        Ventana: Dialog;
        a: Integer;
        ColumnValues: array[100] of Decimal;
        ColumnCaptions: array[100] of Text[100];
        NoOfColumns: Integer;
        Desde: Date;
        Hasta: Date;
        ColumnStyle1: Text;
        ColumnStyle2: Text;
        ColumnStyle3: Text;
        ColumnStyle4: Text;
        ColumnStyle5: Text;
        ColumnStyle6: Text;
        ColumnStyle7: Text;
        ColumnStyle8: Text;
        ColumnStyle9: Text;
        ColumnStyle10: Text;
        ColumnStyle11: Text;
        ColumnStyle12: Text;
        TipoEstadistica: Option Compra,Venta,Contabiliad,"Contabilidad Pedidos";


    trigger OnAfterGetRecord()
    var
        ColumnNo: Integer;
        ColumnNo2: Integer;
        Emp: Record "Empresa grupo";
        IxE: Record "Proveedores x Empresa";
    begin
        ColumnStyle1 := 'Standard';
        ColumnStyle2 := 'StrongAccent';
        ColumnStyle3 := 'Standard';
        ColumnStyle4 := 'StrongAccent';
        ColumnStyle5 := 'Standard';
        ColumnStyle6 := 'StrongAccent';
        ColumnStyle7 := 'Standard';
        ColumnStyle8 := 'StrongAccent';
        ColumnStyle9 := 'Standard';
        ColumnStyle10 := 'StrongAccent';
        ColumnStyle11 := 'Standard';
        ColumnStyle12 := 'StrongAccent';
        Clear(ColumnValues);
        ColumnOffset := 0;
        MatrixRec.Reset();
        ColumnNo := 0;
        Emp.Reset();
        Emp.SetRange("Cod. grupo", 'COMPRAS');
        if Emp.FindFirst() then
            repeat
                if ColumnNo = 0 Then
                    ColumnNo := 1 else
                    ColumnNo += 2;
                MatrixRec.SetRange("Código Proveedor", Rec."No.");
                MatrixRec.SetRange(Empresa, Emp.Empresa);
                if MatrixRec.FindFirst() Then
                    repeat

                        if (ColumnNo > ColumnOffset) and (ColumnNo - ColumnOffset <= ArrayLen(ColumnValues)) then begin
                            ColumnValues[ColumnNo - ColumnOffset] := MatrixRec.Saldo;
                            ColumnCaptions[ColumnNo - ColumnOffset] := Emp.Empresa;

                        end;
                    until MatrixRec.Next() = 0;

            Until Emp.Next = 0;
        NoOfColumns := ColumnNo + 1;
        ColumnNo2 := 2;
        Repeat
            //for ColumnNo2 := 2 to ColumnNo + 1  do begin
            if Not IxE.Get('TOTAL', ColumnCaptions[ColumnNo2 - 1 - ColumnOffset]) Then IxE.Init();
            if IxE.Saldo <> 0 Then
                ColumnValues[ColumnNo2 - ColumnOffset] := ColumnValues[ColumnNo2 - 1 - ColumnOffset] / IxE.Saldo * 100;
            ColumnCaptions[ColumnNo2 - ColumnOffset] := '%';
            ColumnNo2 += 2;
        Until ColumnNo2 > NoOfColumns;
    end;

    trigger OnOpenPage()
    var
        EmpGr: Record "Empresa grupo";
    begin
        if (Hasta <> 0D) And (Desde <> 0D) Then Rec.SetRange("Date Filter", Desde, Hasta);
        if Not Rec.get('TOTAL') Then begin
            Rec."No." := 'TOTAL';
            Rec.Name := 'TOTAL';
            Rec.Insert();
        end;
        if not EmpGr.Get('COMPRAS', 'Z-TOTAL') Then begin
            EmpGr."Cod. grupo" := 'COMPRAS';
            EmpGr.Empresa := 'Z-TOTAL';
            EmpGr.Insert();
        end;
    end;

    PROCEDURE Calcular(Interc: Code[20]; Contabilidad: Boolean);
    VAR
        IxE: Record "Proveedores x Empresa";
        IxEmpGr: Record "Proveedores x Empresa";
        Contratos: Record 36;
        rAlb: Record 120;
        r121: Record 121;
        rDev: Record 6650;
        r121D: Record 6651;
        r25: Record 25;
        r21: Record 21;
        r380: Record 380;
        r379: Record 379;
        EmpGr: Record "Empresa grupo";
        Pro: Record "Proveedores Compra";
    BEGIN
        a := 0;
        if Rec.GETFILTER("Date Filter") = '' THEN ERROR('Especifique filtro fecha');
        // IxE.SETRANGE(IxE."Código Proveedor", Rec."No.");
        IxE.SetRange("Código Proveedor", 'TOTAL');
        IxE.DeleteAll();
        IxE.Reset();
        IxEmpGr.Reset();
        IxEmpGr.ModifyAll(Saldo, 0);
        IxE.SetRange(Empresa, 'Z-TOTAL');
        IxE.DeleteAll();
        IxE.SetRange(Empresa, 'TOTAL');
        IxE.DeleteAll();
        IxE.Reset();
        Commit();
        Rec.Copyfilter("Date Filter", IxE."Date Filter");
        Ventana.OPEN('#########1## de #########2##');
        Ventana.UPDATE(2, IxE.COUNT);
        if IxE.FINDFIRST THEN
            REPEAT
                a += 1;
                Ventana.UPDATE(1, a);
                IxE.Saldo := 0;
                IxE.Desde := Desde;
                IxE.Hasta := Hasta;

                IxE.Saldo := TotalesDocumentos(IxE.Proveedor, IxE.Empresa, Ixe.Desde, Ixe.Hasta, Contabilidad);
                IxE.MODIFY;
                COMMIT;

            UNTIL IxE.NEXT = 0;
        IxE.Reset();

        if Pro.FindFirst() Then
            repeat
                IxE.SetRange("Código Proveedor", pro."No.");
                if IxE.FindFirst() Then begin
                    IxEmpGr."Código Proveedor" := Pro."No.";
                    IxEmpGr.Empresa := 'Z-TOTAL';
                    IxEmpGr.Saldo := 0;
                    IxEmpGr.Desde := Desde;
                    IxEmpGr.Hasta := Hasta;
                    IxEmpGr.Insert();
                    repeat
                        IxEmpGr.Saldo += IxE.Saldo;
                    until IxE.Next() = 0;
                    IxEmpGr.Modify();
                end;
            until Pro.Next() = 0;
        IxE.Reset();
        EmpGr.SetRange("Cod. grupo", 'COMPRAS');
        if EmpGr.FindFirst() Then
            repeat
                IxE.SetRange(Empresa, EmpGr.Empresa);
                if IxE.FindFirst() Then begin
                    IxEmpGr."Código Proveedor" := 'TOTAL';
                    IxEmpGr.Empresa := EmpGr.Empresa;
                    IxEmpGr.Saldo := 0;
                    IxEmpGr.Desde := Desde;
                    IxEmpGr.Hasta := Hasta;
                    IxEmpGr.Insert();
                    repeat
                        if IxE."Código Proveedor" <> 'TOTAL' then
                            IxEmpGr.Saldo += IxE.Saldo;
                        IxEmpGr.Modify();
                    until IxE.Next() = 0;
                end;
            until EmpGr.Next() = 0;
        Ventana.CLOSE;
        // Sql.Close;
        // CLEAR(Sql);
    END;

    PROCEDURE TotalesDocumentos(No: Code[20]; pEmpresa: Text[30]; Desde: Date; Hasta: Date; Contabilidad: Boolean): Decimal;
    VAR
        CabCompra: Record 38;
        PurchLine: Record 39;
        Importe: Decimal;
        BImporte: Decimal;
        BImpBorFac: Decimal;
        BImpBorAbo: Decimal;
        BImpFac: Decimal;
        BImpAbo: Decimal;
        BTotImp: Decimal;
        BTotCont: Decimal;
        GlEntry: Record 17;
        Albaranes: Record 120;
    BEGIN
        //FCL-31/05/04. Obtengo totales de borradores y facturas correspondientes a este contrato.
        // Contrato.Get(Contrato."Document Type"::Order, No);
        Importe := 0;
        if pEmpresa = 'Z-TOTAL' Then exit(0);

        CabCompra.RESET;
        CabCompra.CHANGECOMPANY(pEmpresa);
        CabCompra.SetRange("Order Date", Desde, Hasta);
        CabCompra.SETRANGE("Buy-From Vendor No.", No);
        CabCompra.SETFILTER("Document Type", '%1|%2',
           CabCompra."Document Type"::Order, CabCompra."Document Type"::"Return Order");
        if CabCompra.FIND('-') THEN BEGIN
            REPEAT
                if Contabilidad then begin
                    Albaranes.CHANGECOMPANY(pEmpresa);
                    Albaranes.SetRange("Order No.", CabCompra."No.");
                    If Albaranes.Find('-') then
                        repeat
                            GlEntry.CHANGECOMPANY(pEmpresa);
                            GlEntry.SETRANGE("Document No.", Albaranes."No.");
                            GlEntry.SETRANGE("Document Type", GlEntry."Document Type"::Receipt);
                            GlEntry.SETRANGE("G/L Account No.", '6', '69999999999');
                            if GlEntry.FIND('-') THEN
                                REPEAT
                                    Importe += GlEntry.Amount;
                                UNTIL GlEntry.NEXT = 0;
                        until Albaranes.Next() = 0;
                end else begin
                    PurchLine.CHANGECOMPANY(pEmpresa);
                    PurchLine.SETRANGE(PurchLine."Document Type", CabCompra."Document Type");
                    PurchLine.SETRANGE(PurchLine."Document No.", CabCompra."No.");

                    if PurchLine.FINDFIRST THEN
                        REPEAT
                            Importe += (PurchLine.Quantity * PurchLine."Direct Unit Cost" * (1 - PurchLine."Line Discount %" / 100));
                        //BImporte += (PurchLine.Quantity * PurchLine."Direct Unit Cost" * (1 - PurchLine."Line Discount %" / 100));
                        UNTIL PurchLine.NEXT = 0;
                end;
            UNTIL CabCompra.NEXT = 0;
        END;

        EXIT(Importe);
    END;

    PROCEDURE CalcularVenta(Interc: Code[20]);
    VAR
        IxE: Record "Proveedores x Empresa";
        IxEmpGr: Record "Proveedores x Empresa";
        SalesLinet: Record 37 temporary;
        Contratos: Record 36;
        rAlb: Record 120;
        r121: Record 121;
        rDev: Record 6650;
        r121D: Record 6651;
        r25: Record 25;
        r21: Record 21;
        r380: Record 380;
        r379: Record 379;
        EmpGr: Record "Empresa grupo";
        Pro: Record "Proveedores Compra";
    BEGIN

        if Rec.GETFILTER("Date Filter") = '' THEN ERROR('Especifique filtro fecha');
        // IxE.SETRANGE(IxE."Código Proveedor", Rec."No.");
        IxE.SetRange("Código Proveedor", 'TOTAL');
        IxE.DeleteAll();
        IxE.Reset();
        IxEmpGr.Reset();
        IxEmpGr.ModifyAll(Saldo, 0);
        a := 0;
        IxE.SetRange(Empresa, 'Z-TOTAL');
        IxE.DeleteAll();
        IxE.SetRange(Empresa, 'TOTAL');
        IxE.DeleteAll();
        IxE.Reset();
        Commit();
        Rec.Copyfilter("Date Filter", IxE."Date Filter");
        Ventana.OPEN('#########1## de #########2##');
        Ventana.UPDATE(2, IxE.COUNT);
        if IxE.FINDFIRST THEN
            REPEAT
                a += 1;
                Ventana.UPDATE(1, a);
                IxE.Saldo := 0;
                IxE.Desde := Desde;
                IxE.Hasta := Hasta;
                DrillDownVenta(IxE.Empresa, ixe."Código Proveedor", IxE.Saldo, false);
                IxE.MODIFY;
                COMMIT;

            UNTIL IxE.NEXT = 0;
        IxE.Reset();
        if Pro.FindFirst() Then
            repeat
                IxE.SetRange("Código Proveedor", pro."No.");
                if IxE.FindFirst() Then begin
                    IxEmpGr."Código Proveedor" := Pro."No.";
                    IxEmpGr.Empresa := 'Z-TOTAL';
                    IxEmpGr.Saldo := 0;
                    IxEmpGr.Desde := Desde;
                    IxEmpGr.Hasta := Hasta;
                    IxEmpGr.Insert();
                    repeat
                        IxEmpGr.Saldo += IxE.Saldo;
                    until IxE.Next() = 0;
                    IxEmpGr.Modify();
                end;
            until Pro.Next() = 0;
        IxE.Reset();
        EmpGr.SetRange("Cod. grupo", 'COMPRAS');
        if EmpGr.FindFirst() Then
            repeat
                IxE.SetRange(Empresa, EmpGr.Empresa);
                if IxE.FindFirst() Then begin
                    IxEmpGr."Código Proveedor" := 'TOTAL';
                    IxEmpGr.Empresa := EmpGr.Empresa;
                    IxEmpGr.Saldo := 0;
                    IxEmpGr.Desde := Desde;
                    IxEmpGr.Hasta := Hasta;
                    IxEmpGr.Insert();
                    repeat
                        if IxE."Código Proveedor" <> 'TOTAL' then
                            IxEmpGr.Saldo += IxE.Saldo;
                        IxEmpGr.Modify();
                    until IxE.Next() = 0;
                end;
            until EmpGr.Next() = 0;
        Ventana.CLOSE;
        // Sql.Close;
        // CLEAR(Sql);
    END;

    PROCEDURE CalcularContabilidad(Interc: Code[20]);
    VAR
        IxE: Record "Proveedores x Empresa";
        IxEmpGr: Record "Proveedores x Empresa";
        MovContabilidad: Record 17;
        rAlb: Record 120;
        r121: Record 121;
        rDev: Record 6650;
        r121D: Record 6651;
        r25: Record 25;
        r21: Record 21;
        r380: Record 380;
        r379: Record 379;
        EmpGr: Record "Empresa grupo";
        Pro: Record "Proveedores Compra";
    BEGIN
        if Rec.GETFILTER("Date Filter") = '' THEN ERROR('Especifique filtro fecha');
        // IxE.SETRANGE(IxE."Código Proveedor", Rec."No.");
        IxE.SetRange("Código Proveedor", 'TOTAL');
        IxE.DeleteAll();
        IxE.Reset();
        IxEmpGr.Reset();
        IxEmpGr.ModifyAll(Saldo, 0);
        a := 0;
        IxE.SetRange(Empresa, 'Z-TOTAL');
        IxE.DeleteAll();
        IxE.SetRange(Empresa, 'TOTAL');
        IxE.DeleteAll();
        IxE.Reset();
        Commit();
        Rec.Copyfilter("Date Filter", IxE."Date Filter");
        Ventana.OPEN('#########1## de #########2##');
        Ventana.UPDATE(2, IxE.COUNT);
        if IxE.FINDFIRST THEN
            REPEAT
                a += 1;
                Ventana.UPDATE(1, a);
                IxE.Saldo := 0;
                IxE.Desde := Desde;
                IxE.Hasta := Hasta;
                IxE.Saldo := TotalesDocumentosContab(IxE.Proveedor, IxE.Empresa, Ixe.Desde, Ixe.Hasta);
                IxE.MODIFY;
                COMMIT;

            UNTIL IxE.NEXT = 0;
        IxE.Reset();
        if Pro.FindFirst() Then
            repeat
                IxE.SetRange("Código Proveedor", pro."No.");
                if IxE.FindFirst() Then begin
                    IxEmpGr."Código Proveedor" := Pro."No.";
                    IxEmpGr.Empresa := 'Z-TOTAL';
                    IxEmpGr.Saldo := 0;
                    IxEmpGr.Desde := Desde;
                    IxEmpGr.Hasta := Hasta;
                    IxEmpGr.Insert();
                    repeat
                        IxEmpGr.Saldo += IxE.Saldo;
                    until IxE.Next() = 0;
                    IxEmpGr.Modify();
                end;
            until Pro.Next() = 0;
        IxE.Reset();
        EmpGr.SetRange("Cod. grupo", 'COMPRAS');
        if EmpGr.FindFirst() Then
            repeat
                IxE.SetRange(Empresa, EmpGr.Empresa);
                if IxE.FindFirst() Then begin
                    IxEmpGr."Código Proveedor" := 'TOTAL';
                    IxEmpGr.Empresa := EmpGr.Empresa;
                    IxEmpGr.Saldo := 0;
                    IxEmpGr.Desde := Desde;
                    IxEmpGr.Hasta := Hasta;
                    IxEmpGr.Insert();
                    repeat
                        if IxE."Código Proveedor" <> 'TOTAL' then
                            IxEmpGr.Saldo += IxE.Saldo;
                        IxEmpGr.Modify();
                    until IxE.Next() = 0;
                end;
            until EmpGr.Next() = 0;
        Ventana.CLOSE;
        // Sql.Close;
        // CLEAR(Sql);
    END;



    PROCEDURE TotalesDocumentosContab(No: Code[20]; pEmpresa: Text[30]; Desde: Date; Hasta: Date): Decimal;
    VAR
        rCabVenta: Record 36;
        Conta: Record 17;
        rCabCompra: Record 38;
        PurchLine: Record 39;
        Importe: Decimal;
        BImporte: Decimal;
        BImpBorFac: Decimal;
        BImpBorAbo: Decimal;
        BImpFac: Decimal;
        BImpAbo: Decimal;
        BTotImp: Decimal;
        BTotCont: Decimal;

    BEGIN
        //FCL-31/05/04. Obtengo totales de borradores y facturas correspondientes a este contrato.
        // Contrato.Get(Contrato."Document Type"::Order, No);
        Importe := 0;
        if pEmpresa = 'Z-TOTAL' Then exit(0);

        Conta.RESET;
        Conta.CHANGECOMPANY(pEmpresa);

        Conta.SetRange("Posting Date", Desde, Hasta);
        Conta.SETRANGE("Source No.", No);
        Conta.SETRange("Source Type", Conta."Source Type"::Vendor);
        Conta.SetRange("G/L Account No.", '6', '69999999999');
        if Conta.FIND('-') THEN
            REPEAT
                Importe += Conta.Amount;
            UNTIL Conta.NEXT = 0;

        EXIT(Importe);
    END;

    PROCEDURE Desplegar(VAR Inter: Record "Proveedores Compra");
    VAR
        rEmpresa: Record Company;
        Vendor: Record 23;
        RxE: Record "Proveedores x Empresa";
        Cli: Code[20];
        Prov: Code[20];
        Control: Codeunit ControlProcesos;
    BEGIN
        //rEmpresa.SETFILTER(rEmpresa."Clave Recursos",'<>%1','');
        rEmpresa.SetRange("Evaluation Company", false);
        if rEmpresa.FINDFIRST THEN
            REPEAT
                if Control.Permiso_Empresas(rEmpresa.Name) then begin


                    Vendor.CHANGECOMPANY(rEmpresa.Name);

                    Vendor.SETRANGE("VAT Registration No.", Inter."VAT Registration No.");
                    if NOT Vendor.FINDFIRST THEN Prov := '' ELSE Prov := Vendor."No.";
                    if RxE.GET(Inter."No.", rEmpresa.Name) THEN RxE.DELETE;
                    RxE.INIT;
                    RxE.Empresa := rEmpresa.Name;
                    RxE."Código proveedor" := Inter."No.";
                    RxE.Proveedor := Prov;
                    if (Prov <> '') THEN
                        RxE.INSERT;


                end;
            UNTIL rEmpresa.NEXT = 0;
    END;

    PROCEDURE DrillDown(a: Integer)
    VAR
        CabCompra: Record 38;
        Prov: Record "Proveedores x Empresa";
        GlEntry: Record 17;
        Albaranes: Record 120;
        Impo: Decimal;
    BEGIN
        // r17.SETCURRENTKEY(
        // "G/L Account No.", "Source Type", "Source No.", "Posting Date", "Document Type", "Gen. Posting Type");
        case TipoEstadistica of
            TipoEstadistica::Venta:
                DrillDownVenta(ColumnCaptions[a], Rec."No.", Impo, true);
            TipoEstadistica::Contabiliad:
                DrillDownContable(a);
            TipoEstadistica::Compra:
                begin
                    CabCompra.ChangeCompany(ColumnCaptions[a]);
                    Prov.SetRange("Código Proveedor", Rec."No.");
                    Prov.SetRange(Empresa, ColumnCaptions[a]);
                    if Not Prov.FindFirst() then Prov.Init();
                    CabCompra.SETRANGE("Buy-From Vendor No.", prov.Proveedor);//, ColumnCaptions[a] + '999999');
                    CabCompra.SetRange("Order Date", Prov.Desde, Prov.Hasta);

                    Page.RUNMODAL(0, CabCompra);
                end;
            TipoEstadistica::"Contabilidad Pedidos":
                begin
                    CabCompra.ChangeCompany(ColumnCaptions[a]);
                    Prov.SetRange("Código Proveedor", Rec."No.");
                    Prov.SetRange(Empresa, ColumnCaptions[a]);
                    if Not Prov.FindFirst() then Prov.Init();
                    CabCompra.SETRANGE("Buy-From Vendor No.", prov.Proveedor);//, ColumnCaptions[a] + '999999');
                    CabCompra.SetRange("Order Date", Prov.Desde, Prov.Hasta);
                    If CabCompra.Find('-') then
                        repeat
                            Albaranes.CHANGECOMPANY(ColumnCaptions[a]);
                            Albaranes.SetRange("Order No.", CabCompra."No.");
                            If Albaranes.Find('-') then
                                repeat
                                    GlEntry.CHANGECOMPANY(ColumnCaptions[a]);
                                    GlEntry.SETRANGE("Document No.", Albaranes."No.");
                                    GlEntry.SETRANGE("Document Type", GlEntry."Document Type"::Receipt);
                                    GlEntry.SETRANGE("G/L Account No.", '6', '69999999999');
                                    if GlEntry.FIND('-') THEN
                                        REPEAT
                                            Impo += GlEntry.Amount;
                                        UNTIL GlEntry.NEXT = 0;
                                until Albaranes.Next() = 0;
                        until CabCompra.Next() = 0;
                end;
        End;


    end;

    PROCEDURE DrillDownTodos(Contabilidad: Boolean)
    VAR
        CabCompra: Record 38;
        LinCompra: Record 39;
        LinCompraT: Record 39 temporary;
        CabCompraT: Record 38 temporary;
        Prov: Record "Proveedores x Empresa";
        Impo: Decimal;
        ProveedoresCompra: Record "Proveedores Compra";
        Empresa: Record Company;
        Line: Integer;
        Albaranes: Record 120;
        GlEntry: Record 17;
        GlentryT: Record 17 temporary;
        Res: Record Resource;
        Job: Record Job;
    BEGIN
        LinCompraT.DeleteAll();
        ProveedoresCompra.RESET;
        if ProveedoresCompra.FindFirst() Then
            repeat
                Prov.SetRange("Código Proveedor", ProveedoresCompra."No.");
                if Prov.FindFirst() then
                    repeat
                        if Empresa.get(Prov.Empresa) then begin
                            CabCompra.ChangeCompany(Prov.Empresa);
                            Job.ChangeCompany(Prov.Empresa);
                            Prov.SetRange("Código Proveedor", ProveedoresCompra."No.");
                            CabCompra.ChangeCompany(Prov.Empresa);
                            CabCompra.SetRange("Document Type", CabCompra."Document Type"::Order, CabCompra."Document Type"::"Return Order");
                            CabCompra.SETRANGE("Buy-From Vendor No.", prov.Proveedor);//, ColumnCaptions[a] + '999999');
                            CabCompra.SetRange("Order Date", Prov.Desde, Prov.Hasta);
                            LinCompra.ChangeCompany(Prov.Empresa);

                            if CabCompra.Find('-') then
                                repeat
                                    Albaranes.CHANGECOMPANY(Prov.Empresa);
                                    Res.ChangeCompany(Prov.Empresa);
                                    If Not Job.Get(CabCompra."Nº Proyecto") Then Job.Init();
                                    Albaranes.SetRange("Order No.", CabCompra."No.");
                                    If Albaranes.Find('-') then
                                        repeat
                                            GlEntry.CHANGECOMPANY(Prov.Empresa);
                                            GlEntry.SETRANGE("Document No.", Albaranes."No.");
                                            GlEntry.SETRANGE("Document Type", GlEntry."Document Type"::Receipt);
                                            GlEntry.SETRANGE("G/L Account No.", '6', '69999999999');
                                            if GlEntry.FIND('-') THEN
                                                REPEAT
                                                    GlentryT := GlEntry;
                                                    GlentryT.Comment := Prov.Empresa;
                                                    if GlEntryT.Insert() then;
                                                UNTIL GlEntry.NEXT = 0;
                                        until Albaranes.Next() = 0;
                                    if Contabilidad then begin
                                    end else begin
                                        LinCompra.SETRANGE("Document Type", CabCompra."Document Type");
                                        LinCompra.SETRANGE("Document No.", CabCompra."No.");
                                        if LinCompra.FindFirst() then
                                            repeat
                                                LinCompraT := LinCompra;
                                                LinCompraT."IC Item Reference No." := ProveedoresCompra."No.";
                                                if Res.Get(LinCompraT."No.") then
                                                    LinCompraT.Description := Res.Name;
                                                LinCompraT."Description 2" := ProveedoresCompra."Name";
                                                LinCompraT."Order Date" := CabCompra."Order Date";
                                                LinCompraT."Planned Receipt Date" := CabCompra."Order Date";
                                                LinCompraT."Fecha inicial recurso" := Job."Starting Date";
                                                LinCompraT."Posting Group" := ProveedoresCompra."VAT Registration No.";
                                                LinCompraT."Empresa Origen" := Prov.Empresa;
                                                LinCompraT."Empresa Venta" := Prov.Empresa;
                                                Line := LinCompraT."Line No.";
                                                If LinCompraT.Quantity <> 0 Then
                                                    repeat
                                                        LinCompraT."Line No." := Line;
                                                        Line += 10000;
                                                    until LinCompraT.Insert();
                                            until LinCompra.NEXT = 0;
                                    end;
                                until CabCompra.Next() = 0;
                        end;
                    until Prov.Next() = 0;
            until ProveedoresCompra.Next() = 0;
        commit;
        if Contabilidad then
            Page.RUNMODAL(0, GlentryT) else
            Page.RUNMODAL(7001182, LinCompraT);


    end;



    procedure DrillDownVenta(Empresa: Text[30]; Provedor: Code[20]; VAR Saldo: Decimal; Mostrar: Boolean);

    begin
        if CalculadoVentas = false then begin
            CalculaVenta();
            CalculadoVentas := true;
        end;
        rLinVentaT.Reset();
        if Empresa <> '' then rlinventat.SetRange("Empresa Venta", Empresa);
        if Provedor <> '' then rlinventat.SetRange("IC Item Reference No.", Provedor);
        if Mostrar then begin
            Page.RUNMODAL(7001183, rLinVentaT);
            exit;
        end;
        if rLinVentaT.FindFirst() then
            repeat
                Saldo += rLinVentaT."Line Amount";
            until rLinVentaT.Next() = 0;
    end;

    PROCEDURE CalculaVenta();
    VAR
        rCabCompra: Record 38;
        rCabVenta: Record 36;
        PurchLine: Record 39;

        SalesLine: Record 37;
        Prov: Record "Proveedores x Empresa";
        ProveedoresCompra: Record "Proveedores Compra";
        Empresa: Record Company;
        Res: Record Resource;
        Job: Record Job;

    BEGIN
        // r17.SETCURRENTKEY(
        // "G/L Account No.", "Source Type", "Source No.", "Order Date", "Document Type", "Gen. Posting Type");
        rLinVentaT.DeleteAll();
        if ProveedoresCompra.FindFirst() Then
            repeat
                Prov.SetRange("Código Proveedor", ProveedoresCompra."No.");
                if Prov.FindFirst() then
                    repeat
                        if Empresa.Get(Prov.Empresa) then begin
                            rCabCompra.RESET;
                            rCabCompra.CHANGECOMPANY(Prov.Empresa);
                            rCabVenta.CHANGECOMPANY(Prov.Empresa);
                            rCabCompra.SetRange("Order Date", Desde, Hasta);
                            rCabCompra.SETRANGE("Buy-From Vendor No.", Prov.Proveedor);
                            rCabCompra.SETFILTER("Document Type", '%1|%2',
                            rCabCompra."Document Type"::Order, rCabCompra."Document Type"::"Return Order");
                            if rCabCompra.FIND('-') THEN BEGIN
                                REPEAT
                                    SalesLine.CHANGECOMPANY(Prov.Empresa);
                                    Res.ChangeCompany(Prov.Empresa);
                                    rCabVenta.CHANGECOMPANY(Prov.Empresa);
                                    SalesLine.SETRANGE(SalesLine."Document Type", rCabVenta."Document Type"::Order);
                                    Job.ChangeCompany(Prov.Empresa);
                                    rCabVenta.SetRange("Nº Proyecto", rCabCompra."Nº Proyecto");
                                    If Job.Get(rCabCompra."Nº Proyecto") Then begin
                                        if Job."Proyecto en empresa origen" <> '' then begin
                                            rCabVenta.SetRange("Nº Proyecto", Job."Proyecto en empresa origen");
                                            SalesLine.ChangeCompany(job."Empresa Origen");
                                            rCabVenta.ChangeCompany(job."Empresa Origen");
                                        end;
                                    end;
                                    rCabVenta.SetRange("Document Type", rCabVenta."Document Type"::Order);
                                    if rCabVenta.FindFirst() then begin
                                        SalesLine.SETRANGE(SalesLine."Document No.", rCabVenta."No.");

                                        if SalesLine.FINDFIRST THEN
                                            REPEAT
                                                PurchLine.CHANGECOMPANY(Prov.Empresa);
                                                PurchLine.SETRANGE("Document Type", rCabCompra."Document Type");
                                                PurchLine.SETRANGE("Document No.", rCabCompra."No.");
                                                if Copystr(SalesLine."No.", 1, 2) in ['MP', 'GR', 'IB', 'MA', 'MN', 'ME']
                                                then
                                                    salesline."No." := Copystr(SalesLine."No.", 3, 20);
                                                PurchLine.SetRange("No.", SalesLine."No.");
                                                If PurchLine.FindFirst() Then begin
                                                    rLinVentaT := SalesLine;
                                                    rLinVentaT := SalesLine;
                                                    rLinVentaT."Document No." := LetraEmpresa(rLinVentaT."Document No.", Prov.Empresa);
                                                    if Res.Get(SalesLine."No.") then
                                                        rLinVentaT.Description := Res.Name;
                                                    rLinVentaT."IC Item Reference No." := ProveedoresCompra."No.";
                                                    rLinVentaT."Planned Delivery Date" := rCabcompra."Order Date";
                                                    rlinVentat."Shipment Date" := rCabCompra."Fecha Firma";
                                                    rLinVentaT."Originally Ordered No." := rCabCompra."No.";
                                                    rLinVentaT."Description 2" := ProveedoresCompra."Name";
                                                    if Job."Starting Date" <> 0D then
                                                        rLinVentaT."Fecha inicial recurso" := Job."Starting Date";
                                                    rLinVentaT."Posting Group" := ProveedoresCompra."VAT Registration No.";
                                                    rLinVentaT."Empresa Venta" := Prov.Empresa;
                                                    rLinVentaT."Line Amount" := (SalesLine.Quantity * SalesLine."Unit Price" * (1 - SalesLine."Line Discount %" / 100));
                                                    If rLinVentaT.Quantity <> 0 Then
                                                        if rLinventat.Insert() then;
                                                end;

                                            //BImporte += (SalesLine.Quantity * SalesLine."Direct Unit Cost" * (1 - SalesLine."Line Discount %" / 100));
                                            UNTIL SalesLine.NEXT = 0;
                                    end;
                                UNTIL rCabCompra.NEXT = 0;
                            END;
                        end;
                    until Prov.Next() = 0;
            Until ProveedoresCompra.Next() = 0;
        commit;
        rLinVentaT.Reset();

    END;

    procedure LetraEmpresa(No: Code[20]; Empresa: Text[30]): Code[20]
    var
        Inf: Record "Company Information";
        letra: Code[20];
    begin
        inf.ChangeCompany(Empresa);
        Inf.Get();
        letra := Inf."Clave Recursos";
        if letra <> CopyStr(No, 1, 2) then
            No := letra + No;
        exit(No);
    end;

    local procedure DrillDownContableTodos()
    var
        Conta: Record 17;
        Contat: Record 17 temporary;
        Prov: Record "Proveedores x Empresa";
        ProveedoresCompra: Record "Proveedores Compra";
        Empresa: Record Company;
        EntryNo: Integer;
    begin
        EntryNo := 0;
        Contat.DeleteAll();
        if ProveedoresCompra.FindFirst() Then
            repeat
                Prov.SetRange("Código Proveedor", ProveedoresCompra."No.");
                if Prov.FindFirst() then
                    repeat
                        if Empresa.Get(Prov.Empresa) then begin
                            Conta.RESET;
                            Conta.CHANGECOMPANY(Prov.Empresa);

                            Conta.SetRange("Posting Date", Desde, Hasta);
                            Conta.SETRANGE("Source No.", prov.Proveedor);
                            Conta.SETRange("Source Type", Conta."Source Type"::Vendor);
                            Conta.SetRange("G/L Account No.", '6', '69999999999');
                            if Conta.FIND('-') THEN
                                REPEAT
                                    Contat := Conta;
                                    EntryNo += 1;
                                    Contat."Entry No." := EntryNo;
                                    Contat.Comment := prov.Empresa;
                                    Contat.Insert();
                                UNTIL Conta.NEXT = 0;
                        end;
                    until Prov.Next() = 0;
            Until ProveedoresCompra.Next() = 0;
        Commit();
        Page.RUNMODAL(0, Contat);
    end;

    local procedure DrillDownContable(a: Integer)
    var
        Conta: Record 17;
        Contat: Record 17 temporary;
        Prov: Record "Proveedores x Empresa";

        Empresa: Record Company;
        EntryNo: Integer;
    begin
        EntryNo := 0;
        Contat.DeleteAll();

        Prov.SetRange("Código Proveedor", Rec."No.");
        Prov.SetRange(Empresa, ColumnCaptions[a]);
        if Prov.FindFirst() then begin

            if Empresa.Get(Prov.Empresa) then begin
                Conta.RESET;
                Conta.CHANGECOMPANY(Prov.Empresa);

                Conta.SetRange("Posting Date", Desde, Hasta);
                Conta.SETRANGE("Source No.", prov.Proveedor);
                Conta.SETRange("Source Type", Conta."Source Type"::Vendor);
                Conta.SetRange("G/L Account No.", '6', '69999999999');
                if Conta.FIND('-') THEN
                    REPEAT
                        Contat := Conta;
                        EntryNo += 1;
                        Contat."Entry No." := EntryNo;
                        Contat.Comment := prov.Empresa;
                        Contat.Insert();
                    UNTIL Conta.NEXT = 0;
            end;
        End;
        Commit();
        Page.RUNMODAL(0, Contat);
    end;
}
page 7001182 "Lineas Compra"
{
    ApplicationArea = All;
    Editable = false;
    PageType = List;
    LinksAllowed = false;
    SourceTable = "Purchase Line";


    layout
    {
        area(content)
        {
            repeater(Control1)
            {
                ShowCaption = false;
                field("Tipo Documento"; 'Compra')
                {
                    Caption = 'Tipo Documento';
                    ApplicationArea = All;
                    ToolTip = 'Specifies the type of document that you are about to create.';
                }
                field("Document No."; Rec."Document No.")
                {
                    Caption = 'Nº Documento';
                    ApplicationArea = All;
                    ToolTip = 'Specifies the document number.';
                }
                field("Buy-from Vendor No."; Rec."Buy-from Vendor No.")
                {
                    Caption = 'Compra-a Nº Proveedor';
                    ApplicationArea = All;
                    ToolTip = 'Specifies the name of the vendor who delivered the items.';
                }
                field(BuyfromVendorName; Rec."Description 2")
                {
                    Caption = 'Compra-a Nombre Proveedor';
                    ApplicationArea = All;
                    ToolTip = 'Specifies the name of the vendor who delivered the items.';
                }
                field(Empresa; Rec."Empresa Venta")
                {
                    Caption = 'Empresa';
                    ApplicationArea = All;
                    ToolTip = 'Specifies the name of the vendor who delivered the items.';
                }
                field("Line No."; Rec."Line No.")
                {
                    ApplicationArea = All;
                    ToolTip = 'Specifies the line''s number.';
                    Visible = false;
                }
                field(Type; Rec.Type)
                {
                    ApplicationArea = All;
                    ToolTip = 'Specifies the line type.';
                }
                field("No."; Rec."No.")
                {
                    ApplicationArea = All;
                    ToolTip = 'Specifies the number of the involved entry or record, according to the specified number series.';
                }
                field("Variant Code"; Rec."Variant Code")
                {
                    ApplicationArea = Planning;
                    ToolTip = 'Specifies the variant of the item on the line.';
                    Visible = false;
                }
                field(Description; Rec.Description)
                {
                    ApplicationArea = All;
                    ToolTip = 'Specifies a description of the entry of the product to be purchased. To add a non-transactional text line, fill in the Description field only.';
                }

                field("Cif"; Rec."Posting Group")
                {
                    Caption = 'Cif';
                    ApplicationArea = All;
                    ToolTip = 'Specifies the code for the location where the items on the line will be located.';
                    Visible = true;
                }
                field("Fecha inico proyecto"; Rec."Fecha inicial recurso")
                {
                    Caption = 'Fecha inico proyecto';
                    ApplicationArea = All;
                    ToolTip = 'Specifies the code for the location where the items on the line will be located.';
                    Visible = true;
                }
                field("Order Date"; Rec."Order Date")
                {
                    ApplicationArea = All;
                    ToolTip = 'Specifies the code for the location where the items on the line will be located.';
                    Visible = true;
                }
                field("Planned Receipt Date"; Rec."Planned Receipt Date")
                {
                    ApplicationArea = All;
                    Caption = 'Fecha lanzamiento Pedido';
                    ToolTip = 'Specifies the code for the location where the items on the line will be located.';
                    Visible = true;
                }
                field(Quantity; Rec.Quantity)
                {
                    ApplicationArea = All;
                    ToolTip = 'Specifies the number of units of the item specified on the line.';
                }
                field("Reserved Qty. (Base)"; Rec."Reserved Qty. (Base)")
                {
                    ApplicationArea = Reservation;
                    ToolTip = 'Specifies the value in the Reserved Quantity field, expressed in the base unit of measure.';
                }
                field("Unit of Measure Code"; Rec."Unit of Measure Code")
                {
                    ApplicationArea = All;
                    ToolTip = 'Specifies how each unit of the item or resource is measured, such as in pieces or hours. By default, the value in the Base Unit of Measure field on the item or resource card is inserted.';
                }
                field("Direct Unit Cost"; Rec."Direct Unit Cost")
                {
                    ApplicationArea = All;
                    ToolTip = 'Specifies the cost of one unit of the selected item or resource.';
                }
                field("Indirect Cost %"; Rec."Indirect Cost %")
                {
                    ApplicationArea = All;
                    ToolTip = 'Specifies the percentage of the item''s last purchase cost that includes indirect costs, such as freight that is associated with the purchase of the item.';
                    Visible = false;
                }
                field("Unit Cost (LCY)"; Rec."Unit Cost (LCY)")
                {
                    ApplicationArea = All;
                    ToolTip = 'Specifies the cost, in LCY, of one unit of the item or resource on the line.';
                    Visible = false;
                }
                field("Unit Price (LCY)"; Rec."Unit Price (LCY)")
                {
                    ApplicationArea = All;
                    ToolTip = 'Specifies the price, in LCY, of one unit of the item or resource. You can enter a price manually or have it entered according to the Price/Profit Calculation field on the related card.';
                    Visible = false;
                }
                field("Line Amount"; Rec."Line Amount")
                {
                    ApplicationArea = All;
                    BlankZero = true;
                    ToolTip = 'Specifies the net amount, excluding any invoice discount amount, that must be paid for products on the line.';
                }
                field("Job No."; Rec."Job No.")
                {
                    ApplicationArea = Jobs;
                    ToolTip = 'Specifies the number of the related job. If you fill in this field and the Job Task No. field, then a job ledger entry will be posted together with the purchase line.';
                    Visible = false;
                }
                field("Job Task No."; Rec."Job Task No.")
                {
                    ApplicationArea = Jobs;
                    ToolTip = 'Specifies the number of the related job task.';
                    Visible = false;
                }
                field("Job Line Type"; Rec."Job Line Type")
                {
                    ApplicationArea = Jobs;
                    ToolTip = 'Specifies a Job Planning Line together with the posting of a job ledger entry.';
                    Visible = false;
                }
                field("Shortcut Dimension 1 Code"; Rec."Shortcut Dimension 1 Code")
                {
                    ApplicationArea = Dimensions;
                    ToolTip = 'Specifies the code for Shortcut Dimension 1, which is one of two global dimension codes that you set up in the General Ledger Setup window.';
                    Visible = false;
                }
                field("Shortcut Dimension 2 Code"; Rec."Shortcut Dimension 2 Code")
                {
                    ApplicationArea = Dimensions;
                    ToolTip = 'Specifies the code for Shortcut Dimension 2, which is one of two global dimension codes that you set up in the General Ledger Setup window.';
                    Visible = false;
                }
                field("ShortcutDimCode[3]"; ShortcutDimCode[3])
                {
                    ApplicationArea = Dimensions;
                    CaptionClass = '1,2,3';
                    TableRelation = "Dimension Value".Code WHERE("Global Dimension No." = CONST(3),
                                                                  "Dimension Value Type" = CONST(Standard),
                                                                  Blocked = CONST(false));
                    Visible = false;
                }
                field("ShortcutDimCode[4]"; ShortcutDimCode[4])
                {
                    ApplicationArea = Dimensions;
                    CaptionClass = '1,2,4';
                    TableRelation = "Dimension Value".Code WHERE("Global Dimension No." = CONST(4),
                                                                  "Dimension Value Type" = CONST(Standard),
                                                                  Blocked = CONST(false));
                    Visible = false;
                }
                field("ShortcutDimCode[5]"; ShortcutDimCode[5])
                {
                    ApplicationArea = Dimensions;
                    CaptionClass = '1,2,5';
                    TableRelation = "Dimension Value".Code WHERE("Global Dimension No." = CONST(5),
                                                                  "Dimension Value Type" = CONST(Standard),
                                                                  Blocked = CONST(false));
                    Visible = false;
                }
                field("ShortcutDimCode[6]"; ShortcutDimCode[6])
                {
                    ApplicationArea = Dimensions;
                    CaptionClass = '1,2,6';
                    TableRelation = "Dimension Value".Code WHERE("Global Dimension No." = CONST(6),
                                                                  "Dimension Value Type" = CONST(Standard),
                                                                  Blocked = CONST(false));
                    Visible = false;
                }
                field("ShortcutDimCode[7]"; ShortcutDimCode[7])
                {
                    ApplicationArea = Dimensions;
                    CaptionClass = '1,2,7';
                    TableRelation = "Dimension Value".Code WHERE("Global Dimension No." = CONST(7),
                                                                  "Dimension Value Type" = CONST(Standard),
                                                                  Blocked = CONST(false));
                    Visible = false;
                }
                field("ShortcutDimCode[8]"; ShortcutDimCode[8])
                {
                    ApplicationArea = Dimensions;
                    CaptionClass = '1,2,8';
                    TableRelation = "Dimension Value".Code WHERE("Global Dimension No." = CONST(8),
                                                                  "Dimension Value Type" = CONST(Standard),
                                                                  Blocked = CONST(false));
                    Visible = false;
                }
                field("Expected Receipt Date"; Rec."Expected Receipt Date")
                {
                    Visible = false;
                    ApplicationArea = All;
                    ToolTip = 'Specifies the date that you expect the items to be available in your warehouse.';
                }
                field("Outstanding Quantity"; Rec."Outstanding Quantity")
                {
                    ApplicationArea = All;
                    ToolTip = 'Specifies how many units on the order line have not yet been received.';
                }
                field("Outstanding Amount (LCY)"; Rec."Outstanding Amount (LCY)")
                {
                    ApplicationArea = All;
                    ToolTip = 'Specifies the amount for the items on the order that have not yet been received in LCY.';
                    Visible = false;
                }
                field("Amt. Rcd. Not Invoiced (LCY)"; Rec."Amt. Rcd. Not Invoiced (LCY)")
                {
                    ApplicationArea = All;
                    ToolTip = 'Specifies the sum, in LCY, for items that have been received but have not yet been invoiced. The value in the Amt. Rcd. Not Invoiced (LCY) field is used for entries in the Purchase Line table of document type Order to calculate and update the contents of this field.';
                    Visible = false;
                }
            }
        }
    }

    actions
    {
        area(navigation)
        {
            group("&Line")
            {
                Caption = '&Line';
                Image = Line;
                action("Show Document")
                {
                    ApplicationArea = All;
                    Caption = 'Mostrar Documento';
                    Image = View;
                    ShortCutKey = 'Shift+F7';
                    ToolTip = 'Open the document that the selected line exists on.';

                    trigger OnAction()
                    var
                        PageManagement: Codeunit "Page Management";
                    begin
                        PurchHeader.Get(Rec."Document Type", Rec."Document No.");
                        PageManagement.PageRun(PurchHeader);

                        OnShowDocumentOnAfterOnAction(PurchHeader);
                    end;
                }



            }
        }
        area(Promoted)
        {
            group(Category_Process)
            {
                Caption = 'Process', Comment = 'Generated from the PromotedActionCategories property index 1.';

                actionref("Show Document_Promoted"; "Show Document")
                {
                }

            }
        }
    }

    trigger OnAfterGetRecord()
    begin
        Rec.ShowShortcutDimCode(ShortcutDimCode);
    end;

    trigger OnNewRecord(BelowxRec: Boolean)
    begin
        Clear(ShortcutDimCode);
    end;

    trigger OnOpenPage()
    begin
        DetachLinesVisible := Rec.GetFilter("Attached to Line No.") <> '';
    end;

    var
        PurchHeader: Record "Purchase Header";
        DetachLinesVisible: Boolean;

    protected var
        ShortcutDimCode: array[8] of Code[20];

    [IntegrationEvent(true, false)]
    local procedure OnShowDocumentOnAfterOnAction(var PurchHeader: Record "Purchase Header")
    begin
    end;
}
page 7001183 "Lineas Venta"
{
    ApplicationArea = All;
    Editable = false;
    PageType = List;
    LinksAllowed = false;
    SourceTable = "Sales Line";


    layout
    {
        area(content)
        {
            repeater(Control1)
            {
                ShowCaption = false;
                field("Tipo Cocumento"; 'Venta')
                {
                    Caption = 'Tipo Documento';
                    ApplicationArea = All;
                    ToolTip = 'Specifies the type of document that you are about to create.';
                }
                field("Documento Compra"; Rec."Originally Ordered No.")
                {
                    Caption = 'Nº Documento Compra';
                    ApplicationArea = All;
                    ToolTip = 'Specifies the name of the vendor who delivered the items.';
                }
                field("Buy-from Vendor No."; Rec."Item Reference No.")
                {
                    Caption = 'Compra-a Nº Proveedor';
                    ApplicationArea = All;
                    ToolTip = 'Specifies the name of the vendor who delivered the items.';
                }
                field(BuyfromVendorName; Rec."Description 2")
                {
                    Caption = 'Compra-a Nombre Proveedor';
                    ApplicationArea = All;
                    ToolTip = 'Specifies the name of the vendor who delivered the items.';
                }

                field("Contrato"; Rec."Document No.")
                {
                    Caption = 'Contrato';
                    ApplicationArea = All;
                    ToolTip = 'Specifies the document number.';
                }
                field(Empresa; Rec."Empresa Venta")
                {
                    Caption = 'Empresa';
                    ApplicationArea = All;
                    ToolTip = 'Specifies the name of the vendor who delivered the items.';
                }
                field("Line No."; Rec."Line No.")
                {
                    ApplicationArea = All;
                    ToolTip = 'Specifies the line''s number.';
                    Visible = false;
                }
                field(Type; Rec.Type)
                {
                    ApplicationArea = All;
                    ToolTip = 'Specifies the line type.';
                }
                field("No."; Rec."No.")
                {
                    ApplicationArea = All;
                    ToolTip = 'Specifies the number of the involved entry or record, according to the specified number series.';
                }
                field("Variant Code"; Rec."Variant Code")
                {
                    ApplicationArea = Planning;
                    ToolTip = 'Specifies the variant of the item on the line.';
                    Visible = false;
                }
                field(Description; Rec.Description)
                {
                    ApplicationArea = All;
                    ToolTip = 'Specifies a description of the entry of the product to be purchased. To add a non-transactional text line, fill in the Description field only.';
                }

                field("Cif"; Rec."Posting Group")
                {
                    Caption = 'Cif';
                    ApplicationArea = All;
                    ToolTip = 'Specifies the code for the location where the items on the line will be located.';
                    Visible = true;
                }
                field("Fecha inico proyecto"; Rec."Fecha inicial recurso")
                {
                    Caption = 'Fecha inico proyecto';
                    ApplicationArea = All;
                    ToolTip = 'Specifies the code for the location where the items on the line will be located.';
                    Visible = true;
                }
                field("Order Date"; Rec."Planned Delivery Date")
                {
                    Caption = 'Fecha Pedido';
                    ApplicationArea = All;
                    ToolTip = 'Specifies the code for the location where the items on the line will be located.';
                    Visible = true;
                }
                field("Planned Receipt Date"; Rec."Shipment Date")
                {
                    ApplicationArea = All;
                    Caption = 'Fecha Firma';
                    ToolTip = 'Specifies the code for the location where the items on the line will be located.';
                    Visible = true;
                }
                field(Quantity; Rec.Quantity)
                {
                    ApplicationArea = All;
                    ToolTip = 'Specifies the number of units of the item specified on the line.';
                }
                field("Reserved Qty. (Base)"; Rec."Reserved Qty. (Base)")
                {
                    ApplicationArea = Reservation;
                    ToolTip = 'Specifies the value in the Reserved Quantity field, expressed in the base unit of measure.';
                }
                field("Unit of Measure Code"; Rec."Unit of Measure Code")
                {
                    ApplicationArea = All;
                    ToolTip = 'Specifies how each unit of the item or resource is measured, such as in pieces or hours. By default, the value in the Base Unit of Measure field on the item or resource card is inserted.';
                }
                field("Unit price"; -Rec."Unit Price")
                {
                    ApplicationArea = All;
                    ToolTip = 'Specifies the cost of one unit of the selected item or resource.';
                }

                field("Unit Cost (LCY)"; Rec."Unit Cost (LCY)")
                {
                    ApplicationArea = All;
                    ToolTip = 'Specifies the cost, in LCY, of one unit of the item or resource on the line.';
                    Visible = false;
                }

                field("Line Amount"; -Rec."Line Amount")
                {
                    ApplicationArea = All;
                    BlankZero = true;
                    ToolTip = 'Specifies the net amount, excluding any invoice discount amount, that must be paid for products on the line.';
                }
                field("Job No."; Rec."Job No.")
                {
                    ApplicationArea = Jobs;
                    ToolTip = 'Specifies the number of the related job. If you fill in this field and the Job Task No. field, then a job ledger entry will be posted together with the purchase line.';
                    Visible = false;
                }
                field("Job Task No."; Rec."Job Task No.")
                {
                    ApplicationArea = Jobs;
                    ToolTip = 'Specifies the number of the related job task.';
                    Visible = false;
                }

                field("Shortcut Dimension 1 Code"; Rec."Shortcut Dimension 1 Code")
                {
                    ApplicationArea = Dimensions;
                    ToolTip = 'Specifies the code for Shortcut Dimension 1, which is one of two global dimension codes that you set up in the General Ledger Setup window.';
                    Visible = false;
                }
                field("Shortcut Dimension 2 Code"; Rec."Shortcut Dimension 2 Code")
                {
                    ApplicationArea = Dimensions;
                    ToolTip = 'Specifies the code for Shortcut Dimension 2, which is one of two global dimension codes that you set up in the General Ledger Setup window.';
                    Visible = false;
                }
                field("ShortcutDimCode[3]"; ShortcutDimCode[3])
                {
                    ApplicationArea = Dimensions;
                    CaptionClass = '1,2,3';
                    TableRelation = "Dimension Value".Code WHERE("Global Dimension No." = CONST(3),
                                                                  "Dimension Value Type" = CONST(Standard),
                                                                  Blocked = CONST(false));
                    Visible = false;
                }
                field("ShortcutDimCode[4]"; ShortcutDimCode[4])
                {
                    ApplicationArea = Dimensions;
                    CaptionClass = '1,2,4';
                    TableRelation = "Dimension Value".Code WHERE("Global Dimension No." = CONST(4),
                                                                  "Dimension Value Type" = CONST(Standard),
                                                                  Blocked = CONST(false));
                    Visible = false;
                }
                field("ShortcutDimCode[5]"; ShortcutDimCode[5])
                {
                    ApplicationArea = Dimensions;
                    CaptionClass = '1,2,5';
                    TableRelation = "Dimension Value".Code WHERE("Global Dimension No." = CONST(5),
                                                                  "Dimension Value Type" = CONST(Standard),
                                                                  Blocked = CONST(false));
                    Visible = false;
                }
                field("ShortcutDimCode[6]"; ShortcutDimCode[6])
                {
                    ApplicationArea = Dimensions;
                    CaptionClass = '1,2,6';
                    TableRelation = "Dimension Value".Code WHERE("Global Dimension No." = CONST(6),
                                                                  "Dimension Value Type" = CONST(Standard),
                                                                  Blocked = CONST(false));
                    Visible = false;
                }
                field("ShortcutDimCode[7]"; ShortcutDimCode[7])
                {
                    ApplicationArea = Dimensions;
                    CaptionClass = '1,2,7';
                    TableRelation = "Dimension Value".Code WHERE("Global Dimension No." = CONST(7),
                                                                  "Dimension Value Type" = CONST(Standard),
                                                                  Blocked = CONST(false));
                    Visible = false;
                }
                field("ShortcutDimCode[8]"; ShortcutDimCode[8])
                {
                    ApplicationArea = Dimensions;
                    CaptionClass = '1,2,8';
                    TableRelation = "Dimension Value".Code WHERE("Global Dimension No." = CONST(8),
                                                                  "Dimension Value Type" = CONST(Standard),
                                                                  Blocked = CONST(false));
                    Visible = false;
                }
                field("Outstanding Quantity"; Rec."Outstanding Quantity")
                {
                    ApplicationArea = All;
                    ToolTip = 'Specifies how many units on the order line have not yet been received.';
                }
                field("Outstanding Amount (LCY)"; Rec."Outstanding Amount (LCY)")
                {
                    ApplicationArea = All;
                    ToolTip = 'Specifies the amount for the items on the order that have not yet been received in LCY.';
                    Visible = false;
                }

            }
        }
    }

    actions
    {
        area(navigation)
        {
            group("&Line")
            {
                Caption = '&Line';
                Image = Line;
                action("Show Document")
                {
                    ApplicationArea = All;
                    Caption = 'Mostrar Documento';
                    Image = View;
                    ShortCutKey = 'Shift+F7';
                    ToolTip = 'Open the document that the selected line exists on.';

                    trigger OnAction()
                    var
                        PageManagement: Codeunit "Page Management";
                    begin
                        SalesHeader.Get(Rec."Document Type", Rec."Document No.");
                        PageManagement.PageRun(SalesHeader);

                        OnShowDocumentOnAfterOnAction(SalesHeader);
                    end;
                }



            }
        }
        area(Promoted)
        {
            group(Category_Process)
            {
                Caption = 'Process', Comment = 'Generated from the PromotedActionCategories property index 1.';

                actionref("Show Document_Promoted"; "Show Document")
                {
                }

            }
        }
    }

    trigger OnAfterGetRecord()
    begin
        Rec.ShowShortcutDimCode(ShortcutDimCode);
    end;

    trigger OnNewRecord(BelowxRec: Boolean)
    begin
        Clear(ShortcutDimCode);
    end;

    trigger OnOpenPage()
    begin
        DetachLinesVisible := Rec.GetFilter("Attached to Line No.") <> '';
    end;

    var
        SalesHeader: Record "Sales Header";
        DetachLinesVisible: Boolean;

    protected var
        ShortcutDimCode: array[8] of Code[20];

    [IntegrationEvent(true, false)]
    local procedure OnShowDocumentOnAfterOnAction(var PurchHeader: Record "Sales Header")
    begin
    end;
}