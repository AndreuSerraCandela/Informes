pageextension 93002 Intercambioext extends "Ficha Intercambio"
{
    actions
    {
        addafter("&Calcular")
        {
            action("Exportar a Excel")
            {
                ApplicationArea = All;
                Image = Excel;

                trigger OnAction()
                Var
                    RowNo: Integer;
                    ColumnNo: Integer;
                    rEmp: Record "Company Information";
                    //rEmp2 : Record 2000000006;
                    a: Integer;
                    RxE: Record "Intercambio x Empresa";
                    rMovCli: Record "Cust. Ledger Entry";
                    DetMovCli: Record "Detailed Cust. Ledg. Entry";
                    DetMovPro: Record "Detailed Vendor Ledg. Entry";
                    rMovPro: Record "Vendor Ledger Entry";
                    ExcelStream: OutStream;
                    Secuencia: Integer;
                    ficheros: Record Ficheros;
                    Intstream: Instream;

                    Saldo: Decimal;
                    SaldoLinea: Decimal;
                    SaldoCli: Decimal;
                    SaldoPro: Decimal;
                BEGIN
                    RxE.SETRANGE(RxE."Código Intercambio", Rec."No.");
                    EnterCell(1, 1, COMPANYNAME, true, FALSE, false, false, '', TempExcelBuffer."Cell Type"::Text, '', 0, '', '');
                    //        RowNo, ColumnNo, CellValue  , Bold, Italic, UnderLine, DobleUnderLine, NumberFormat, CellType,                         Fuente, Tamaño, Color, ColorFondo
                    EnterCell(
                        2,
                        1,
                        'A la atención de ' + Rec.Name,
                        false,
                        FALSE,
                        FALSE,
                        false,
                        '',
                        TempExcelBuffer."Cell Type"::Text,
                        '',
                        0,
                        '',
                        '');
                    EnterCell(
                        3,
                        1,
                        'Palma de Mallorca a ' + Format(Today, 0, '<Day,2> de <Month text> de <Year4>'),
                        false,
                        FALSE,
                        FALSE,
                        false,
                        '',
                        TempExcelBuffer."Cell Type"::Text,
                        '',
                        0,
                        '',
                        '');
                    EnterCell(
                        4,
                        1,
                        'Según nuestro acuerdo de intercambio. a fecha de hoy, hemos compensado',
                        false,
                        FALSE,
                        FALSE,
                        false,
                        '',
                        TempExcelBuffer."Cell Type"::Text,
                        '',
                        0,
                        '',
                        '');
                    EnterCell(
                        5,
                        1,
                        'la/s siguiente/s factura/s:',
                        false,
                        FALSE,
                        FALSE,
                        false,
                        '',
                        TempExcelBuffer."Cell Type"::Text,
                        '',
                        0,
                        '',
                        '');
                    RowNo := 1;
                    If RxE.FINDFIRST THEN
                        REPEAT
                            rMovCli.CHANGECOMPANY(RxE.Empresa);
                            rMovCli.SETRANGE(rMovCli."Customer No.", RxE.Cliente);
                            rMovCli.SETRANGE(rMovCli."Payment Method Code", 'INTERCAM');
                            rMovCli.SETRANGE(Open, TRUE);
                            rEmp.ChangeCompany(RxE.Empresa);
                            rEmp.GET;

                            If rMovCli.FINDSET THEN begin
                                EnterCell(
                                   RowNo + 9,
                                   1,
                                   'Facturas emitidas por ' + rEmp.Name + ' a ' + Rec.Name,
                                   false,
                                   false,
                                   false,
                                false,
                                '',
                                TempExcelBuffer."Cell Type"::Text,
                                '',
                                0,
                                '',
                                '');
                                EnterCell(
                                    RowNo + 12,
                                    1,
                                    'FECHA',
                                    TRUE,
                                    FALSE,
                                    true,
                                false,
                                '',
                                TempExcelBuffer."Cell Type"::Text,
                                '',
                                0,
                                '',
                                '');
                                EnterCell(
                                    RowNo + 12,
                                    2,
                                    'FACTURA',
                                    TRUE,
                                    FALSE,
                                    True,
                                false,
                                '',
                                TempExcelBuffer."Cell Type"::Text,
                                '',
                                0,
                                '',
                                '');
                                EnterCell(
                                    RowNo + 11,
                                    3,
                                    'TOTAL',
                                    TRUE,
                                    FALSE,
                                    false,
                                false,
                                '',
                                TempExcelBuffer."Cell Type"::Text,
                                '',
                                0,
                                '',
                                '');
                                EnterCell(
                                    RowNo + 12,
                                    3,
                                    'FACTURA',
                                    TRUE,
                                    FALSE,
                                    true,
                                false,
                                '',
                                TempExcelBuffer."Cell Type"::Text,
                                '',
                                0,
                                '',
                                '');
                                EnterCell(
                                    RowNo + 10,
                                    4,
                                    'IMPORTE',
                                    TRUE,
                                    FALSE,
                                    false,
                                false,
                                '',
                                TempExcelBuffer."Cell Type"::Text,
                                '',
                                0,
                                '',
                                '');
                                EnterCell(
                                    RowNo + 11,
                                    4,
                                    'PENDIENTE',
                                    TRUE,
                                    FALSE,
                                    false,
                                false,
                                '',
                                TempExcelBuffer."Cell Type"::Text,
                                '',
                                0,
                                '',
                                '');
                                EnterCell(
                                    RowNo + 12,
                                    4,
                                    'LIQUIDAR',
                                    TRUE,
                                    FALSE,
                                    true,
                                false,
                                '',
                                TempExcelBuffer."Cell Type"::Text,
                                '',
                                0,
                                '',
                                '');
                                EnterCell(
                                    RowNo + 11,
                                    5,
                                    'IMPORTE',
                                    TRUE,
                                    FALSE,
                                    false,
                                false,
                                '',
                                TempExcelBuffer."Cell Type"::Text,
                                '',
                                0,
                                '',
                                '');
                                EnterCell(
                                    RowNo + 12,
                                    5,
                                    'LIQUIDADO',
                                    TRUE,
                                    FALSE,
                                    true,
                                false,
                                '',
                                TempExcelBuffer."Cell Type"::Text,
                                '',
                                0,
                                '',
                                '');
                                RowNo += 12;
                                SaldoLinea := 0;
                                repeat
                                    RowNo += 1;
                                    EnterCell(
                                       RowNo + 1,
                                       1,
                                       Format(rMovCli."Posting Date", 0, '<Day,2>/<Month,2>/<Year4>'),
                                       FALSE,
                                       FALSE,
                                       FALSE,
                                        false,
                                        '',
                                        TempExcelBuffer."Cell Type"::Text,
                                        '',
                                        0,
                                        '',
                                        '');
                                    EnterCell(
                                        RowNo + 1,
                                        2,
                                        rMovCli."Document No.",
                                        FALSE,
                                        FALSE,
                                        FALSE,
                                false,
                                '',
                                TempExcelBuffer."Cell Type"::Text,
                                '',
                                0,
                                '',
                                '');
                                    DetMovCli.Reset();
                                    DetMovCli.ChangeCompany(RxE.Empresa);
                                    DetMovCli.SetRange(DetMovCli."Entry Type", DetMovCli."Entry Type"::"Initial Entry");
                                    DetMovCli.SETRANGE("Cust. Ledger Entry No.", rMovCli."Entry No.");
                                    DetMovCli.SetRange("Ledger Entry Amount", true);
                                    DetMovCli.CalcSums("Amount (LCY)");
                                    EnterCell(
                                        RowNo + 1,
                                        3,
                                        Format(DetMovCli."Amount (LCY)", 0),
                                        FALSE,
                                        FALSE,
                                        FALSE,
                                false,
                                '_-* #,##0.00_-;-* #,##0.00_-;_-* "-"_-;_-@_-',
                                TempExcelBuffer."Cell Type"::number,
                                '',
                                0,
                                '',
                                '');
                                    DetMovCli.Reset();
                                    DetMovCli.ChangeCompany(RxE.Empresa);
                                    DetMovCli.SetRange(DetMovCli."Entry Type");
                                    DetMovCli.SETRANGE("Cust. Ledger Entry No.", rMovCli."Entry No.");
                                    DetMovCli.SetRange("Ledger Entry Amount");
                                    DetMovCli.CalcSums("Amount (LCY)");
                                    //'_-* #,##0.00_-;-* #,##0.00_-;_-* "-"_-;_-@_-';
                                    EnterCell(
                                        RowNo + 1,
                                        4,
                                        Format(DetMovCli."Amount (LCY)", 0),
                                        FALSE,
                                        FALSE,
                                        FALSE,
                                false,
                                '_-* #,##0.00_-;-* #,##0.00_-;_-* "-"_-;_-@_-',
                                TempExcelBuffer."Cell Type"::Number,
                                '',
                                0,
                                '',
                                '');
                                    EnterCell(
                                       RowNo + 1,
                                       5,
                                       Format(DetMovCli."Amount (LCY)", 0),
                                       FALSE,
                                       FALSE,
                                       FALSE,
                                false,
                                '_-* #,##0.00_-;-* #,##0.00_-;_-* "-"_-;_-@_-',
                                TempExcelBuffer."Cell Type"::Number,
                                '',
                                0,
                                '',
                                '');
                                    SaldoLinea += DetMovCli."Amount (LCY)";
                                until rMovCli.NEXT = 0;
                                RowNo += 1;
                                EnterCell(
                                   RowNo + 1,
                                   4,
                                   'Subtotal:',
                                   true,
                                   FALSE,
                                   FALSE,
                                false,
                                '',
                                TempExcelBuffer."Cell Type"::Text,
                                '',
                                0,
                                '',
                                '');
                                EnterCell(
                                    RowNo + 1,
                                    5,
                                    Format(SaldoLinea, 0),
                                    true,
                                    FALSE,
                                    FALSE,
                                false,
                                '_-* #,##0.00_-;-* #,##0.00_-;_-* "-"_-;_-@_-',
                                TempExcelBuffer."Cell Type"::Number,
                                '',
                                0,
                                '',
                                '');
                                SaldoCli += SaldoLinea;
                            end;
                            rMovPro.CHANGECOMPANY(RxE.Empresa);
                            rMovPro.SETRANGE(rMovPro."Vendor No.", RxE.Proveedor);
                            rMovPro.SETRANGE(rMovPro."Payment Method Code", 'INTERCAM');
                            rMovPro.SETRANGE(Open, TRUE);
                            rEmp.ChangeCompany(RxE.Empresa);
                            rEmp.GET;
                            If rMovPro.FINDSET THEN begin
                                EnterCell(
                                   RowNo + 9,
                                   1,
                                   'Facturas emitidas por ' + Rec.Name + ' a ' + rEmp.Name,
                                   false,
                                   false,
                                   false,
                                false,
                                '',
                                TempExcelBuffer."Cell Type"::Text,
                                '',
                                0,
                                '',
                                '');
                                EnterCell(
                                    RowNo + 12,
                                    1,
                                    'FECHA',
                                    TRUE,
                                    FALSE,
                                    true,
                                false,
                                '',
                                TempExcelBuffer."Cell Type"::Text,
                                '',
                                0,
                                '',
                                '');
                                EnterCell(
                                    RowNo + 12,
                                    2,
                                    'FACTURA',
                                    TRUE,
                                    FALSE,
                                    True,
                                false,
                                '',
                                TempExcelBuffer."Cell Type"::Text,
                                '',
                                0,
                                '',
                                '');
                                EnterCell(
                                    RowNo + 11,
                                    3,
                                    'TOTAL',
                                    TRUE,
                                    FALSE,
                                    FALSE,
                                false,
                                '',
                                TempExcelBuffer."Cell Type"::Text,
                                '',
                                0,
                                '',
                                '');
                                EnterCell(
                                    RowNo + 12,
                                    3,
                                    'FACTURA',
                                    TRUE,
                                    FALSE,
                                    true,
                                false,
                                '',
                                TempExcelBuffer."Cell Type"::Text,
                                '',
                                0,
                                '',
                                '');
                                EnterCell(
                                    RowNo + 10,
                                    4,
                                    'IMPORTE',
                                    TRUE,
                                    FALSE,
                                    FALSE,
                                false,
                                '',
                                TempExcelBuffer."Cell Type"::Text,
                                '',
                                0,
                                '',
                                '');
                                EnterCell(
                                    RowNo + 11,
                                    4,
                                    'PENDIENTE',
                                    TRUE,
                                    FALSE,
                                    FALSE,
                                false,
                                '',
                                TempExcelBuffer."Cell Type"::Text,
                                '',
                                0,
                                '',
                                '');
                                EnterCell(
                                    RowNo + 12,
                                    4,
                                    'LIQUIDAR',
                                    TRUE,
                                    FALSE,
                                    true,
                                false,
                                '',
                                TempExcelBuffer."Cell Type"::Text,
                                '',
                                0,
                                '',
                                '');
                                EnterCell(
                                    RowNo + 11,
                                    5,
                                    'IMPORTE',
                                    TRUE,
                                    FALSE,
                                    FALSE,
                                false,
                                '',
                                TempExcelBuffer."Cell Type"::Text,
                                '',
                                0,
                                '',
                                '');
                                EnterCell(
                                    RowNo + 12,
                                    5,
                                    'LIQUIDADO',
                                    TRUE,
                                    FALSE,
                                    true,
                                false,
                                '',
                                TempExcelBuffer."Cell Type"::Text,
                                '',
                                0,
                                '',
                                '');
                                RowNo += 12;
                                SaldoLinea := 0;
                                repeat
                                    RowNo += 1;
                                    EnterCell(
                                       RowNo + 1,
                                       1,
                                       Format(rMovPro."Posting Date", 0, '<Day,2>/<Month,2>/<Year4>'),
                                       FALSE,
                                       FALSE,
                                       FALSE,
                                false,
                                '',
                                TempExcelBuffer."Cell Type"::Text,
                                '',
                                0,
                                '',
                                '');
                                    EnterCell(
                                        RowNo + 1,
                                        2,
                                        rMovPro."External Document No.",
                                        FALSE,
                                        FALSE,
                                        FALSE,
                                false,
                                '',
                                TempExcelBuffer."Cell Type"::Text,
                                '',
                                0,
                                '',
                                '');
                                    DetMovPro.Reset();
                                    DetMovPro.ChangeCompany(RxE.Empresa);
                                    DetMovPro.SetRange(DetMovPro."Entry Type", DetMovPro."Entry Type"::"Initial Entry");
                                    DetMovPro.SETRANGE("Vendor Ledger Entry No.", rMovPro."Entry No.");
                                    DetMovPro.SetRange("Ledger Entry Amount", true);
                                    DetMovPro.CalcSums("Amount (LCY)");
                                    EnterCell(
                                        RowNo + 1,
                                        3,
                                        Format(DetMovPro."Amount (LCY)", 0),
                                        FALSE,
                                        FALSE,
                                        FALSE,
                                false,
                                '_-* #,##0.00_-;-* #,##0.00_-;_-* "-"_-;_-@_-',
                                TempExcelBuffer."Cell Type"::Number,
                                '',
                                0,
                                '',
                                '');

                                    DetMovPro.ChangeCompany(RxE.Empresa);
                                    DetMovPro.SetRange(DetMovPro."Entry Type");
                                    DetMovPro.SETRANGE("Vendor Ledger Entry No.", rMovPro."Entry No.");
                                    DetMovPro.SetRange("Ledger Entry Amount");
                                    DetMovPro.CalcSums("Amount (LCY)");
                                    EnterCell(
                                        RowNo + 1,
                                        4,
                                        Format(DetMovPro."Amount (LCY)", 0),
                                        FALSE,
                                        FALSE,
                                        FALSE,
                                false,
                                '_-* #,##0.00_-;-* #,##0.00_-;_-* "-"_-;_-@_-',
                                TempExcelBuffer."Cell Type"::Number,
                                '',
                                0,
                                '',
                                '');
                                    EnterCell(
                                       RowNo + 1,
                                       5,
                                       Format(DetMovPro."Amount (LCY)", 0),
                                       FALSE,
                                       FALSE,
                                       FALSE,
                                false,
                                '_-* #,##0.00_-;-* #,##0.00_-;_-* "-"_-;_-@_-',
                                TempExcelBuffer."Cell Type"::Number,
                                '',
                                0,
                                '',
                                '');
                                    SaldoLinea += DetMovPro."Amount (LCY)";
                                until rMovPro.NEXT = 0;
                                RowNo += 1;
                                EnterCell(
                                    RowNo + 1,
                                    4,
                                    'Subtotal:',
                                    true,
                                    FALSE,
                                    FALSE,
                                    false,
                                    '',
                                    TempExcelBuffer."Cell Type"::Text,
                                    '',
                                    0,
                                    '',
                                    '');
                                EnterCell(
                                    RowNo + 1,
                                    5,
                                    Format(SaldoLinea, 0),
                                    true,
                                    FALSE,
                                    FALSE,
                                    false,
                                    '_-* #,##0.00_-;-* #,##0.00_-;_-* "-"_-;_-@_-',
                                    TempExcelBuffer."Cell Type"::Number,
                                    '',
                                    0,
                                    '',
                                    '');
                                SaldoPro += SaldoLinea;
                            end;
                        UNTIL RxE.NEXT = 0;
                    //Esperamos merezcan su conformidad.
                    EnterCell(
                        RowNo + 2,
                        1,
                        'Esperamos merezcan su conformidad.',
                        FALSE,
                        FALSE,
                        FALSE,
                        FALSE,
                        '',
                        TempExcelBuffer."Cell Type"::Text,
                        '',
                        0,
                        '',
                        '');

                    //Aprovecho para notificarles que el saldo que mantenemos con ustedes a día de hoy, una vez aplicados 
                    EnterCell(
                        RowNo + 3,
                        1,
                        'Aprovecho para notificarles que el saldo que mantenemos con ustedes a día de hoy, una vez aplicados',
                        FALSE,
                        FALSE,
                        FALSE,
                        FALSE,
                        '',
                        TempExcelBuffer."Cell Type"::Text,
                        '',
                        0,
                        '',
                        '');
                    //los intercambios que les hemos detallado arriba, asciende a 3.715,50  € a favor de Malla, S.A.
                    rEmp.ChangeCompany(CompanyName);
                    rEmp.Get();
                    Saldo := SaldoCli + SaldoPro;
                    EnterCell(
                        RowNo + 4,
                        1,
                        'los intercambios que les hemos detallado arriba, asciende a ' + Format(Saldo) + ' € a favor de ' + rEmp.Name + '.',
                        FALSE,
                        FALSE,
                        FALSE,
                        FALSE,
                        '',
                        TempExcelBuffer."Cell Type"::Text,
                        '',
                        0,
                        '',
                        '');
                    ficheros.Reset();

                    If ficheros.FindLast() then Secuencia := ficheros.Secuencia + 1 else Secuencia := 1;
                    ficheros.Secuencia := Secuencia;
                    ficheros."Nombre fichero" := 'Intercambio' + '.xlsx';
                    ficheros.Proceso := 'ENVIARXLS';
                    repeat
                        ficheros.Secuencia := Secuencia;
                        Secuencia += 1;
                    Until ficheros.Insert();
                    ficheros.CalcFields(Fichero);
                    ficheros.Fichero.CreateOutStream(ExcelStream);
                    TempExcelBuffer.CreateNewBook('Intercambio');
                    TempExcelBuffer.WriteSheet('', '', '', Orientacion::Vertical);
                    TempExcelBuffer.CloseBook();
                    TempExcelBuffer.SetFriendlyFilename('Intercambio');
                    TempExcelBuffer.SaveToStream(ExcelStream, true);
                    //TempExcelBuffer.CreateBook('Informe.xls', 'Informe');
                    ficheros.Modify();
                    ficheros.CalcFields(Fichero);
                    ficheros.Fichero.CreateInStream(Intstream);
                    DownloadFromStream(Intstream, 'Guardar', 'C:\Temp', 'ALL Files (*.*)|*.*', ficheros."Nombre fichero");
                    ficheros.Delete();
                    //TempExcelBuffer.CreateSheet('Saldos','Saldos',COMPANYNAME,USERID);
                    //TempExcelBuffer.GiveUserControl;

                END;
            }
        }
    }
    local procedure EnterCell(
        RowNo: Integer;
        ColumnNo: Integer;
        CellValue: Text[250];
        Bold: Boolean;
        Italic: Boolean;
        UnderLine: Boolean;
        DobleUnderLine: Boolean;
        NumberFormat: Text;
        CellType: Option; Fuente: Text[30]; Tamaño: Integer; Color: Text; ColorFondo: Text)
    begin
        TempExcelBuffer.Init();
        TempExcelBuffer.Validate("Row No.", RowNo);
        TempExcelBuffer.Validate("Column No.", ColumnNo);
        TempExcelBuffer."Cell Value as Text" := CellValue;
        TempExcelBuffer.Formula := '';
        TempExcelBuffer.Bold := Bold;
        TempExcelBuffer.Italic := Italic;
        TempExcelBuffer.Underline := UnderLine;
        TempExcelBuffer."Double Underline" := DobleUnderLine;
        TempExcelBuffer.NumberFormat := NumberFormat;
        TempExcelBuffer."Cell Type" := CellType;
        TempExcelBuffer."Font Name" := Fuente;
        TempExcelBuffer."Font Size" := Tamaño;
        TempExcelBuffer."Font Color" := Color;
        TempExcelBuffer."Background Color" := ColorFondo;
        TempExcelBuffer.Insert();
    end;

    var
        TempExcelBuffer: Record "Excel Buffer 2" temporary;
}