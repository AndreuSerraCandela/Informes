pageextension 93000 ContratosExt extends "Lista Contratos x Empresa"
{

    local procedure ExportExcel(var Filtros: Record "Filtros Informes"; var ExcelStream: OutStream)
    var
        TempExcelBuffer: Record "Excel Buffer" temporary;
        ContratosLblEP: Label 'Contratos';
        ExcelFileNameEPR: Label 'Contratos_%1_%2';
        RecordLink: Record "Record Link";
        RecordLinkMgt: Codeunit "Record Link Management";
        BgText: Text;
        Row: Integer;
        Col: Integer;
        Informes: Record "Informes";
        DesdeFecha: Date;
        HastaFecha: Date;
        DF: DateFormula;
    begin
        TempExcelBuffer.Reset();
        TempExcelBuffer.DeleteAll();
        procesar(true);
        Row := 1;
        // TempExcelBuffer.AddColumn(Rec.FieldCaption("No."), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        // TempExcelBuffer.AddColumn(Rec.FieldCaption(Description), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        //var ExcelBuffer: Record "Excel Buffer"; RowNo: Integer; ColumnNo: Integer; Value: Variant; IsBold: Boolean; IsItalics: Boolean; IsUnderline: Boolean)
        //Crear estos títulos: Nº Contrato	Cliente	Nombre	Importe	Descripción	Agente	Comentario	Inicio	Fin
        TempExcelBuffer.EnterCell(TempExcelBuffer, Row, 1, 'Nº Contrato', false, false, false);
        TempExcelBuffer.EnterCell(TempExcelBuffer, Row, 2, 'Cliente', false, false, false);
        TempExcelBuffer.EnterCell(TempExcelBuffer, Row, 3, 'Nombre', false, false, false);
        TempExcelBuffer.EnterCell(TempExcelBuffer, Row, 4, 'Importe', false, false, false);
        TempExcelBuffer.EnterCell(TempExcelBuffer, Row, 5, 'Descripción', false, false, false);
        TempExcelBuffer.EnterCell(TempExcelBuffer, Row, 6, 'Agente', false, false, false);
        TempExcelBuffer.EnterCell(TempExcelBuffer, Row, 7, 'Comentario', false, false, false);
        TempExcelBuffer.EnterCell(TempExcelBuffer, Row, 8, 'Inicio', false, false, false);
        TempExcelBuffer.EnterCell(TempExcelBuffer, Row, 9, 'Fin', false, false, false);
        //Rec.SetRange("Nº Proveedor", Rec."Nº Proveedor");
        If Filtros.Desde <> DF then DesdeFecha := CalcDate(Filtros.Desde, WorkDate()) else DesdeFecha := 0D;
        If Filtros.Hasta <> DF then HastaFecha := CalcDate(Filtros.Hasta, WorkDate()) else HastaFecha := Calcdate('99A', HastaFecha);
        Case
            Filtros.Campo Of
            Rec.FieldNo("Sell-to Customer No."):
                Rec.SetRange("Sell-to Customer No.", Filtros.Valor);
            Rec.FieldNo("No."):
                Rec.SetFilter("No.", '%1', Filtros.Valor);
            Rec.FieldNo("Salesperson Code"):
                Rec.SetFILTER("Salesperson Code", '%1', Filtros.Valor);
        End;
        if Rec.FindSet() then
            repeat

                TempExcelBuffer.EnterCell(TempExcelBuffer, Row, 2, Rec."No.", false, false, false);
                TempExcelBuffer.EnterCell(TempExcelBuffer, Row, 3, Rec."Sell-to Customer No.", false, false, false);
                TempExcelBuffer.EnterCell(TempExcelBuffer, Row, 4, Rec."Sell-to Customer Name", false, false, false);
                TempExcelBuffer.EnterCell(TempExcelBuffer, Row, 5, Rec.Amount, false, false, false);
                TempExcelBuffer.EnterCell(TempExcelBuffer, Row, 6, Rec."Fecha inicial proyecto", false, false, false);
                TempExcelBuffer.EnterCell(TempExcelBuffer, Row, 7, Rec."Fecha fin proyecto", false, false, false);


            until Rec.Next() = 0;
        TempExcelBuffer.CreateNewBook(ContratosLblEP);
        TempExcelBuffer.WriteSheet(ContratosLblEP, CompanyName, UserId);
        TempExcelBuffer.CloseBook();
        TempExcelBuffer.SetFriendlyFilename(StrSubstNo(ExcelFileNameEPR, CurrentDateTime, UserId));
        TempExcelBuffer.SaveToStream(ExcelStream, true);
    end;


}