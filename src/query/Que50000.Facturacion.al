query 50000 Facturacion
{
    Caption = 'Facturacion';
    QueryType = Normal;

    elements
    {
        dataitem(Sales_Invoice_Header; "Sales Invoice Header")
        {
            column(BilltoCustomerNo; "Bill-to Customer No.")
            {
            }
            column(BilltoName; "Bill-to Name")
            {
            }
            column(PostingDate; "Posting Date")
            {
            }
            column(NContrato; "Nº Contrato")
            {
            }
            column(NProyecto; "Nº Proyecto")
            {
            }
            column(Amount; Amount)
            {
            }
            column(AmountIncludingVAT; "Amount Including VAT")
            {
            }
        }
    }

    trigger OnBeforeOpen()
    begin

    end;
}
