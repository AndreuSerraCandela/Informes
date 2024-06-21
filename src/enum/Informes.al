enum 7001190 Informes
{
    value(0; "Contratos x Empresa") { }
    value(1; "Estadisticas Contabilidad") { }
    value(2; "Tablas") { }
    value(3; "Web Service") { }
    value(4; "Informes Financieros") { }
    value(5; "Saldo InterEmpresas") { }
}
enum 7001191 Funciones
{
    //'Importe,Vendedor,GetTotImp,ImporteIva,GetImpBorFac,GetImpBorAbo,GetImpFac,GetImpAbo,GetTotCont'
    value(0; " ") { Caption = '------'; }
    value(1; "Importe") { Caption = 'Importe'; }
    value(2; "Vendedor") { Caption = 'Vendedor'; }
    value(3; "GetTotImp") { Caption = 'Facturas-Abonos'; }
    value(4; "ImporteIva") { Caption = 'ImporteIva'; }
    value(5; "GetImpBorFac") { Caption = 'Importe Facturas Borrador'; }
    value(6; "GetImpBorAbo") { Caption = 'Importe Abonos Borrador'; }
    value(7; "GetImpFac") { Caption = 'Importe Facturas'; }
    value(8; "GetImpAbo") { Caption = 'Importe Abonos'; }
    value(9; "GetTotCont") { Caption = 'Total Contrato'; }
    value(10; "Cliente_Proveedor") { Caption = 'Nombre Tercero'; }
    value(11; "Año") { Caption = 'Año'; }
    value(12; "Mes") { Caption = 'Mes'; }
    value(13; "Semana") { Caption = 'Semana'; }
    value(14; "Diferencia") { Caption = 'Diferencia'; }
    value(15; "GetTotContNew") { Caption = 'Total Contrato o Total Facturado'; }
    value(16; "Columna") { Caption = 'Columna'; }

}
enumextension 90144 EnumEscenarioInformes extends "Email Scenario"
{
    value(90003; Informes)
    {
        Caption = 'Informes';
    }


}
