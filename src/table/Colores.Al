table 7001238 Colores
{
    LookupPageId = Colores;
    DrillDownPageId = Colores;
    fields
    {
        field(1; Code; Code[20])
        {
            DataClassification = ToBeClassified;

            trigger OnValidate()
            begin
                if Code = '' then
                    error('Code is required');
            end;
        }
        field(2; "Color Value"; Code[20])
        {
            DataClassification = ToBeClassified;
            trigger OnValidate()
            begin
                if "Color Value" = '' then
                    error('Color Value is required');
                MakePicture();
                "Color Excel" := '00' + copystr("Color Value", 2);
            end;
        }
        field(3; Name; Text[30])
        {
            DataClassification = ToBeClassified;
            trigger OnValidate()
            begin
                if Name = '' then
                    error('Name is required');
            end;
        }
        field(4; "Color Excel"; Text[30])
        {
            DataClassification = ToBeClassified;
        }
        field(15; "Sample Picture"; Media)
        {
            Caption = 'Color';
        }
    }
    keys
    {
        key(PK; "Code")
        {
            Clustered = true;
        }
    }
    var
        BlueCodeTok: Label 'AZUL', Comment = 'Blue';
        BlueGrayCodeTok: Label 'AZUL_GRISAZEO', Comment = 'Blue gray';
        DarkBlueCodeTok: Label 'AZUL_OSCURO', Comment = 'dark blue';
        GreenCodeTok: Label 'VERDE', Comment = 'Green';
        DarkGreenCodeTok: Label 'VERDE_OSCURO', Comment = 'dark green';
        OrangeCodeTok: Label 'NARANJA', Comment = 'Orange';
        DarkOrangeTok: Label 'NARANJA_OSCURO', Comment = 'dark orange';
        RedCodeTok: Label 'ROJO', Comment = 'Red';
        PurpleCodeTok: Label 'MORADO', Comment = 'Purple';
        DarkPurpleCodeTok: Label 'MORADO_OSCURO', Comment = 'dark purple';
        YellowCodeTok: Label 'AMARILLO', Comment = 'yellow';
        TealCodeTok: Label 'VIOLETA', Comment = 'teal';
        BlueTxt: Label 'Azul';
        BlueGrayTxt: Label 'Azul grisaceo';
        DarkBlueTxt: Label 'Azul oscuro';
        GreenTxt: Label 'Verde';
        DarkGreenTxt: Label 'Verde oscuro';
        OrangeTxt: Label 'Naranja';
        DarkOrangeTxt: Label 'Naranja oscuro';
        RedTxt: Label 'Rojo';
        PurpleTxt: Label 'Morado';
        DarkPurpleTxt: Label 'Morado oscuro';
        YellowTxt: Label 'Amarillo';
        TealTxt: Label 'Verde azulado';


    procedure MakePicture()
    var
        TempBlob: Codeunit "Temp Blob";
        Bitmap: DotNet Bitmap;
        Graphics: DotNet Graphics;
        Color: DotNet Color;
        ColorTranslator: DotNet ColorTranslator;
        SolidColorBrush: DotNet SolidBrush;
        ImageFormat: DotNet ImageFormat;
        InStr: InStream;
    begin
        Bitmap := Bitmap.Bitmap(100, 100);
        Graphics := Graphics.FromImage(Bitmap);

        Color := ColorTranslator.FromHtml("Color Value");
        SolidColorBrush := SolidColorBrush.SolidBrush(Color);
        Graphics.FillEllipse(SolidColorBrush, 0, 0, 100, 100);
        Graphics.Dispose();

        TempBlob.CreateInStream(InStr);
        Bitmap.Save(InStr, ImageFormat.Png);

        "Sample Picture".ImportStream(InStr, '');
        Bitmap.Dispose();
    end;

    procedure CreateDefaultBrandColors()
    begin
        CreateBrandColor(RedCodeTok, RedTxt, '#B51725');
        CreateBrandColor(DarkOrangeTok, DarkOrangeTxt, '#DE371C');
        CreateBrandColor(OrangeCodeTok, OrangeTxt, '#FF5709');
        CreateBrandColor(YellowCodeTok, YellowTxt, '#E29D00');
        CreateBrandColor(GreenCodeTok, GreenTxt, '#25892F');
        CreateBrandColor(DarkGreenCodeTok, DarkGreenTxt, '#005C4D');
        CreateBrandColor(TealCodeTok, TealTxt, '#00A199');
        CreateBrandColor(BlueCodeTok, BlueTxt, '#008DD3');
        CreateBrandColor(DarkBlueCodeTok, DarkBlueTxt, '#003A6C');
        CreateBrandColor(BlueGrayCodeTok, BlueGrayTxt, '#536076');
        CreateBrandColor(DarkPurpleCodeTok, DarkPurpleTxt, '#3A327D');
        CreateBrandColor(PurpleCodeTok, PurpleTxt, '#8F65B6');
    end;


    local procedure CreateBrandColor(BrandColorCode: Code[20]; BrandColorName: Text[30]; ColorValue: Code[10])
    var
        Colores: Record "Colores";
    begin
        if Colores.Get(Code) then
            exit;
        Colores.Code := BrandColorCode;
        Colores.Name := BrandColorName;
        Colores."Color Value" := ColorValue;
        Colores."Color Excel" := '00' + copystr(ColorValue, 2);
        Colores.MakePicture();
        Colores.Insert();
    end;
}
//crear pagina para la tabla Colores
page 7001185 "Colores"
{
    PageType = List;
    SourceTable = Colores;
    ApplicationArea = All;
    UsageCategory = Lists;

    Caption = 'Colores';
    layout
    {
        area(content)
        {

            repeater(Group)
            {
                ShowAsTree = true;

                field("Code"; Rec.Code)
                {
                    ApplicationArea = All;

                }
                field(Name; Rec.Name)
                {
                    ApplicationArea = All;
                }

                field("Color Value"; Rec.Code)
                {
                    ApplicationArea = All;
                }
                field("Color Excel"; Rec."Color Excel")
                {
                    ApplicationArea = All;
                }
                field("Sample Picture"; Rec."Sample Picture")
                {
                    ApplicationArea = All;
                }
            }
        }
    }
    actions
    {
        area(processing)
        {
            action("Crear x defecto")
            {
                ApplicationArea = All;
                Caption = 'Crear x defecto';
                Image = Create;
                trigger OnAction()
                begin
                    Rec.CreateDefaultBrandColors();
                end;
            }
        }
        area(Promoted)
        {
            actionref(CrearxDefecto_Promoted; "Crear x defecto")
            {

            }
        }

    }


}