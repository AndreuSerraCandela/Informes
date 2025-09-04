dotnet
{
    assembly("Microsoft.Dynamics.Nav.OpenXml")
    {
        type("Microsoft.Dynamics.Nav.OpenXml.Spreadsheet.WorkbookWriter"; "Microsoft.Dynamics.Nav.OpenXml.Spreadsheet.WorkbookWriter")
        {
            //IsControlAddIn = true;
        }
        //type(IEnumerator;IEnumerator){}

    }

    //IsControlAddIn =
    assembly("DocumentFormat.OpenXml")
    {
        type("DocumentFormat.OpenXml.Spreadsheet.OrientationValues"; OrientationValues)
        {

        }
        type(DocumentFormat.OpenXml.Spreadsheet.Fill; Fill)
        {
            //IsControlAddIn = true;
        }

        type(DocumentFormat.OpenXml.Spreadsheet.PatternFill; PatternFill)
        {
            //IsControlAddIn = true;
        }
        type(DocumentFormat.OpenXml.Spreadsheet.BackgroundColor; BackgroundColor)
        {
            //IsControlAddIn = true;
        }
        type(DocumentFormat.OpenXml.Spreadsheet.ForegroundColor; ForegroundColor)
        {
            //IsControlAddIn = true;
        }
        type(DocumentFormat.OpenXml.Spreadsheet.Font; Font)
        {
            //IsControlAddIn = true;
        }
        type(DocumentFormat.OpenXml.HexBinaryValue; HexBinaryValue)
        {
            //IsControlAddIn = true;
        }
        type(DocumentFormat.OpenXml.Spreadsheet.PatternValues; PatternValues)
        {


            //IsControlAddIn = true;
        }
        type("DocumentFormat.OpenXml.Spreadsheet.Color"; FontColor)
        { }





    }
    assembly(mscorlib)
    {
        type("System.Text.StringBuilder"; TextStringBuilder)
        { }

        type("System.String"; String)
        { }



    }

    assembly(System.Text.Encoding)
    {
        type("System.Text.Encoding"; Encoding)
        { }
    }
    assembly("System.IO")
    {

        type("System.IO.StreamWriter"; StreamWriter)
        { }
        type("System.IO.Stream"; Stream)
        { }

    }
    assembly("System.IO.FileSystem.Primitives")
    {
        type("System.IO.FileMode"; FileMode)
        { }
    }
    assembly("System.Xml")
    {
        type("System.Xml.XmlDocument"; XmlDocument)
        {
            //IsControlAddIn = true;
        }
        type("System.Xml.XmlTextWriter"; XmlTextWriter)
        { }


    }
    assembly("System.Drawing.Common")
    {
        type("System.Drawing.Bitmap"; BitMap) { }
        type("System.Drawing.Graphics"; Graphics) { }
    }






}
