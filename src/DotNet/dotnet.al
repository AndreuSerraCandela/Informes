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
        type("System.IO.FileMode"; FileMoode)
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






}
