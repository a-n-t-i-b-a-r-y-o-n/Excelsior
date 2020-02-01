using System;
using System.Collections.Generic;
using System.Text;

namespace Excelsior
{
    // Helper class XML for namespaces
    public static class xmlns
    {
        // Known OpenXML SpreadsheetML schemas
        public static class OpenXML
        {
            public static String Main = @"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            public static String Relationships = @"http://schemas.openxmlformats.org/package/2006/relationships";

            public static class OfficeDocument
            {
                public static class Relationships
                {
                    public static String Worksheet = @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
                    public static String SharedStrings = @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
                    public static String Theme = @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
                    public static String Styles = @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
                }
            }
        }

        public static class MS
        {
            // Known Microsoft SpreadsheetML schemas, ordered by year
            public static class x2010
            {
                public static String Main = @"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main";
            }

            public static class x2018
            {
                public static String CalcFeatures = @"http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures";
            }
        }

    }
}
