using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml;

namespace Excelsior
{
    public class Workbook
    {
        // Full path to the file
        public String Filename { get; set; }

        // Internal workbook name
        public String Name { get; }

        // Parsed from the .rels file; pairs of internal ids and filenames
        List<Relationship> Relationships { get; set; }

        // Parsed from sharedStrings.xml; numerically-indexed String content of cells
        List<String> SharedStrings { get; set; }

        // Sheet objects contained by the workbook
        public List<Sheet> Sheets { get; set; }


        public Workbook(String filename)
        {
            Filename = filename;
            Sheets = new List<Sheet>();

            TryParse();
        }

        private Boolean TryParse()
        {
            try
            {
                // Unfortunately, everything is read into memory at once to keep from having to actually
                // *extract* things and touching the filesystem
                // You can just read .xlsx files as zip files though. Works well.
                using (FileStream file = new FileStream(Filename, FileMode.Open))
                {
                    // Extract the Excel file contents for reading
                    using (ZipArchive archive = new ZipArchive(file, ZipArchiveMode.Read))
                    {
                        // Attempt to locate the main workbook xml file...
                        if (archive.Entries.Any(entry => entry.FullName == "xl/workbook.xml"))
                        {
                            // Read Workbook's relationships file
                            Relationships = new List<Relationship>();
                            try
                            {
                                // Load file from the _rels directory
                                XmlDocument rels = new XmlDocument();
                                rels.Load(archive.Entries.Where(entry => entry.FullName == "xl/_rels/workbook.xml.rels").First().Open());

                                XmlNamespaceManager relsmgr = new XmlNamespaceManager(rels.NameTable);
                                relsmgr.AddNamespace("r", xmlns.OpenXML.Relationships);

                                // Read all the <Relationship> tags
                                XmlNodeList relationships = rels.SelectNodes("//r:Relationships/r:Relationship", relsmgr);
                                
                                foreach (XmlNode n in relationships)
                                {
                                    Relationships.Add(new Relationship(n.Attributes.GetNamedItem("Id").Value, n.Attributes.GetNamedItem("Type").Value, n.Attributes.GetNamedItem("Target").Value));
                                }
                            }
                            catch (Exception e)
                            {
                                Console.WriteLine("[!] {0}:\n\t{1}\n\n\t{2}", e.GetType().ToString(), e.Message, e.StackTrace);
                            }

                            // Halt parsing with no Relationships file
                            if (Relationships.Count == 0)
                                return false;

                            // Locate and process sharedString file
                            if(Relationships.Any(rel => rel.Type.Contains("sharedStrings")))
                            {
                                SharedStrings = new List<String>();
                                try
                                {
                                    // Load file
                                    XmlDocument sharedStrings = new XmlDocument();
                                    sharedStrings.Load(archive.Entries.Where((entry) => entry.FullName == "xl/sharedStrings.xml").First().Open());

                                    XmlNamespaceManager ssmgr = new XmlNamespaceManager(sharedStrings.NameTable);
                                    ssmgr.AddNamespace("s", xmlns.OpenXML.Main);

                                    // Read all <t> tags
                                    XmlNodeList strings = sharedStrings.SelectNodes("//s:sst/s:si/s:t", ssmgr);
                                    foreach (XmlNode n in strings)
                                    {
                                        SharedStrings.Add(n.InnerText);
                                    }

                                }
                                catch (Exception e)
                                {
                                    Console.WriteLine("[!] {0}:\n\t{1}\n\n\t{2}", e.GetType().ToString(), e.Message, e.StackTrace);
                                }
                            }
                            else
                            {
                                Console.WriteLine("[-] No shared strings file.");
                            }
                            
                            // Read in actual workbook.xml
                            XmlDocument Xml = new XmlDocument();
                            try
                            {
                                Xml.Load(archive.Entries.Where((entry) => entry.FullName == "xl/workbook.xml").First().Open());
                            }
                            catch (Exception e)
                            {
                                Console.WriteLine("[!] {0}:\n\t{1}\n\n\t{2}", e.GetType().ToString(), e.Message, e.StackTrace);
                            }

                            // Namespaces
                            XmlNamespaceManager xmlmgr = new XmlNamespaceManager(Xml.NameTable);

                            // Add a default namespace (needed to find things with XPath for SelectNodes() )
                            xmlmgr.AddNamespace("x", xmlns.OpenXML.Main);

                            // Add relationships namespace
                            xmlmgr.AddNamespace("r", xmlns.OpenXML.Relationships);

                            // Create basic Sheet objects from the <sheet> nodes + other info
                            XmlNodeList sheets = Xml.SelectNodes("//x:workbook/x:sheets/x:sheet", xmlmgr);
                            foreach (XmlNode n in sheets)
                            {
                                // Create sheet with parsed name
                                Sheet s = new Sheet(n.Attributes.GetNamedItem("name").Value);

                                // Try to parse sheet id to integer
                                int id = -1;
                                int.TryParse(n.Attributes.GetNamedItem("sheetId").Value, System.Globalization.NumberStyles.Integer, CultureInfo.CurrentCulture, out id);
                                s.SheetID = id;

                                // Get relationship id for filename resolution
                                s.rID = n.Attributes.GetNamedItem("r:id").Value;

                                // Resolve filename from relationship id
                                if (Relationships.Any(rel => rel.ID == s.rID))
                                    s.Filename = Relationships.Where(rel => rel.ID == s.rID).First().Target;

                                Console.WriteLine("Adding sheet...");
                                Sheets.Add(s);
                            }

                            // Read each sheetN.xml file, resolving sharedStrings and such as we go
                            foreach (Sheet s in Sheets)
                            {
                                XmlDocument sheet = new XmlDocument();
                                try
                                {
                                    sheet.Load(archive.Entries.Where(entry => entry.FullName == "xl/" + s.Filename).First().Open());
                                } catch (Exception e)
                                {
                                    Console.WriteLine("[!] {0}:\n\t{1}\n\n\t{2}", e.GetType().ToString(), e.Message, e.StackTrace);
                                }

                                XmlNamespaceManager smgr = new XmlNamespaceManager(sheet.NameTable);

                                smgr.AddNamespace("s", xmlns.OpenXML.Main);


                                // Read the child nodes of this sheet's <row> tags
                                XmlNodeList rows = sheet.SelectNodes("//s:worksheet/s:sheetData/s:row", smgr);
                                foreach(XmlNode row in rows)
                                {
                                    // Get provided row index, minus one to switch from 1-based to 0-based indices
                                    int r = int.Parse(row.Attributes.GetNamedItem("r").Value)-1;
                                    // Read the <c> cell tags in each row
                                    foreach(XmlNode cell in row.ChildNodes)
                                    {
                                        // Identify column
                                        String colLetter = IndexExtensions.ColumnRegex.Match(cell.Attributes.GetNamedItem("r").Value).Value;
                                        int col = colLetter.ToColumn();

                                        // Add column if it doesn't yet exist
                                        if (!s.Columns.ContainsKey(col))
                                            s.Columns[col] = new Column(colLetter);

                                        // Identify cell's value (dereferencing shared strings if need be)
                                        String value = String.Empty;
                                        try
                                        {
                                            // Handle the value, depending on the type of cell
                                            switch (cell.Attributes.GetNamedItem("t").Value)
                                            {
                                                // Shared string
                                                case "s":
                                                    int idx = int.Parse(cell.ChildNodes[0].InnerText);
                                                    value = SharedStrings[idx];
                                                    break;

                                                // Reference to string in another cell
                                                case "str":
                                                    // TODO: Handle this
                                                    value = "Not Implemented!";
                                                    break;

                                                // Not handled
                                                default:
                                                    // TODO: Handle this
                                                    value = "Not handled!";
                                                    break;
                                            }
                                        } catch (NullReferenceException)
                                        {
                                            // This cell doesn't have a type - it's just a plain value
                                            value = cell.ChildNodes[0].InnerText;
                                        }

                                        s.Columns[col].Cells[r] = value;

                                    }

                                }

                            }

                            return true;

                        }
                        else
                        {
                            Console.WriteLine("[!] workbook.xml not found.");
                            return false;
                        }
                    }
                }
            } catch (IOException e)
            {
                // This is a handy workaround for open locked files in Windozzze
                // 1. Figure out the user's temp directory
                // 2. Drop a shell to copy the file there
                // 3. Recursively call the function with the new filename

                // Double-check we're not _already_ trying to open the copied file
                if (Filename.Contains(Path.GetTempPath().ToString()))
                {
                    Console.WriteLine("[!] Unable to access Temp file!\n\n{0}:\t{1}\n\t{2}\n", e.GetType().ToString(), e.Message, e.StackTrace);
                    return false;
                }

                //                      Temp directory                              Filename without the path
                String NewFilename = Path.GetTempPath().ToString() + Filename.Substring(Filename.LastIndexOf(Path.DirectorySeparatorChar) + 1);

                // Prepare subshell
                Process p = new Process();
                ProcessStartInfo ps = new ProcessStartInfo("CMD.EXE", "/C COPY /Y " + Filename + ' ' + NewFilename);

                // Keep the operation quiet
                ps.CreateNoWindow = true;
                
                // Go!
                p.StartInfo = ps;
                p.Start();
                p.WaitForExit();

                // Update our local filename
                Filename = NewFilename;

                return TryParse();
            } catch (Exception e)
            {
                Console.WriteLine("[!] {0}:\n\t{1}\n\n\t{2}", e.GetType().ToString(), e.Message, e.StackTrace);
                return false;
            }
        }

    }

}
