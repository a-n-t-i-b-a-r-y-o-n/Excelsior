using System;
using System.Collections.Generic;
using System.Text;

namespace Excelsior
{
    // "I just don't see why we have to put labels on it..."

    // Several types of objects within an xlsx file seem to have an associated .rels file describing
    // things they need e.g. their sharedStrings file, sheets in the workbook, etc.
    // This object represents one of those files for a given parent object.
    class Relationship
    {
        public String ID { get; set; }
        public String Type { get; set; }
        public String Target { get; set; }
        public Relationship() { }
        public Relationship(String pID, String pType, String pTarget)
        {
            ID = pID;
            Type = pType;
            Target = pTarget;
        }
    }
}
