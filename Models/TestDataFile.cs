using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EasyMig.Models
{
    public class TestDataFile
    {
        public Int32 Id { get; set; }
        public string Name { get; set; }
        public DateTime CreatedOn { get; set; }
        public string DownloadPath { get; set; }
    }

    public class Mandate
    {
        public Int32 rowId { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string City { get; set; }
    }

    public class Mandate2
    {
        public int Material_No { get; set; }
        public int Plant { get; set; }

    }
    public class Mandate3
    {
        public int Material_No { get; set; }
        public int Sales_Org { get; set; }
       public int Dist_Channel { get; set; }
    }
    public class Mandate4
    {
        public int Legacy_Mat_No { get; set; }
        public string Class { get; set; }
        public string Class_Character { get; set; }
        public string Class_Value { get; set; }

    }

}