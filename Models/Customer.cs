using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EasyMig.Models
{
    public class Customer
    {
        public Guid Id { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string City { get; set; }
        public string Phone { get; set; }
        public string ZipCode { get; set; }
    }

    public class MaterialPlant
    {
        public int Material_No { get; set; }
        public int Plant { get; set; }
    }
    public class MaterialSales
    {
        public int Material_No { get; set; }
        public int Sales_Org { get; set; }
        public int Dist_Channel { get; set; }
    }

    public class MaterialClassification
    {
        public int Legacy_Mat_No { get; set; }
        public string Class { get; set; }
        public string Class_Character { get; set; }
        public string Class_Value { get; set; }

    }
}





