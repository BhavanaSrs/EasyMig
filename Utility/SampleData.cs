using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using EasyMig.Models;
using StackExchange.Redis;
using Bogus;

namespace EasyMig.Utility
{
    public class SampleCustomerRepository
    {
        public IEnumerable<Customer> GetCustomers(Int32 noOfRecords)
        {
            Randomizer.Seed = new Random(123456);
            var customerGenerator = new Faker<Customer>()
                .RuleFor(c => c.Id, Guid.NewGuid())
                .RuleFor(c => c.FirstName, f => f.Person.FirstName)
                .RuleFor(c => c.LastName, f => f.Person.LastName)
                .RuleFor(c => c.City, f => f.Person.Address.City)
                .RuleFor(c => c.ZipCode, f => f.Person.Address.ZipCode)
                .RuleFor(c => c.Phone, f => f.Phone.PhoneNumber());

            return customerGenerator.Generate(noOfRecords);

        }


        public IEnumerable<Models.MaterialPlant> GetMaterialPlant(Int32 noOfRecords)
        {
            //Randomizer.Seed = new Random(12);
            Random random = new Random();
            //var materialplantGenerator = new Faker<Models.MaterialPlant>()
            //   .RuleFor(c => c.Material_No, random.Next(1,10000))
            //   .RuleFor(c => c.Plant, random.Next(1, 100000));
            List<Models.MaterialPlant> model = new List<MaterialPlant>();
            for (int i = 0; i < noOfRecords; i++)
            {
                var m = new Models.MaterialPlant
                {
                    Material_No = random.Next(1, 20),
                    Plant = random.Next(1, 10000)
                };
                model.Add(m);
            }
            return model;
            //return materialplantGenerator.Generate(noOfRecords);
        }
        public IEnumerable<Models.MaterialSales> GetMaterialSales(Int32 noOfRecords)
        {

            Randomizer.Seed = new Random(12);
            var materialsalesGenerator = new Faker<Models.MaterialSales>()

               .RuleFor(c => c.Material_No, 1)
               .RuleFor(c => c.Sales_Org, 1)
               .RuleFor(c => c.Dist_Channel, 1);

            return materialsalesGenerator.Generate(noOfRecords);
        }

        public IEnumerable<MaterialClassification> GetMaterialClassification(int noOfRecords)
        {
            Randomizer.Seed = new Random(12);
            var materialClassificationGenerator = new Faker<MaterialClassification>()
               .RuleFor(c => c.Legacy_Mat_No, 1)
               .RuleFor(c => c.Class, "1")
               .RuleFor(c => c.Class_Character, "1")
               .RuleFor(c => c.Class_Value, "1");

            return materialClassificationGenerator.Generate(noOfRecords);
        }
    }
}
