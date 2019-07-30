using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using EasyMig.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Drawing;
using System.Drawing;
using System.IO;
using System.Configuration;
using System.Data.SqlClient;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace EasyMig.Utility
{
    public class TestData
    {
        //customers

        public void GenerateTestData(Int32 noOfRecords)
        {
            EasyMig.Utility.SampleCustomerRepository repository = new SampleCustomerRepository();
            // to get sample data
            var customers = repository.GetCustomers(noOfRecords);

            //Create datatable
            DataTable dt = CreateDataTable(customers);

            //export data to excel file 
            string fileName = ExportToExcel(dt);

            SaveFile(fileName);
        }


        //material Plant

        public List<Models.MaterialPlant> GenerateMaterialPlantTestData(int Plantrecords)
        {
            EasyMig.Utility.SampleCustomerRepository repository = new SampleCustomerRepository();
            // to get sample data
            var materialplant = repository.GetMaterialPlant(Plantrecords);
            // var m = GeneratePlantData(materialplant);

            //Create datatable
            DataTable dt4 = CreateDataTableMaterialPlant(materialplant);

            //export data to excel file 
            string fileName = ExportToExcel(dt4);

            SaveFile(fileName);
            return  materialplant.ToList();
        }

        public void GenerateMaterialSalesTestData(int Plantrecords)
        {
            EasyMig.Utility.SampleCustomerRepository repository = new SampleCustomerRepository();
            // to get sample data
            var materialsales = repository.GetMaterialSales(Plantrecords);
            // var m = GeneratePlantData(materialplant);

            //Create datatable
            DataTable dt7 = CreateDataTableMaterialSales(materialsales);

            //export data to excel file 
            string fileName = ExportToExcel(dt7);

            SaveFile(fileName);
        }

        /// <summary>
        /// This method used to get the test data which already created
        /// </summary>
        /// <returns></returns>
        public List<EasyMig.Models.TestDataFile> GetTestData()
        {
            List<EasyMig.Models.TestDataFile> testDataFiles = GetSampleDataFiles();
            return testDataFiles;
        }

        /// <summary>
        /// This method is to get test data files from database
        /// </summary>
        /// <returns></returns>
        public List<Models.TestDataFile> GetSampleDataFiles()
        {
            string strCon = ConfigurationManager.ConnectionStrings["easyMigCon"].ConnectionString.ToString();
            string strPath = ConfigurationManager.AppSettings["filePath"].ToString();
            SqlConnection con = new SqlConnection(strCon);
            SqlDataAdapter da = new SqlDataAdapter();
            DataSet ds = new DataSet();
            List<Models.TestDataFile> objFile = new List<TestDataFile>();
            try
            {
                SqlCommand cmd = new SqlCommand();
                cmd.CommandText = "SELECT * FROM TestDataFile";
                cmd.Connection = con;
                da.SelectCommand = cmd;
                da.Fill(ds);

                if (ds.Tables[0].Rows.Count > 0)
                {
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        TestDataFile obj = new TestDataFile();
                        obj.Id = Convert.ToInt32(ds.Tables[0].Rows[i]["Id"]);
                        obj.Name = ds.Tables[0].Rows[i]["Name"].ToString();
                        obj.CreatedOn = Convert.ToDateTime(ds.Tables[0].Rows[i]["CreatedOn"]);
                        obj.DownloadPath = strPath + ds.Tables[0].Rows[i]["Name"].ToString();
                        objFile.Add(obj);
                    }
                }
                return objFile;

            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (con.State == ConnectionState.Open)
                    con.Close();
            }
        }

        public void SaveFile(string fileName)
        {
            string strCon = ConfigurationManager.ConnectionStrings["easyMigCon"].ConnectionString.ToString();
            SqlConnection con = new SqlConnection(strCon);
            try
            {
                SqlCommand cmd = new SqlCommand();
                cmd.CommandText = "INSERT INTO TestDataFile (Name,CreatedOn) VALUES ('" + fileName + "','" + DateTime.Now + "')";
                cmd.Connection = con;
                con.Open();
                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (con.State == ConnectionState.Open)
                    con.Close();
            }
        }

        public DataTable CreateDataTable(IEnumerable<Customer> customers)
        {
            DataTable Dt = new DataTable("SampleData");

            Dt.Columns.Add(new DataColumn("Id", typeof(string))); 
            Dt.Columns.Add(new DataColumn("FirstName", typeof(string)));
            Dt.Columns.Add(new DataColumn("LastName", typeof(string)));
            Dt.Columns.Add(new DataColumn("City", typeof(string)));
            Dt.Columns.Add(new DataColumn("Phone", typeof(string)));
            Dt.Columns.Add(new DataColumn("ZipCode", typeof(string)));
            DataRow dr;

            Customer c1;

            foreach (object obj in customers)
            {
                c1 = (Customer)obj;
                dr = Dt.NewRow();
                //Add rows  
                dr["Id"] = c1.Id;
                dr["FirstName"] = c1.FirstName;
                dr["LastName"] = c1.LastName;
                dr["City"] = c1.City;
                dr["Phone"] = c1.Phone;
                dr["ZipCode"] = c1.ZipCode;



                Dt.Rows.Add(dr);
            }
            return Dt;
        }



        //material Plant list
        public List<MaterialPlant> GeneratePlantData(IEnumerable<MaterialPlant> materialplant)
        {
            List<MaterialPlant> materialPlant = new List<MaterialPlant>();
            DataSet dt4 = new DataSet();

            //List<int> materialNos = new List<int> { 1100000040, 1100000041, 1100000042, 1100000043, 1100000044 };
            // List<int> plants = new List<int> { 1000, 1100, 1200, 1300, 1400, 1500, 1600 };


            //foreach (var matNo in materialNos) // This is to get each Material No.
            //{
            //    foreach (var plant in plants) // This is to get each Language No
            //    {
            //        materialPlant.Add(new MaterialPlant { Material_No = matNo, Plant = plant });

            //    }

            //}
            return materialPlant;
        }


        
        public DataTable CreateDataTableMaterialPlant(IEnumerable<MaterialPlant> materialPlants)
        {
            //material Plant
            List<int> plants = new List<int> { 1000, 1100, 1200, 1300, 1400, 1500, 1600 };

            Int64 matNo = 1100000034;
            DataTable Dt4 = new DataTable("SampleData4");
            DataRow dr4;
            Dt4.Columns.Add(new DataColumn("Material_No", typeof(Int64)));
            Dt4.Columns.Add(new DataColumn("Plant", typeof(Int32)));

            DataTable dt4 = new DataTable("MaterialPlant");
            dt4.Columns.Add("MaterialNo", typeof(Int64));
            dt4.Columns.Add("Plant", typeof(Int64));

            for (int i = 0; i <= 20; i++)
            {
                int rnp = RandomNumber1(1, 5);
                i = i + rnp;
                Int64 randomPlant = 0;
                for (int j = 0;j< rnp;j++)
                {
                    dr4 = dt4.NewRow();
                    dr4["MaterialNo"] = matNo;
                    do
                    {
                        randomPlant = plants[RandomNumber1(1, 7)];
                    } while (dt4.AsEnumerable().Where(x => x.Field<Int64>("MaterialNo") == matNo).Select(x => x.Field<Int64>("Plant")).Contains(randomPlant));
                    dr4["Plant"] = randomPlant;

                    dt4.Rows.Add(dr4);
                }
                matNo = matNo + 1;
            }
            return dt4;
        }


        private static readonly Random random = new Random();
        private static readonly object syncLock = new object();
        public static int RandomNumber1(int min, int max)
        {
            lock (syncLock)
            { // synchronize
                return random.Next(min, max);
            }
        }
        public DataTable CreateDataTableMaterialSales(IEnumerable<MaterialSales> materialSales)
        {
            //material Sales
            List<int> sales_Org = new List<int> { 1000, 1020, 2000, 3000, 3020, 7500 };
            List<int> chanals = new List<int> { 10,20,12, 30 };

            DataTable Dt4 = new DataTable("SampleData4");
            DataRow dr4;
            Int64 matNo = 1100000034;

            Dt4.Columns.Add(new DataColumn("Material_No", typeof(Int64)));
            Dt4.Columns.Add(new DataColumn("Sales_Org", typeof(Int32)));
            Dt4.Columns.Add(new DataColumn("Dist_Channel", typeof(Int32)));
            

            DataTable dt4 = new DataTable("MaterialPlant");
            dt4.Columns.Add("MaterialNo", typeof(Int64));
            dt4.Columns.Add("Sales_Org", typeof(Int64));
            dt4.Columns.Add("Dist_Channel", typeof(Int64));

            for (int i = 0; i <= 20; i++)
            {
                int rnp = RandomNumber1(1, 4);
                i = i + rnp;
                Int64 randomChannel = 0, randomSales = 0;
                for (int j = 0; j < rnp; j++)
                {
                    dr4 = dt4.NewRow();
                    dr4["MaterialNo"] = matNo;
                    do
                    {
                        randomSales = sales_Org[RandomNumber1(1, 6)];
                    } while (dt4.AsEnumerable().Where(x => x.Field<Int64>("MaterialNo") == matNo).Select(x => x.Field<Int64>("Sales_Org")).Contains(randomSales));
                    
                    do
                    {
                        randomChannel = chanals[RandomNumber1(1,4)];
                    } while (dt4.AsEnumerable().Where(x => x.Field<Int64>("MaterialNo") == matNo).Select(x => x.Field<Int64>("Dist_Channel")).Contains(randomChannel));

                    dr4["Sales_Org"] = randomSales;
                    dr4["Dist_Channel"] = randomChannel;
                    dt4.Rows.Add(dr4);
                }
                matNo = matNo + 1;
            }
            return dt4;
        }

        ////material sales list
        //public List<MaterialSales> GenerateSalesData(IEnumerable<MaterialSales> materialsales)
        //{
        //    List<MaterialSales> materialSales = new List<MaterialSales>();
        //    DataSet dt7 = new DataSet();

        //    return materialSales;
        //}


        ////material sales
        //List<int> sales_org = new List<int> { 1000, 1020, 2000, 3000, 3020, 7500};
        //List<int> dist_channel = new List<int> { 10, 20, 12, 10, 30, 10};
        //public DataTable CreateDataTableMaterialSales(IEnumerable<MaterialSales> materialsales)
        //{

        //    DataRow dr7; 
        //    //DataTable distinctTable = dt.DefaultView.ToTable( /*distinct*/ true);
        //    int rn = 0;

        //    DataTable dt7 = new DataTable("MaterialSales");
        //    dt7.Columns.Add("MaterialNo", typeof(Int64));
        //    dt7.Columns.Add("Sales_Org", typeof(Int64));
        //    dt7.Columns.Add("Dist_Channel", typeof(Int64));

        //    Int64 matNo = 1100000037;
        //    //  Int64 plantNo = 1000;
        //    Int32 sameNoCount = 0;
        //    rn = RandomNumber(1, 6);


        //    for (int i = 1; i <= 20; i++)
        //    {
        //        int rns = RandomNumber(0, 5);
        //        dr7 = dt7.NewRow();
        //        dr7["MaterialNo"] = matNo;
        //        dr7["Sales_Org"] = sales_org[rns];
        //        dr7["Dist_Channel"] = dist_channel[rns];
        //        dt7.Rows.Add(dr7);

        //        if (sameNoCount == rn)
        //        {
        //            matNo++;
        //            sameNoCount = 0;
        //            rn = RandomNumber(1, 6);
        //        }
        //        sameNoCount++;

        //    }

        //    return dt7;
        //}


        public void GenerateMaterialClassificationTestData(int recorder)
        {
            EasyMig.Utility.SampleCustomerRepository repository = new SampleCustomerRepository();
            // to get sample data
            var materialclassification = repository.GetMaterialClassification(recorder);
            var m = GenerateClassificationData(materialclassification);

            //Create datatable
            DataTable dt7 = CreateDataTableMaterialClassification(m);
            //    var materialPlant = GeneratePlantData(records);

            //export data to excel file 
            string fileName = ExportToExcel(dt7);

            SaveFile(fileName);
        }


        //material Classification list
        public List<MaterialClassification> GenerateClassificationData(IEnumerable<MaterialClassification> materialclassification)
        {
            
            List<string> classer = new List<string> { "DPC1_FDD", "MOTOR", "IP_TEL_HW", "GG_SUPPLY_MEDIA", "GG_SUPPLY_MEDIA" }; //These are material Types
            List<string> classchar = new List<string> { "DPCX_FDD", "HD_MOTOR", "IP_TEL_TYPE", "GG_DB_SYSTEM", "GG_OS_SYSTEM" }; //These are material Types
            List<string> classval = new List<string> { "001", "1300", "SYSTEM_1000", "OR", "UX" }; //These are material Types
                                                                                                   // List<int> materialNo = new List<int> { 1100000034, 1100000035, 1100000036, 1100000037, 1100000038, 1100000039 };
                                                                                                   // List<int> Old_Material = new List<int> { 10034,10034,10035,10036,10037,10038,10039};

            var matClassification = materialclassification.ToList();
            List<MaterialClassification> mm = new List<MaterialClassification>();
            int Legacy_Mat_No = 10051;
            foreach (var v in matClassification) // for count Ex- 10
            {
                
                for (int i = 0; i <= 3; i++)
                {
                    matClassification[i].Legacy_Mat_No = Legacy_Mat_No;
                    matClassification[i].Class = classer[i];
                    matClassification[i].Class_Character = classchar[i];
                    matClassification[i].Class_Value = classval[i];

                    mm.Add(new MaterialClassification() { Class = matClassification[i].Class, Class_Character = matClassification[i].Class_Character, Class_Value = matClassification[i].Class_Value });
                }
                Legacy_Mat_No = Legacy_Mat_No + 1;
            }

            return mm;
        }



        public DataTable CreateDataTableMaterialClassification(IEnumerable<MaterialClassification> materialclassification)
        {
            DataTable Dt6 = new DataTable("Sampledata");
            Dt6.Columns.Add(new DataColumn("Class", typeof(string)));

            Dt6.Columns.Add(new DataColumn("Legacy_Mat_No", typeof(Int64)));
            Dt6.Columns.Add(new DataColumn("Class_Character", typeof(string)));
            Dt6.Columns.Add(new DataColumn("Class_Value", typeof(string)));
            DataRow dr6;

            MaterialClassification c6;
            //Int64 legmatno = 10014;
            //Boolean isData = false;
            //Int32 q = 1;

            foreach (object obj in materialclassification)
            {
                c6 = (MaterialClassification)obj;
                dr6 = Dt6.NewRow();
                //Add rows  
                dr6["Legacy_Mat_No"] =  //matNo;
                dr6["Class"] = c6.Class;
                dr6["Class_Value"] = c6.Class_Value;
                dr6["Class_Character"] = c6.Class_Character;

                //isData = isData ? false : true;
                //legmatno++;
                //q++;

                Dt6.Rows.Add(dr6);
            }
            return Dt6;
        }



        public void ReadStructure()
        {
            //TODO:
        }

        public string ExportToExcel(DataTable dt)
        {
            try
            {
                //string strPath = @"D:\Projects\EasyMIG\EasyMig\EasyMig\ExcelExports\";
                string strPath = ConfigurationManager.AppSettings["filePath"].ToString();
                string fName = string.Empty;

                using (ExcelPackage objPackage = new ExcelPackage())
                {
                    ExcelWorksheet objSheet = objPackage.Workbook.Worksheets.Add(dt.TableName);
                    objSheet.Cells["A1"].LoadFromDataTable(dt, true);
                    objSheet.Cells.Style.Font.SetFromFont(new Font("Calibri", 10));
                    objSheet.Cells.AutoFitColumns();
                    //Format the header    
                    using (ExcelRange objRange = objSheet.Cells["A1:XFD1"])
                    {
                        objRange.Style.Font.Bold = true;
                        objRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        objRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    }
                    Random random = new Random();
                    //fName = "SampleData7558.xlsx";
                    fName = "SampleData" + DateTime.Now.ToShortDateString() + random.Next().ToString() + ".xlsx";
                    strPath = strPath + fName;

                    //Write it back to the client    
                    if (File.Exists(strPath))
                        File.Delete(strPath);

                    //Create excel file on physical disk    
                    FileStream objFileStrm = File.Create(strPath);
                    objFileStrm.Close();

                    //Write content to excel file    
                    File.WriteAllBytes(strPath, objPackage.GetAsByteArray());
                }
                return fName;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        public List<Mandate> ValidateExcelFile(string file)
        {
            try
            {
                string strPath = ConfigurationManager.AppSettings["filePath"].ToString();
                DataTable dt = new DataTable();
                bool hasHeader = true;

                using (ExcelPackage exlPackage = new ExcelPackage())
                {
                    using (var fstream = File.OpenRead(file))
                    {
                        exlPackage.Load(fstream);
                    }

                    var ws = exlPackage.Workbook.Worksheets[1];
                    foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                    {
                        dt.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
                    }

                    var startRow = hasHeader ? 2 : 1;
                    for (int rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
                    {
                        var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                        DataRow row = dt.Rows.Add();
                        foreach (var cell in wsRow)
                        {
                            row[cell.Start.Column - 1] = cell.Text;
                        }
                    }
                }

                var mandate = MandatoryCheck(dt);

                return mandate;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        public List<Mandate> MandatoryCheck(DataTable dt)
        {
            try
            {
                string[] mandateColumns = ConfigurationManager.AppSettings["mandateColumns"].ToString().Split(',');
                List<Mandate> lstMandate = new List<Mandate>();
                Mandate mandate = null;
                //DataRow dRow = null;
                DataTable dtResult = dt.Clone();
                int colIndex = 0;

                for (int i = 0; i < mandateColumns.Length; i++)
                {
                    //dRow = dt.Rows[i];
                    colIndex = dt.Columns[mandateColumns[i]].Ordinal;
                    if (dt.Columns["First Name"].ToString() == mandateColumns[i].ToString())
                    {
                        for (int j = 0; j < dt.Rows.Count; j++)
                        {
                            if (string.IsNullOrWhiteSpace(dt.Rows[j][colIndex].ToString()))
                            {
                                mandate = new Mandate();
                                mandate.rowId = Convert.ToInt32(dt.Rows[j][0]);
                                mandate.FirstName = dt.Rows[j][colIndex].ToString();
                                mandate.LastName = dt.Rows[j]["Last Name"].ToString();
                                lstMandate.Add(mandate);
                            }
                        }
                    }
                }
                return lstMandate;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}

