using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;
using System.Text.RegularExpressions;

namespace cap_revision_sec321
{
    public class Shipment
    {
        public int Id { get; set; }
        public string Shipment_Control_Number { get; set; }
        public string Shipment_Type { get; set; }
        public string Shipper_Name { get; set; }
        public string Shipper_Address { get; set; }
        public string Shipper_City { get; set; }
        public string Shipper_Country { get; set; }
        public string Shipper_State { get; set; }
        public string Shipper_Postal { get; set; }
        public string Shipper_Port_of_Lading { get; set; }
        public string Consignee_Name { get; set; }
        public string Consignee_Address { get; set; }
        public string Consignee_City { get; set; }
        public string Consignee_Country { get; set; }
        public string Consignee_State { get; set; }
        public string Consignee_Postal { get; set; }
        public string Product_Description { get; set; }
        public string Product_Qty { get; set; }
        public string Product_UOM { get; set; }
        public string Product_Weight { get; set; }
        public string Product_Unit_of_Weight { get; set; }
        public string Product_Value { get; set; }
        public string Customer_Reference { get; set; }
        public string US_Port_Arrive { get; set; }
        public string Fn_Port_Loading { get; set; }
        public string Fn_Port_Reciept { get; set; }
        public string Origin { get; set; }
        public static List<Shipment> Import_To_Grid(string FilePath, string Sheet)
        {
            var lstShipments = new List<Shipment>();
            string constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FilePath + ";Extended Properties=\"Excel 12.0 Xml;HDR=NO;\"";
            using (OleDbConnection conn = new OleDbConnection(constr))
            {
                conn.Open();
                OleDbCommand command = new OleDbCommand("Select * from [" + Sheet + "$]", conn);
                OleDbDataReader reader = command.ExecuteReader();



                if (reader.HasRows)
                {
                    int Count = 0;
                    while (reader.Read())
                    {
                        var oShipment = new Shipment();

                        if (Count == 0)
                        {

                        }
                        else
                        {
                            oShipment.Id = int.Parse(reader[0].ToString());
                            oShipment.Shipment_Control_Number = reader[1].ToString();
                            oShipment.Shipment_Type = reader[2].ToString();
                            oShipment.Shipper_Name = reader[3].ToString();
                            oShipment.Shipper_Address = reader[4].ToString();
                            oShipment.Shipper_City = reader[5].ToString();
                            oShipment.Shipper_Country = reader[6].ToString();
                            oShipment.Shipper_State = reader[7].ToString();
                            oShipment.Shipper_Postal = reader[8].ToString();
                            oShipment.Shipper_Port_of_Lading = reader[9].ToString();
                            oShipment.Consignee_Name = reader[10].ToString();
                            oShipment.Consignee_Address = reader[11].ToString();
                            oShipment.Consignee_City = reader[12].ToString();
                            oShipment.Consignee_Country = reader[13].ToString();
                            oShipment.Consignee_State = reader[14].ToString();
                            oShipment.Consignee_Postal = reader[15].ToString();
                            oShipment.Product_Description = reader[16].ToString();
                            oShipment.Product_Qty = reader[17].ToString();
                            oShipment.Product_UOM = reader[18].ToString();
                            oShipment.Product_Weight = reader[19].ToString();
                            oShipment.Product_Unit_of_Weight = reader[20].ToString();
                            oShipment.Product_Value = reader[21].ToString();
                            oShipment.Customer_Reference = reader[22].ToString();
                            oShipment.US_Port_Arrive = reader[23].ToString();
                            oShipment.Fn_Port_Loading = reader[24].ToString();
                            oShipment.Fn_Port_Reciept = reader[25].ToString();
                            oShipment.Origin = reader[26].ToString();
                            lstShipments.Add(oShipment);

                        }

                        Count++;

                    }
                }
                return lstShipments;
            }
        }
    }
    public class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            //C:\Users\PJUAREZ\source\repos\cap_revision_sec321\cap_revision_sec321\bin\Debug\netcoreapp3.1\PLUMA_NACIONAL.xlsx
            string FilePath = @"C:\Users\PJUAREZ\source\repos\cap_revision_sec321\cap_revision_sec321\bin\Debug\netcoreapp3.1\files\PLUMA_NACIONAL.xlsx";
            string Sheet = "Shipments";

            List<Shipment> lstShipments = Shipment.Import_To_Grid(FilePath, Sheet);

            lstShipments.ForEach(oShipment => {


                bool consignee_name_check = ConsigneeNameCheck(oShipment.Consignee_Name);
                
                if (!consignee_name_check) 
                {
                    Log("ERROR IN LINE " + oShipment.Id + " DETECTED. VALUE: " + oShipment.Consignee_Name + " - ConsigneeName");
                }

                bool shipper_address_check = ShipperAddressCheck(oShipment.Shipper_Address);

                if (!shipper_address_check)
                {
                    Log("ERROR IN LINE " + oShipment.Id + " DETECTED. VALUE: " + oShipment.Shipper_Address + " - ShipperAddress");
                }

                bool shipper_postal_check = ShipperPostalCheck(oShipment.Shipper_Postal);

                if (!shipper_postal_check)
                {
                    Log("ERROR IN LINE " + oShipment.Id + " DETECTED. VALUE: " + oShipment.Shipper_Postal + " - ShipperPostal");
                }

                bool product_value_check = ProductValueCheck(oShipment.Product_Value);

                if (!shipper_postal_check)
                {
                    Log("ERROR IN LINE " + oShipment.Id + " DETECTED. VALUE: " + oShipment.Product_Value + " - ProductValue");
                }

                bool consignee_state_check = ConsigneeStateCheck(oShipment.Consignee_State);

                if (!consignee_state_check)
                {
                    Log("ERROR IN LINE " + oShipment.Id + " DETECTED. VALUE: " + oShipment.Consignee_State + " - ConsigneeState");
                }

                bool shipper_state_check = ShipperStateCheck(oShipment.Shipper_State);

                if (!shipper_state_check)
                {
                    Log("ERROR IN LINE " + oShipment.Id + " DETECTED. VALUE: " + oShipment.Shipper_State + " - ShipperState");
                }

                bool shipper_country_check = ShipperCountryCheck(oShipment.Shipper_Country);

                if (!shipper_country_check)
                {
                    Log("ERROR IN LINE " + oShipment.Id + " DETECTED. VALUE: " + oShipment.Shipper_Country + " - ShipperCountry");
                }

                bool shipper_name_check = ShipperNameCheck(oShipment.Shipper_Name);

                if (!shipper_name_check)
                {
                    Log("ERROR IN LINE " + oShipment.Id + " DETECTED. VALUE: " + oShipment.Shipper_Name + " - ShipperName");
                }

                bool consignee_address_check = ConsigneeAddressCheck(oShipment.Consignee_Address);

                if (!consignee_address_check)
                {
                    Log("ERROR IN LINE " + oShipment.Id + " DETECTED. VALUE: " + oShipment.Consignee_Address + " - ConsigneeAddress");
                }
            });

            Log(lstShipments.Count + " LINES PROCESSED.");
        }
        
        public static bool ConsigneeNameCheck(string password)
        {
            bool result = false;

            Regex rgx = new Regex(@"^[a-zA-Z0-9_ ]*$");

            if (rgx.IsMatch(password))
            {
                result = true;
            }

            return result;
        }
        public static bool ShipperAddressCheck(string password)
        {
            bool result = false;

            Regex rgx = new Regex(@"^[a-zA-Z0-9_ ]*$");

            if (rgx.IsMatch(password))
            {
                result = true;
            }

            return result;
        }

        public static bool ShipperPostalCheck(string password)
        {
            bool result = false;

            Regex rgx = new Regex(@"^[0-9]{5}$");

            if (rgx.IsMatch(password))
            {
                result = true;
            }

            return result;
        }
        public static bool ProductValueCheck(string password)
        {
            bool result = false;

            Regex rgx = new Regex(@"^[0-9]*$");

            if (rgx.IsMatch(password))
            {
                result = true;
            }

            return result;
        }
        public static bool ConsigneeStateCheck(string password)
        {
            bool result = false;

            Regex rgx = new Regex(@"^[A-Z]{2}$");

            if (rgx.IsMatch(password))
            {
                result = true;
            }

            return result;
        }
        public static bool ShipperStateCheck(string password)
        {
            bool result = false;

            Regex rgx = new Regex(@"^[A-Z]{3}$");

            if (rgx.IsMatch(password))
            {
                result = true;
            }

            return result;
        }

        public static bool ShipperCountryCheck(string password)
        {
            bool result = false;

            Regex rgx = new Regex(@"^[A-Z]{2}$");

            if (rgx.IsMatch(password))
            {
                result = true;
            }

            return result;
        }

        public static bool ShipperNameCheck(string password)
        {
            bool result = false;

            Regex rgx = new Regex(@"^[A-Z_ ]*$");

            if (rgx.IsMatch(password))
            {
                result = true;
            }

            return result;
        }
        public static bool ConsigneeAddressCheck(string password)
        {
            bool result = false;

            Regex rgx = new Regex(@"^[0-9]+[ ][A-Z0-9][A-Z0-9 ]*$");

            if (rgx.IsMatch(password))
            {
                result = true;
            }

            return result;
        }
        //Consignee Address
        public static void LogToFile(string message, string time_stamp)
        {//DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt")
            using (StreamWriter sw = new StreamWriter(@"C:\Users\PJUAREZ\source\repos\cap_revision_sec321\cap_revision_sec321\bin\Debug\netcoreapp3.1\Files\Logs.txt", true))
            {
                sw.WriteLine(time_stamp + " " + message);
            }
        }
        public static void LogToScreen(string message, string time_stamp)
        {
            Console.WriteLine(time_stamp + " " + message);
        }

        public static void Log(string message)
        {
            string time_stamp = DateTime.Now.ToString("dd'/'MM'/'yyyy HH:mm:ss");
            LogToFile(message, time_stamp);
            LogToScreen(message, time_stamp);
        }
    }
    
}

