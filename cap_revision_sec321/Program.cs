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
            //C:\Users\PJUAREZ\source\repos\cap_revision_sec321\cap_revision_sec321\bin\Debug\netcoreapp3.1\
            string FilePath = @"files\PLUMA_NACIONAL.xlsx";
            string Sheet = "Shipments";

            List<Shipment> lstShipments = Shipment.Import_To_Grid(FilePath, Sheet);

            lstShipments.ForEach(oShipment => {

                bool shipment_control_check = ShipmentControlNumberCheck(oShipment.Shipment_Control_Number);

                if (!shipment_control_check)
                {
                    Log("ERROR IN LINE " + oShipment.Id + " DETECTED. VALUE: " + oShipment.Shipment_Control_Number + " - ShipmentControlNumber");
                }

                bool shipment_type_check = ShipmentTypeCheck(oShipment.Shipment_Type);

                if (!shipment_type_check)
                {
                    Log("ERROR IN LINE " + oShipment.Id + " DETECTED. VALUE: " + oShipment.Shipment_Type + " - ShipmentType");
                }

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
                
                bool shipper_port_of_landing_check = ShipperPortofLadingCheck(oShipment.Shipper_Port_of_Lading);

                if (!shipper_port_of_landing_check)
                {
                    Log("ERROR IN LINE " + oShipment.Id + " DETECTED. VALUE: " + oShipment.Shipper_Port_of_Lading + " - ShipperPortofLading");
                }

                bool product_value_check = ProductValueCheck(oShipment.Product_Value);

                if (!shipper_postal_check)
                {
                    Log("ERROR IN LINE " + oShipment.Id + " DETECTED. VALUE: " + oShipment.Product_Value + " - ProductValue");
                }

                bool customer_reference_check = CustomerReferenceCheck(oShipment.Customer_Reference);

                if (!customer_reference_check)
                {
                    Log("ERROR IN LINE " + oShipment.Id + " DETECTED. VALUE: " + oShipment.Customer_Reference + " - CustomerReference");
                }
                
                bool us_port_arrive_check = USPortArriveCheck(oShipment.US_Port_Arrive);

                if (!us_port_arrive_check)
                {
                    Log("ERROR IN LINE " + oShipment.Id + " DETECTED. VALUE: " + oShipment.US_Port_Arrive + " - USPortArrive");
                }
                
                bool fn_port_loading_check = FnPortLoadingCheck(oShipment.Fn_Port_Loading);

                if (!fn_port_loading_check)
                {
                    Log("ERROR IN LINE " + oShipment.Id + " DETECTED. VALUE: " + oShipment.Fn_Port_Loading + " - FnPortLoading");
                }
                
                bool fn_port_reciept_check = FnPortRecieptCheck(oShipment.Fn_Port_Reciept);

                if (!fn_port_reciept_check)
                {
                    Log("ERROR IN LINE " + oShipment.Id + " DETECTED. VALUE: " + oShipment.Fn_Port_Reciept + " - FnPortReciept");
                }

                bool consignee_state_check = ConsigneeStateCheck(oShipment.Consignee_State);

                if (!consignee_state_check)
                {
                    Log("ERROR IN LINE " + oShipment.Id + " DETECTED. VALUE: " + oShipment.Consignee_State + " - ConsigneeState");
                }
                
                bool consignee_portal_check = ConsigneePostalCheck(oShipment.Consignee_Postal);

                if (!consignee_portal_check)
                {
                    Log("ERROR IN LINE " + oShipment.Id + " DETECTED. VALUE: " + oShipment.Consignee_Postal + " - ConsigneePostal");
                }
                
                bool product_description_check = ProductDescriptionCheck(oShipment.Product_Description);

                if (!product_description_check)
                {
                    Log("ERROR IN LINE " + oShipment.Id + " DETECTED. VALUE: " + oShipment.Product_Description + " - ProductDescription");
                }
                
                bool product_qty_check = ProductQtyCheck(oShipment.Product_Qty);

                if (!product_qty_check)
                {
                    Log("ERROR IN LINE " + oShipment.Id + " DETECTED. VALUE: " + oShipment.Product_Qty + " - ProductQty");
                }
                
                bool product_uom_check = ProductUOMCheck(oShipment.Product_UOM);

                if (!product_uom_check)
                {
                    Log("ERROR IN LINE " + oShipment.Id + " DETECTED. VALUE: " + oShipment.Product_UOM + " - ProductUOM");
                }

                bool product_weight_check = ProductWeightCheck(oShipment.Product_Weight);

                if (!product_weight_check)
                {
                    Log("ERROR IN LINE " + oShipment.Id + " DETECTED. VALUE: " + oShipment.Product_Weight + " - ProductWeight");
                }
                
                bool product_unit_of_weight_check = ProductUnitofWeightCheck(oShipment.Product_Unit_of_Weight);

                if (!product_unit_of_weight_check)
                {
                    Log("ERROR IN LINE " + oShipment.Id + " DETECTED. VALUE: " + oShipment.Product_Unit_of_Weight + " - ProductUnitofWeight");
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

                bool consignee_city_check = ConsigneeCityCheck(oShipment.Consignee_City);

                if (!consignee_city_check)
                {
                    Log("ERROR IN LINE " + oShipment.Id + " DETECTED. VALUE: " + oShipment.Consignee_City + " - ConsigneeCity");
                }

                bool consignee_country_check = ConsigneeCountryCheck(oShipment.Consignee_Country);

                if (!consignee_country_check)
                {
                    Log("ERROR IN LINE " + oShipment.Id + " DETECTED. VALUE: " + oShipment.Consignee_Country + " - ConsigneeCountry");
                }


                bool shipper_city_check = ShipperCityCheck(oShipment.Shipper_City);

                if (!shipper_city_check)
                {
                    Log("ERROR IN LINE " + oShipment.Id + " DETECTED. VALUE: " + oShipment.Shipper_City + " - ShipperCity");
                }

                bool origin_check = OriginCheck(oShipment.Origin);

                if (!origin_check)
                {
                    Log("ERROR IN LINE " + oShipment.Id + " DETECTED. VALUE: " + oShipment.Origin + " - Origin");
                }

            });

            Log(lstShipments.Count + " LINES PROCESSED.");
        }
        public static bool ShipmentControlNumberCheck(string password)
        {
            bool result = false;

            Regex rgx = new Regex(@"^[A-Z][0-9]*$");

            if (rgx.IsMatch(password))
            {
                result = true;
            }

            return result;
        }
        public static bool ShipmentTypeCheck(string password)
        {
            bool result = false;

            Regex rgx = new Regex(@"^SECTION 321$");

            if (rgx.IsMatch(password))
            {
                result = true;
            }

            return result;
        }
        public static bool ConsigneeNameCheck(string password)
        {
            bool result = false;

            Regex rgx = new Regex(@"^[a-zA-Z0-9 ]*$");

            if (rgx.IsMatch(password))
            {
                result = true;
            }

            return result;
        }
        public static bool ShipperAddressCheck(string password)
        {
            bool result = false;

            Regex rgx = new Regex(@"^[a-zA-Z0-9 ]*$");

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
        public static bool ShipperPortofLadingCheck(string password)
        {
            bool result = false;

            Regex rgx = new Regex(@"^[a-zA-Z ]*$");

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
        public static bool CustomerReferenceCheck(string password)
        {
            bool result = false;

            Regex rgx = new Regex(@"^[A-Z][0-9]*$");

            if (rgx.IsMatch(password))
            {
                result = true;
            }

            return result;
        }
        public static bool USPortArriveCheck(string password)
        {
            bool result = false;

            Regex rgx = new Regex(@"^[0-9]{4}$");

            if (rgx.IsMatch(password))
            {
                result = true;
            }

            return result;
        }
        
        public static bool FnPortLoadingCheck(string password)
        {
            bool result = false;

            Regex rgx = new Regex(@"^[0-9]{5}$");

            if (rgx.IsMatch(password))
            {
                result = true;
            }

            return result;
        }
        public static bool FnPortRecieptCheck(string password)
        {
            bool result = false;

            Regex rgx = new Regex(@"^[0-9]{5}$");

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
        public static bool ConsigneePostalCheck(string password)
        {
            bool result = false;

            Regex rgx = new Regex(@"^[0-9]{5}$");

            if (rgx.IsMatch(password))
            {
                result = true;
            }

            return result;
        }
        public static bool ProductDescriptionCheck(string password)
        {
            bool result = false;

            Regex rgx = new Regex(@"^[a-zA-Z0-9 ]*$");

            if (rgx.IsMatch(password))
            {
                result = true;
            }

            return result;
        }
        public static bool ProductQtyCheck(string password)
        {
            bool result = false;

            Regex rgx = new Regex(@"^[0-9]*$");

            if (rgx.IsMatch(password))
            {
                result = true;
            }

            return result;
        }
        public static bool ProductUOMCheck(string password)
        {
            bool result = false;

            Regex rgx = new Regex(@"^[A-Z]*$");

            if (rgx.IsMatch(password))
            {
                result = true;
            }

            return result;
        }
        
        public static bool ProductWeightCheck(string password)
        {
            bool result = false;

            Regex rgx = new Regex(@"^[0-9]*$");

            if (rgx.IsMatch(password))
            {
                result = true;
            }

            return result;
        }
        
        public static bool ProductUnitofWeightCheck(string password)
        {
            bool result = false;

            Regex rgx = new Regex(@"^[A-Z]*$");

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

            Regex rgx = new Regex(@"^[A-Z ]*$");

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
        public static bool ConsigneeCityCheck(string password)
        {
            bool result = false;

            Regex rgx = new Regex(@"^[A-Z ]*$");

            if (rgx.IsMatch(password))
            {
                result = true;
            }

            return result;
        }
        public static bool ConsigneeCountryCheck(string password)
        {
            bool result = false;

            Regex rgx = new Regex(@"^[A-Z]{2}$");

            if (rgx.IsMatch(password))
            {
                result = true;
            }

            return result;
        }
        public static bool ShipperCityCheck(string password)
        {
            bool result = false;

            Regex rgx = new Regex(@"^[A-Z]+$");

            if (rgx.IsMatch(password))
            {
                result = true;
            }

            return result;
        }
        public static bool OriginCheck(string password)
        {
            bool result = false;

            Regex rgx = new Regex(@"^[A-Z]{2}$");

            if (rgx.IsMatch(password))
            {
                result = true;
            }

            return result;
        }

        public static void LogToFile(string message, string time_stamp)
        {//DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt")
            using (StreamWriter sw = new StreamWriter(@"Files\Logs.txt", true))
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

