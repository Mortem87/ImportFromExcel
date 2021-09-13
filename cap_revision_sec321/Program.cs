using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;
using System.Text.RegularExpressions;

namespace cap_revision_sec321
{
    public class Shipment
    {
        private int Id { get; set; }
        public string ShipmentControlNumber { get; set; }
        public string ShipmentType { get; set; }
        public string ShipperName { get; set; }
        public string ShipperAddress { get; set; }
        public string ShipperCity { get; set; }
        public string ShipperCountry { get; set; }
        public string ShipperState { get; set; }
        public string ShipperPostal { get; set; }
        public string ShipperPortofLading { get; set; }
        public string ConsigneeName { get; set; }
        public string ConsigneeAddress { get; set; }
        public string ConsigneeCity { get; set; }
        public string ConsigneeCountry { get; set; }
        public string ConsigneeState { get; set; }
        public string ConsigneePostal { get; set; }
        public string ProductDescription { get; set; }
        public string ProductQty { get; set; }
        public string ProductUOM { get; set; }
        public string ProductWeight { get; set; }
        public string ProductUnitofWeight { get; set; }
        public string ProductValue { get; set; }
        public string CustomerReference { get; set; }
        public string USPortArrive { get; set; }
        public string FnPortLoading { get; set; }
        public string FnPortReciept { get; set; }
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
                            oShipment.ShipmentControlNumber = reader[1].ToString();
                            oShipment.ShipmentType = reader[2].ToString();
                            oShipment.ShipperName = reader[3].ToString();
                            oShipment.ShipperAddress = reader[4].ToString();
                            oShipment.ShipperCity = reader[5].ToString();
                            oShipment.ShipperCountry = reader[6].ToString();
                            oShipment.ShipperState = reader[7].ToString();
                            oShipment.ShipperPostal = reader[8].ToString();
                            oShipment.ShipperPortofLading = reader[9].ToString();
                            oShipment.ConsigneeName = reader[10].ToString();
                            oShipment.ConsigneeAddress = reader[11].ToString();
                            oShipment.ConsigneeCity = reader[12].ToString();
                            oShipment.ConsigneeCountry = reader[13].ToString();
                            oShipment.ConsigneeState = reader[14].ToString();
                            oShipment.ConsigneePostal = reader[15].ToString();
                            oShipment.ProductDescription = reader[16].ToString();
                            oShipment.ProductQty = reader[17].ToString();
                            oShipment.ProductUOM = reader[18].ToString();
                            oShipment.ProductWeight = reader[19].ToString();
                            oShipment.ProductUnitofWeight = reader[20].ToString();
                            oShipment.ProductValue = reader[21].ToString();
                            oShipment.CustomerReference = reader[22].ToString();
                            oShipment.USPortArrive = reader[23].ToString();
                            oShipment.FnPortLoading = reader[24].ToString();
                            oShipment.FnPortReciept = reader[25].ToString();
                            oShipment.Origin = reader[26].ToString();
                            lstShipments.Add(oShipment);

                        }

                        Count++;

                    }
                }
                return lstShipments;
            }
        }
        private static bool ShipmentControlNumberCheck(string password)
        {
            bool result = false;

            Regex rgx = new Regex(@"^[A-Z][0-9]*$");

            if (rgx.IsMatch(password))
            {
                result = true;
            }

            return result;
        }
        private static bool ShipmentTypeCheck(string password)
        {
            bool result = false;

            Regex rgx = new Regex(@"^SECTION 321$");

            if (rgx.IsMatch(password))
            {
                result = true;
            }

            return result;
        }
        private static bool ConsigneeNameCheck(string password)
        {
            bool result = false;

            Regex rgx = new Regex(@"^[a-zA-Z0-9 ]*$");

            if (rgx.IsMatch(password))
            {
                result = true;
            }

            return result;
        }
        private static bool ShipperAddressCheck(string password)
        {
            bool result = false;

            Regex rgx = new Regex(@"^[a-zA-Z0-9 ]*$");

            if (rgx.IsMatch(password))
            {
                result = true;
            }

            return result;
        }

        private static bool ShipperPostalCheck(string password)
        {
            bool result = false;

            Regex rgx = new Regex(@"^[0-9]{5}$");

            if (rgx.IsMatch(password))
            {
                result = true;
            }

            return result;
        }
        private static bool ShipperPortofLadingCheck(string password)
        {
            bool result = false;

            Regex rgx = new Regex(@"^[a-zA-Z ]*$");

            if (rgx.IsMatch(password))
            {
                result = true;
            }

            return result;
        }
        private static bool ProductValueCheck(string password)
        {
            bool result = false;

            Regex rgx = new Regex(@"^[0-9]*$");

            if (rgx.IsMatch(password))
            {
                result = true;
            }

            return result;
        }
        private static bool CustomerReferenceCheck(string password)
        {
            bool result = false;

            Regex rgx = new Regex(@"^[A-Z][0-9]*$");

            if (rgx.IsMatch(password))
            {
                result = true;
            }

            return result;
        }
        private static bool USPortArriveCheck(string password)
        {
            bool result = false;

            Regex rgx = new Regex(@"^[0-9]{4}$");

            if (rgx.IsMatch(password))
            {
                result = true;
            }

            return result;
        }

        private static bool FnPortLoadingCheck(string password)
        {
            bool result = false;

            Regex rgx = new Regex(@"^[0-9]{5}$");

            if (rgx.IsMatch(password))
            {
                result = true;
            }

            return result;
        }
        private static bool FnPortRecieptCheck(string password)
        {
            bool result = false;

            Regex rgx = new Regex(@"^[0-9]{5}$");

            if (rgx.IsMatch(password))
            {
                result = true;
            }

            return result;
        }
        private static bool ConsigneeStateCheck(string password)
        {
            bool result = false;

            Regex rgx = new Regex(@"^[A-Z]{2}$");

            if (rgx.IsMatch(password))
            {
                result = true;
            }

            return result;
        }
        private static bool ConsigneePostalCheck(string password)
        {
            bool result = false;

            Regex rgx = new Regex(@"^[0-9]{5}$");

            if (rgx.IsMatch(password))
            {
                result = true;
            }

            return result;
        }
        private static bool ProductDescriptionCheck(string password)
        {
            bool result = false;

            Regex rgx = new Regex(@"^[a-zA-Z0-9 ]*$");

            if (rgx.IsMatch(password))
            {
                result = true;
            }

            return result;
        }
        private static bool ProductQtyCheck(string password)
        {
            bool result = false;

            Regex rgx = new Regex(@"^[0-9]*$");

            if (rgx.IsMatch(password))
            {
                result = true;
            }

            return result;
        }
        private static bool ProductUOMCheck(string password)
        {
            bool result = false;

            Regex rgx = new Regex(@"^[A-Z]*$");

            if (rgx.IsMatch(password))
            {
                result = true;
            }

            return result;
        }

        private static bool ProductWeightCheck(string password)
        {
            bool result = false;

            Regex rgx = new Regex(@"^[0-9]*$");

            if (rgx.IsMatch(password))
            {
                result = true;
            }

            return result;
        }

        private static bool ProductUnitofWeightCheck(string password)
        {
            bool result = false;

            Regex rgx = new Regex(@"^[A-Z]*$");

            if (rgx.IsMatch(password))
            {
                result = true;
            }

            return result;
        }
        private static bool ShipperStateCheck(string password)
        {
            bool result = false;

            Regex rgx = new Regex(@"^[A-Z]{3}$");

            if (rgx.IsMatch(password))
            {
                result = true;
            }

            return result;
        }

        private static bool ShipperCountryCheck(string password)
        {
            bool result = false;

            Regex rgx = new Regex(@"^[A-Z]{2}$");

            if (rgx.IsMatch(password))
            {
                result = true;
            }

            return result;
        }

        private static bool ShipperNameCheck(string password)
        {
            bool result = false;

            Regex rgx = new Regex(@"^[A-Z ]*$");

            if (rgx.IsMatch(password))
            {
                result = true;
            }

            return result;
        }
        private static bool ConsigneeAddressCheck(string password)
        {
            bool result = false;

            Regex rgx = new Regex(@"^[0-9]+[ ][A-Z0-9][A-Z0-9 ]*$");

            if (rgx.IsMatch(password))
            {
                result = true;
            }

            return result;
        }
        private static bool ConsigneeCityCheck(string password)
        {
            bool result = false;

            Regex rgx = new Regex(@"^[A-Z ]*$");

            if (rgx.IsMatch(password))
            {
                result = true;
            }

            return result;
        }
        private static bool ConsigneeCountryCheck(string password)
        {
            bool result = false;

            Regex rgx = new Regex(@"^[A-Z]{2}$");

            if (rgx.IsMatch(password))
            {
                result = true;
            }

            return result;
        }
        private static bool ShipperCityCheck(string password)
        {
            bool result = false;

            Regex rgx = new Regex(@"^[A-Z]+$");

            if (rgx.IsMatch(password))
            {
                result = true;
            }

            return result;
        }
        private static bool OriginCheck(string password)
        {
            bool result = false;

            Regex rgx = new Regex(@"^[A-Z]{2}$");

            if (rgx.IsMatch(password))
            {
                result = true;
            }

            return result;
        }
        public static void Validator(List<Shipment> lstShipments)
        {
            lstShipments.ForEach(oShipment => {

                bool shipment_control_check = ShipmentControlNumberCheck(oShipment.ShipmentControlNumber);

                if (!shipment_control_check)
                {
                    Log("ERROR IN LINE " + oShipment.Id + " DETECTED. VALUE: " + oShipment.ShipmentControlNumber + " - ShipmentControlNumber");
                }

                bool shipment_type_check = ShipmentTypeCheck(oShipment.ShipmentType);

                if (!shipment_type_check)
                {
                    Log("ERROR IN LINE " + oShipment.Id + " DETECTED. VALUE: " + oShipment.ShipmentType + " - ShipmentType");
                }

                bool consignee_name_check = ConsigneeNameCheck(oShipment.ConsigneeName);

                if (!consignee_name_check)
                {
                    Log("ERROR IN LINE " + oShipment.Id + " DETECTED. VALUE: " + oShipment.ConsigneeName + " - ConsigneeName");
                }

                bool shipper_address_check = ShipperAddressCheck(oShipment.ShipperAddress);

                if (!shipper_address_check)
                {
                    Log("ERROR IN LINE " + oShipment.Id + " DETECTED. VALUE: " + oShipment.ShipperAddress + " - ShipperAddress");
                }

                bool shipper_postal_check = ShipperPostalCheck(oShipment.ShipperPostal);

                if (!shipper_postal_check)
                {
                    Log("ERROR IN LINE " + oShipment.Id + " DETECTED. VALUE: " + oShipment.ShipperPostal + " - ShipperPostal");
                }

                bool shipper_port_of_landing_check = ShipperPortofLadingCheck(oShipment.ShipperPortofLading);

                if (!shipper_port_of_landing_check)
                {
                    Log("ERROR IN LINE " + oShipment.Id + " DETECTED. VALUE: " + oShipment.ShipperPortofLading + " - ShipperPortofLading");
                }

                bool product_value_check = ProductValueCheck(oShipment.ProductValue);

                if (!shipper_postal_check)
                {
                    Log("ERROR IN LINE " + oShipment.Id + " DETECTED. VALUE: " + oShipment.ProductValue + " - ProductValue");
                }

                bool customer_reference_check = CustomerReferenceCheck(oShipment.CustomerReference);

                if (!customer_reference_check)
                {
                    Log("ERROR IN LINE " + oShipment.Id + " DETECTED. VALUE: " + oShipment.CustomerReference + " - CustomerReference");
                }

                bool us_port_arrive_check = USPortArriveCheck(oShipment.USPortArrive);

                if (!us_port_arrive_check)
                {
                    Log("ERROR IN LINE " + oShipment.Id + " DETECTED. VALUE: " + oShipment.USPortArrive + " - USPortArrive");
                }

                bool fn_port_loading_check = FnPortLoadingCheck(oShipment.FnPortLoading);

                if (!fn_port_loading_check)
                {
                    Log("ERROR IN LINE " + oShipment.Id + " DETECTED. VALUE: " + oShipment.FnPortLoading + " - FnPortLoading");
                }

                bool fn_port_reciept_check = FnPortRecieptCheck(oShipment.FnPortReciept);

                if (!fn_port_reciept_check)
                {
                    Log("ERROR IN LINE " + oShipment.Id + " DETECTED. VALUE: " + oShipment.FnPortReciept + " - FnPortReciept");
                }

                bool consignee_state_check = ConsigneeStateCheck(oShipment.ConsigneeState);

                if (!consignee_state_check)
                {
                    Log("ERROR IN LINE " + oShipment.Id + " DETECTED. VALUE: " + oShipment.ConsigneeState + " - ConsigneeState");
                }

                bool consignee_portal_check = ConsigneePostalCheck(oShipment.ConsigneePostal);

                if (!consignee_portal_check)
                {
                    Log("ERROR IN LINE " + oShipment.Id + " DETECTED. VALUE: " + oShipment.ConsigneePostal + " - ConsigneePostal");
                }

                bool product_description_check = ProductDescriptionCheck(oShipment.ProductDescription);

                if (!product_description_check)
                {
                    Log("ERROR IN LINE " + oShipment.Id + " DETECTED. VALUE: " + oShipment.ProductDescription + " - ProductDescription");
                }

                bool product_qty_check = ProductQtyCheck(oShipment.ProductQty);

                if (!product_qty_check)
                {
                    Log("ERROR IN LINE " + oShipment.Id + " DETECTED. VALUE: " + oShipment.ProductQty + " - ProductQty");
                }

                bool product_uom_check = ProductUOMCheck(oShipment.ProductUOM);

                if (!product_uom_check)
                {
                    Log("ERROR IN LINE " + oShipment.Id + " DETECTED. VALUE: " + oShipment.ProductUOM + " - ProductUOM");
                }

                bool product_weight_check = ProductWeightCheck(oShipment.ProductWeight);

                if (!product_weight_check)
                {
                    Log("ERROR IN LINE " + oShipment.Id + " DETECTED. VALUE: " + oShipment.ProductWeight + " - ProductWeight");
                }

                bool product_unit_of_weight_check = ProductUnitofWeightCheck(oShipment.ProductUnitofWeight);

                if (!product_unit_of_weight_check)
                {
                    Log("ERROR IN LINE " + oShipment.Id + " DETECTED. VALUE: " + oShipment.ProductUnitofWeight + " - ProductUnitofWeight");
                }

                bool shipper_state_check = ShipperStateCheck(oShipment.ShipperState);

                if (!shipper_state_check)
                {
                    Log("ERROR IN LINE " + oShipment.Id + " DETECTED. VALUE: " + oShipment.ShipperState + " - ShipperState");
                }

                bool shipper_country_check = ShipperCountryCheck(oShipment.ShipperCountry);

                if (!shipper_country_check)
                {
                    Log("ERROR IN LINE " + oShipment.Id + " DETECTED. VALUE: " + oShipment.ShipperCountry + " - ShipperCountry");
                }

                bool shipper_name_check = ShipperNameCheck(oShipment.ShipperName);

                if (!shipper_name_check)
                {
                    Log("ERROR IN LINE " + oShipment.Id + " DETECTED. VALUE: " + oShipment.ShipperName + " - ShipperName");
                }

                bool consignee_address_check = ConsigneeAddressCheck(oShipment.ConsigneeAddress);

                if (!consignee_address_check)
                {
                    Log("ERROR IN LINE " + oShipment.Id + " DETECTED. VALUE: " + oShipment.ConsigneeAddress + " - ConsigneeAddress");
                }

                bool consignee_city_check = ConsigneeCityCheck(oShipment.ConsigneeCity);

                if (!consignee_city_check)
                {
                    Log("ERROR IN LINE " + oShipment.Id + " DETECTED. VALUE: " + oShipment.ConsigneeCity + " - ConsigneeCity");
                }

                bool consignee_country_check = ConsigneeCountryCheck(oShipment.ConsigneeCountry);

                if (!consignee_country_check)
                {
                    Log("ERROR IN LINE " + oShipment.Id + " DETECTED. VALUE: " + oShipment.ConsigneeCountry + " - ConsigneeCountry");
                }


                bool shipper_city_check = ShipperCityCheck(oShipment.ShipperCity);

                if (!shipper_city_check)
                {
                    Log("ERROR IN LINE " + oShipment.Id + " DETECTED. VALUE: " + oShipment.ShipperCity + " - ShipperCity");
                }

                bool origin_check = OriginCheck(oShipment.Origin);

                if (!origin_check)
                {
                    Log("ERROR IN LINE " + oShipment.Id + " DETECTED. VALUE: " + oShipment.Origin + " - Origin");
                }

            });

            Log(lstShipments.Count + " LINES PROCESSED.");
        }
        public static void LogToFile(string message, string time_stamp)
        {
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

            Shipment.Validator(lstShipments);

            
        }
    }
}

