using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace RQ4RMA
{
    class RQ4RMA
    {
        [STAThread]
        static void Main(string[] args)
        {
            //string RQ4_SERVICE_URL = "https://vmi7.iqmetrix.net/VMIService.asmx"; // TEST
            string RQ4_SERVICE_URL = "https://vmi1.iqmetrix.net/VMIService.asmx"; // PRODUCTION
            //Guid RQ4_CLIENT_ID = new Guid("{aac62064-a98d-4c1e-95ac-6fc585c6bef8}"); // PRODUCTION ARCH
            Guid RQ4_CLIENT_ID = new Guid("{23f1f07a-3a9c-4ea6-bc9a-efa628afb56d}"); // PRODUCTION LIPSEY
            //Guid RQ4_CLIENT_ID = new Guid("{aad193b3-f55a-43ef-b28a-c796c1e608de}"); // PRODUCTION CA WIRELESS
            //Guid RQ4_CLIENT_ID = new Guid("{ab3d009d-77ee-4df8-b10e-f651f344a218}"); // PRODUCTION MAYCOM
            //Guid RQ4_CLIENT_ID = new Guid("{6a443ad0-b057-4330-b4e6-8b0579a52863}"); // SOLUTIONS CENTER
            //Guid RQ4_CLIENT_ID = new Guid("{fb6ed46e-5bd9-445c-9308-6161a504a933}"); // PRODUCTION NAWS
            //Guid RQ4_CLIENT_ID = new Guid("{54d739a1-3ce9-425e-bb28-cf934209f088}"); // PRODUCTION TOUCHTEL
            //Guid RQ4_CLIENT_ID = new Guid("{3949d039-069a-4ea0-a616-9adafcf553d7}"); //PRODUCTION MOBILE CENTER

            string infile = null;
            string outfile = null;
            string errfile = null;

            using (OpenFileDialog ofd = new OpenFileDialog()) {
                ofd.Title = "input csv";
                ofd.Filter = "csv files (*.csv)|*.csv";

                if (ofd.ShowDialog() == DialogResult.OK) {
                    infile = ofd.FileName;
                }
            }

            if ((infile == null) || (!System.IO.File.Exists(infile))) return;
            outfile = infile + "-out.csv";
            if ((outfile == null) || (outfile == "")) return;
            errfile = infile + "-error.csv";
            if ((errfile == null) || (errfile == null)) return;

            Dictionary<String, RMA> rmas = new Dictionary<string, RMA>();

            System.IO.StreamWriter LOG = new System.IO.StreamWriter(outfile);
            LOG.AutoFlush = true;
            LOG.WriteLine("id,po");

            Microsoft.VisualBasic.FileIO.TextFieldParser csvfile = new Microsoft.VisualBasic.FileIO.TextFieldParser(infile);
            csvfile.Delimiters = new string[]{ "," };
            csvfile.TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited;

            csvfile.ReadFields(); // headers

            // Internal ID	Document Number	PO/Check Number	Memo (Main)	Display Name	Quantity	Item Rate	storeid	rqsku	rqid

            while (!csvfile.EndOfData) {
                String[] fields = csvfile.ReadFields();

                if ((fields != null) && (fields.Length == 10)) { 
                    if (!rmas.ContainsKey(fields[1])) {
                        RMA r = new RMA();
                        r.OrderID = fields[0];
                        r.OrderNumber = fields[1];
                        r.PONumber = fields[2];
                        r.Memo = fields[3];
                        r.StoreID = fields[7];
                        rmas.Add(r.OrderNumber, r);
                    }

                    RMAItem ri = new RMAItem();
                    ri.VendorSKU = fields[4];
                    ri.Quantity = Math.Abs(Convert.ToInt16(fields[5]));
                    ri.ItemRate = Convert.ToDecimal(fields[6]);
                    ri.RQ4SKU = fields[8];
                    ri.RQ4ItemID = fields[9];

                    rmas[fields[1]].ItemList.Add(ri);
                }
            }

            try {
                csvfile.Close();
            } catch {
            }

            Console.WriteLine("RMAS FOUND: " + rmas.Count);
            Console.WriteLine();
            Console.WriteLine("PUSH ENTER TO START");
            Console.ReadLine();

            if (rmas.Count == 0) return;

            int rmanum = 1;

            RQ4.VMIService vmi = new RQ4.VMIService();
            vmi.CookieContainer = new System.Net.CookieContainer();
            vmi.Url = RQ4_SERVICE_URL;

            RQ4.VendorIdentity vid = new RQ4.VendorIdentity();
            vid.Username = "c2w1r3l355";
            vid.Password = "acc350r135";
            vid.VendorID = new Guid("F9B982C3-E7B1-4FD5-9C24-ABE752E058C7");
            vid.Client = new RQ4.ClientAgent();
            vid.Client.ClientID = RQ4_CLIENT_ID;
           
            foreach (KeyValuePair<string,RMA> kvp in rmas) {
                Console.Write(rmanum + "/" + rmas.Count + " - " + kvp.Key + " -> ");

                vid.Client.StoreID = Convert.ToInt16(kvp.Value.StoreID);

                RQ4.ReturnMerchandiseAuthorization rma = new RQ4.ReturnMerchandiseAuthorization();
                rma.VendorRMANumber = kvp.Value.OrderNumber;
                rma.Comments = kvp.Value.OrderNumber; // <--------- required

                List<RQ4.RMAProduct> items = new List<RQ4.RMAProduct>();

                foreach(RMAItem item in kvp.Value.ItemList) {
                    RQ4.RMAProduct prod = new RQ4.RMAProduct { RQProductID = Convert.ToInt16(item.RQ4ItemID), RQProductSku = item.RQ4SKU, TotalQuantity = item.Quantity, NonSellableQuantity = 0, UnitCost = item.ItemRate, ActionTaken = RQ4.ActionTaken.Credit };
                    items.Add(prod);
                }

                rma.ProductData = items.ToArray();

                RQ4.ReturnMerchandiseAuthorization outrma = null;
                
                try {
                    outrma = vmi.CreateRMA(vid, rma);

                    Console.WriteLine(outrma.RMAIDByStore);

                    LOG.WriteLine(kvp.Value.OrderID + "," + outrma.RMAIDByStore);
                    LOG.Flush();
                } catch (System.Web.Services.Protocols.SoapException se) {
                    Console.WriteLine("error");
                    
                    WriteOrder(kvp.Value, errfile);

                    LOG.WriteLine(kvp.Value.OrderID + "," + Quote(se.Message).Replace("\r", "").Replace("\n", ""));
                    LOG.Flush();
                } catch (Exception ex) {
                    Console.WriteLine("error");

                    WriteOrder(kvp.Value, errfile);

                    LOG.WriteLine(kvp.Value.OrderID + ",error");
                    LOG.Flush();
                }

                rmanum++;

                System.Threading.Thread.Sleep(250);
            }

            try {
                LOG.Flush();
                LOG.Close();
            } catch {
            }

            Console.WriteLine();
            Console.WriteLine("done");
            Console.ReadLine();
        }

        static string Quote(string what)
        {
            return "\"" + what + "\"";
        }

        static void WriteOrder(RMA r, string errfile)
        {
            try {
                using (System.IO.StreamWriter fout = new System.IO.StreamWriter(errfile, true))
                {
                    foreach (RMAItem ri in r.ItemList) {
                        fout.Write(Quote(r.OrderID));
                        fout.Write(",");
                        fout.Write(Quote(r.OrderNumber));
                        fout.Write(",");
                        fout.Write(Quote(r.PONumber));
                        fout.Write(",");
                        fout.Write(Quote(r.Memo));
                        fout.Write(",");
                        fout.Write(Quote(ri.VendorSKU));
                        fout.Write(",");
                        fout.Write(Quote(Convert.ToString(ri.Quantity)));
                        fout.Write(",");
                        fout.Write(Quote(Convert.ToString(ri.ItemRate)));
                        fout.Write(",");
                        fout.Write(Quote(r.StoreID));
                        fout.Write(",");
                        fout.Write(Quote(ri.RQ4SKU));
                        fout.Write(",");
                        fout.WriteLine(Quote(ri.RQ4ItemID));
                    }

                    fout.Flush();
                    fout.Close();
                }
            } catch {
            }
        }
    }

    class RMA
    {
        public string OrderID;
        public string OrderNumber;
        public string PONumber;
        public string Memo;
        public string StoreID;
        
        public List<RMAItem> ItemList = new List<RMAItem>();
    }

    class RMAItem
    {
        public string VendorSKU;
        public int Quantity;
        public decimal ItemRate;
        public string RQ4SKU;
        public string RQ4ItemID;
    }
}