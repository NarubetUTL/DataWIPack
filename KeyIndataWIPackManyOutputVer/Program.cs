using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDataReader;
using System.IO;
using System.Data;
using ClosedXML.Excel;


namespace KeyIndataWIPackManyOutputVer
{
    class Program
    {
        private void Runtop()
        {
            DataTable dt_result = new DataTable();
            try
            {
                dt_result.Columns.Add("WI_PACK_ID");
                dt_result.Columns.Add("WI_TYPE");
                dt_result.Columns.Add("ENGINEERING_CODE");
                dt_result.Columns.Add("PRODUCT_APPLICATION");
                dt_result.Columns.Add("DESCRIPTION");
                dt_result.Columns.Add("INSTRUC_OPTN");
                dt_result.Columns.Add("HTB");
                dt_result.Columns.Add("SPECIAL_MAT1");
                dt_result.Columns.Add("SPECIAL_MAT2");
                dt_result.Columns.Add("SPECIAL_MAT3");
                dt_result.Columns.Add("SPECIAL_MAT4");
                dt_result.Columns.Add("SPECIAL_MAT5");
                dt_result.Columns.Add("SPECIAL_MAT1_QTY");
                dt_result.Columns.Add("SPECIAL_MAT2_QTY");
                dt_result.Columns.Add("SPECIAL_MAT3_QTY");
                dt_result.Columns.Add("SPECIAL_MAT4_QTY");
                dt_result.Columns.Add("SPECIAL_MAT5_QTY");
                dt_result.Columns.Add("BAKE_TEMP");
                dt_result.Columns.Add("BAKE_DURATION");
                dt_result.Columns.Add("BAKE_TOOLING");
                dt_result.Columns.Add("PIN1_ORIENTATION");
                dt_result.Columns.Add("PIN1_ORIENTATION_IMG_PATH");
                dt_result.Columns.Add("UNIT_PER_REEL");
                dt_result.Columns.Add("UNIT_PLACEMENT");
                dt_result.Columns.Add("LEADER_POCKET_MAX");
                dt_result.Columns.Add("LEADER_POCKET_MIN");
                dt_result.Columns.Add("TRAILER_POCKET_MAX");
                dt_result.Columns.Add("TRAILER_POCKET_MIN");
                dt_result.Columns.Add("ATTACH_LABEL_FLAG");
                dt_result.Columns.Add("LABEL_POSITION");
                dt_result.Columns.Add("TRAILER_TAPE_FLAG");
                dt_result.Columns.Add("UNIT_PER_TUBE");
                dt_result.Columns.Add("P1_FULL_TUBE");
                dt_result.Columns.Add("P1_FULL_TUBE_FOAM");
                dt_result.Columns.Add("OP_P1_FULL_TUBE");
                dt_result.Columns.Add("OP_P1_FULL_TUBE_FOAM");
                dt_result.Columns.Add("P1_COMBINE_TUBE");
                dt_result.Columns.Add("P1_COMBINE_TUBE_FOAM");
                dt_result.Columns.Add("OP_P1_COMBINE_TUBE");
                dt_result.Columns.Add("OP_P1_COMBINE_TUBE_FOAM");
                dt_result.Columns.Add("P1_PARTIAL_TUBE");
                dt_result.Columns.Add("P1_PARTIAL_TUBE_FOAM");
                dt_result.Columns.Add("OP_P1_PARTIAL_TUBE");
                dt_result.Columns.Add("OP_P1_PARTIAL_TUBE_FOAM");
                dt_result.Columns.Add("PIN1_POSITION");
                dt_result.Columns.Add("PIN1_POSITION_IMG_PATH");
                dt_result.Columns.Add("UNIT_PER_TRAY");
                dt_result.Columns.Add("QTY_COVER_TOP_SIDE");
                dt_result.Columns.Add("QTY_COVER_BOTTOM_SIDE");
                dt_result.Columns.Add("QTY_STACK_TRAY");
                dt_result.Columns.Add("PIN1_ON_TRAY");
                dt_result.Columns.Add("PIN1_ON_TRAY_IMG_PATH");
                dt_result.Columns.Add("PARTIAL_TRAY_DIRECTION");
                dt_result.Columns.Add("PARTIAL_TRAY_IMG_PATH");
                dt_result.Columns.Add("UNIT_PER_CAN");
                dt_result.Columns.Add("CUSHION_POSITION");
                dt_result.Columns.Add("SHIELDING_BAG");
                dt_result.Columns.Add("UNIT_PER_BAG");
                dt_result.Columns.Add("UNIT_PER_WAFER_BOX");
                dt_result.Columns.Add("PACKOUT_TYPE");
                dt_result.Columns.Add("L1_UNIT_PER_REEL");
                dt_result.Columns.Add("L1_UNIT_PER_TUBE");
                dt_result.Columns.Add("L1_UNIT_PER_TRAY");
                dt_result.Columns.Add("L1_UNIT_PER_CAN");
                dt_result.Columns.Add("L1_UNIT_PER_WF_BOX");
                dt_result.Columns.Add("L1_UNIT_PER_BAG");
                dt_result.Columns.Add("L1_CUST_LABEL_FLAG");
                dt_result.Columns.Add("L1_CUST_LABEL_QTY");
                dt_result.Columns.Add("L1_ESD_FLAG");
                dt_result.Columns.Add("L1_ESD_QTY");
                dt_result.Columns.Add("L1_PROTECTIVE_FLAG");
                dt_result.Columns.Add("L1_PROTECTIVE_QTY");
                dt_result.Columns.Add("L1_PINK_FOAM_FLAG");
                dt_result.Columns.Add("L1_WRAP_RUBBER_FLAG");
                dt_result.Columns.Add("L1_WRAP_BUBBLE_FLAG");
                dt_result.Columns.Add("L1_QTY_PER_WRAP_RUBBER");
                dt_result.Columns.Add("L1_QTY_TACK_TRAY_FLAG");
                dt_result.Columns.Add("L1_QTY_COVER_TRAY");
                dt_result.Columns.Add("L1_QTY_STRAP_TRAY");
                dt_result.Columns.Add("L2_UNIT_PER_BAG");
                dt_result.Columns.Add("L2_UNIT_PER_CAN");
                dt_result.Columns.Add("L2_QTY_TUBE_PER_BAG");
                dt_result.Columns.Add("L2_QTY_WF_BOX_PER_BAG");
                dt_result.Columns.Add("L2_QTY_REEL_PER_BAG");
                dt_result.Columns.Add("L2_QTY_BAG_PER_BAG");
                dt_result.Columns.Add("L2_DRY_PACK_FLAG");
                dt_result.Columns.Add("L2_CACUUM_SEAL_FLAG");
                dt_result.Columns.Add("L2_SEAL_LINE");
                dt_result.Columns.Add("L2_MET");
                dt_result.Columns.Add("L2_CUST_LABEL_FLAG");
                dt_result.Columns.Add("L2_CUST_LABEL_QTY");
                dt_result.Columns.Add("L2_ESD_FLAG");
                dt_result.Columns.Add("L2_ESD_QTY");
                dt_result.Columns.Add("L2_CAUTION_FLAG");
                dt_result.Columns.Add("L2_CAUTION_QTY");
                dt_result.Columns.Add("L2_HIC_FLAG");
                dt_result.Columns.Add("L2_HIC_QTY");
                dt_result.Columns.Add("L2_DESICCANT_FLAG");
                dt_result.Columns.Add("L2_DESICCANT_QTY");
                dt_result.Columns.Add("L3_QTY_UNIT_PER_BOX");
                dt_result.Columns.Add("L3_QTY_TUBE_PER_BOX");
                dt_result.Columns.Add("L3_QTY_BAG_PER_BOX");
                dt_result.Columns.Add("L3_QTY_REEL_PER_BOX");
                dt_result.Columns.Add("L3_QTY_WF_BOX_PER_BOX");
                dt_result.Columns.Add("L3_CUST_LABEL_FLAG");
                dt_result.Columns.Add("L3_CUST_LABEL_QTY");
                dt_result.Columns.Add("L3_ESD_FLAG");
                dt_result.Columns.Add("L3_ESD_QTY");
                dt_result.Columns.Add("L3_BUBBLE_FLAG");
                dt_result.Columns.Add("L3_BUBBLE_QTY");
                dt_result.Columns.Add("L3_CAUTION_FLAG");
                dt_result.Columns.Add("L3_CAUTION_QTY");
                dt_result.Columns.Add("L3_QTY_TAPE_LINE");
                dt_result.Columns.Add("UNIQUE_ID");
                dt_result.Columns.Add("STATUS");
                dt_result.Columns.Add("CREATED_BY");
                dt_result.Columns.Add("CREATED_BY_NAME");
                dt_result.Columns.Add("CREATED_DATE");
                dt_result.Columns.Add("UPDATED_BY");
                dt_result.Columns.Add("UPDATED_BY_NAME");
                dt_result.Columns.Add("UPDATED_DATE");
                dt_result.Columns.Add("SPECIAL_MAT6");
                dt_result.Columns.Add("SPECIAL_MAT7");
                dt_result.Columns.Add("SPECIAL_MAT8");
                dt_result.Columns.Add("SPECIAL_MAT9");
                dt_result.Columns.Add("SPECIAL_MAT10");
                dt_result.Columns.Add("SPECIAL_MAT6_QTY");
                dt_result.Columns.Add("SPECIAL_MAT7_QTY");
                dt_result.Columns.Add("SPECIAL_MAT8_QTY");
                dt_result.Columns.Add("SPECIAL_MAT9_QTY");
                dt_result.Columns.Add("SPECIAL_MAT10_QTY");
                dt_result.Columns.Add("L1_UNIT_QTY");
                dt_result.Columns.Add("L2_UNIT_QTY");
                dt_result.Columns.Add("L2_PACK_QTY");
                dt_result.Columns.Add("L3_UNIT_QTY");
                dt_result.Columns.Add("L3_PACK_QTY");

            }
            catch (Exception ex)
            {

                Console.WriteLine(ex);
                Console.ReadLine();
            }
        }

        static void Main(string[] args)
        {
            try
            {
                string mytime = DateTime.Now.ToString("R");
                mytime = mytime.Substring(5);
                string[] timeS = mytime.Split(' ');
                timeS[1] = timeS[1].ToUpper();
                timeS[2] = Convert.ToString(Convert.ToInt32(timeS[2]) % 100);
                var timeList = timeS.ToList();
                timeList.Remove(timeS[3]);
                timeList.Remove(timeS[4]);
                mytime = String.Join("-", timeList);
                Console.WriteLine(mytime);
                var filePath = @"F:\UtacCoop\key-in data WI-Pack\SOURCE ACTL pack_kc0_ob1.xlsx";
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {

                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                        {
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                            {
                                UseHeaderRow = true
                            }
                        });


                        var dt = result.Tables[0];
                        var dt_order = dt.AsEnumerable()

                                         .CopyToDataTable();

                        DataTable dtMain = dt_order;
                        Console.WriteLine("INPUT SUCCESS");
                        Console.WriteLine("Press Enter to Continue");
                        Console.ReadLine();

                        try
                        {

                            /* var dt_result = GetDatatable();*///get new Head Columns

                            //DataTable dt_resultAssy = new DataTable();
                            //DataTable dt_resultRNTO = new DataTable();
                            //DataTable dt_resultRNTF = new DataTable();
                            DataTable dt_result = new DataTable();
                            try
                            {
                                dt_result.Columns.Add("WI_PACK_ID");
                                dt_result.Columns.Add("WI_TYPE");
                                dt_result.Columns.Add("ENGINEERING_CODE");
                                dt_result.Columns.Add("PRODUCT_APPLICATION");
                                dt_result.Columns.Add("DESCRIPTION");
                                dt_result.Columns.Add("INSTRUC_OPTN");
                                dt_result.Columns.Add("HTB");
                                dt_result.Columns.Add("SPECIAL_MAT1");
                                dt_result.Columns.Add("SPECIAL_MAT2");
                                dt_result.Columns.Add("SPECIAL_MAT3");
                                dt_result.Columns.Add("SPECIAL_MAT4");
                                dt_result.Columns.Add("SPECIAL_MAT5");
                                dt_result.Columns.Add("SPECIAL_MAT1_QTY");
                                dt_result.Columns.Add("SPECIAL_MAT2_QTY");
                                dt_result.Columns.Add("SPECIAL_MAT3_QTY");
                                dt_result.Columns.Add("SPECIAL_MAT4_QTY");
                                dt_result.Columns.Add("SPECIAL_MAT5_QTY");
                                dt_result.Columns.Add("BAKE_TEMP");
                                dt_result.Columns.Add("BAKE_DURATION");
                                dt_result.Columns.Add("BAKE_TOOLING");
                                dt_result.Columns.Add("PIN1_ORIENTATION");
                                dt_result.Columns.Add("PIN1_ORIENTATION_IMG_PATH");
                                dt_result.Columns.Add("UNIT_PER_REEL");
                                dt_result.Columns.Add("UNIT_PLACEMENT");
                                dt_result.Columns.Add("LEADER_POCKET_MAX");
                                dt_result.Columns.Add("LEADER_POCKET_MIN");//26
                                dt_result.Columns.Add("TRAILER_POCKET_MAX");
                                dt_result.Columns.Add("TRAILER_POCKET_MIN");
                                dt_result.Columns.Add("ATTACH_LABEL_FLAG");
                                dt_result.Columns.Add("LABEL_POSITION");
                                dt_result.Columns.Add("TRAILER_TAPE_FLAG");
                                dt_result.Columns.Add("UNIT_PER_TUBE");
                                dt_result.Columns.Add("P1_FULL_TUBE");
                                dt_result.Columns.Add("P1_FULL_TUBE_FOAM");
                                dt_result.Columns.Add("OP_P1_FULL_TUBE");
                                dt_result.Columns.Add("OP_P1_FULL_TUBE_FOAM");
                                dt_result.Columns.Add("P1_COMBINE_TUBE");
                                dt_result.Columns.Add("P1_COMBINE_TUBE_FOAM");
                                dt_result.Columns.Add("OP_P1_COMBINE_TUBE");
                                dt_result.Columns.Add("OP_P1_COMBINE_TUBE_FOAM");
                                dt_result.Columns.Add("P1_PARTIAL_TUBE");
                                dt_result.Columns.Add("P1_PARTIAL_TUBE_FOAM");
                                dt_result.Columns.Add("OP_P1_PARTIAL_TUBE");
                                dt_result.Columns.Add("OP_P1_PARTIAL_TUBE_FOAM");
                                dt_result.Columns.Add("PIN1_POSITION");
                                dt_result.Columns.Add("PIN1_POSITION_IMG_PATH");
                                dt_result.Columns.Add("UNIT_PER_TRAY");
                                dt_result.Columns.Add("QTY_COVER_TOP_SIDE");
                                dt_result.Columns.Add("QTY_COVER_BOTTOM_SIDE");
                                dt_result.Columns.Add("QTY_STACK_TRAY");
                                dt_result.Columns.Add("PIN1_ON_TRAY");
                                dt_result.Columns.Add("PIN1_ON_TRAY_IMG_PATH");//26*2
                                dt_result.Columns.Add("PARTIAL_TRAY_DIRECTION");
                                dt_result.Columns.Add("PARTIAL_TRAY_IMG_PATH");
                                dt_result.Columns.Add("UNIT_PER_CAN");
                                dt_result.Columns.Add("CUSHION_POSITION");
                                dt_result.Columns.Add("SHIELDING_BAG");
                                dt_result.Columns.Add("UNIT_PER_BAG");
                                dt_result.Columns.Add("UNIT_PER_WAFER_BOX");
                                dt_result.Columns.Add("PACKOUT_TYPE");
                                dt_result.Columns.Add("L1_UNIT_PER_REEL");
                                dt_result.Columns.Add("L1_UNIT_PER_TUBE");
                                dt_result.Columns.Add("L1_UNIT_PER_TRAY");
                                dt_result.Columns.Add("L1_UNIT_PER_CAN");
                                dt_result.Columns.Add("L1_UNIT_PER_WF_BOX");
                                dt_result.Columns.Add("L1_UNIT_PER_BAG");
                                dt_result.Columns.Add("L1_CUST_LABEL_FLAG");
                                dt_result.Columns.Add("L1_CUST_LABEL_QTY");
                                dt_result.Columns.Add("L1_ESD_FLAG");
                                dt_result.Columns.Add("L1_ESD_QTY");
                                dt_result.Columns.Add("L1_PROTECTIVE_FLAG");
                                dt_result.Columns.Add("L1_PROTECTIVE_QTY");
                                dt_result.Columns.Add("L1_PINK_FOAM_FLAG");
                                dt_result.Columns.Add("L1_WRAP_RUBBER_FLAG");
                                dt_result.Columns.Add("L1_WRAP_BUBBLE_FLAG");
                                dt_result.Columns.Add("L1_QTY_PER_WRAP_RUBBER");
                                dt_result.Columns.Add("L1_QTY_TACK_TRAY_FLAG");
                                dt_result.Columns.Add("L1_QTY_COVER_TRAY");//26*3
                                dt_result.Columns.Add("L1_QTY_STRAP_TRAY");
                                dt_result.Columns.Add("L2_UNIT_PER_BAG");
                                dt_result.Columns.Add("L2_UNIT_PER_CAN");
                                dt_result.Columns.Add("L2_QTY_TUBE_PER_BAG");
                                dt_result.Columns.Add("L2_QTY_WF_BOX_PER_BAG");
                                dt_result.Columns.Add("L2_QTY_REEL_PER_BAG");
                                dt_result.Columns.Add("L2_QTY_BAG_PER_BAG");
                                dt_result.Columns.Add("L2_DRY_PACK_FLAG");
                                dt_result.Columns.Add("L2_CACUUM_SEAL_FLAG");
                                dt_result.Columns.Add("L2_SEAL_LINE");
                                dt_result.Columns.Add("L2_MET");
                                dt_result.Columns.Add("L2_CUST_LABEL_FLAG");
                                dt_result.Columns.Add("L2_CUST_LABEL_QTY");
                                dt_result.Columns.Add("L2_ESD_FLAG");
                                dt_result.Columns.Add("L2_ESD_QTY");
                                dt_result.Columns.Add("L2_CAUTION_FLAG");
                                dt_result.Columns.Add("L2_CAUTION_QTY");
                                dt_result.Columns.Add("L2_HIC_FLAG");
                                dt_result.Columns.Add("L2_HIC_QTY");
                                dt_result.Columns.Add("L2_DESICCANT_FLAG");
                                dt_result.Columns.Add("L2_DESICCANT_QTY");
                                dt_result.Columns.Add("L3_QTY_UNIT_PER_BOX");
                                dt_result.Columns.Add("L3_QTY_TUBE_PER_BOX");
                                dt_result.Columns.Add("L3_QTY_BAG_PER_BOX");
                                dt_result.Columns.Add("L3_QTY_REEL_PER_BOX");
                                dt_result.Columns.Add("L3_QTY_WF_BOX_PER_BOX");//26*4
                                dt_result.Columns.Add("L3_CUST_LABEL_FLAG");
                                dt_result.Columns.Add("L3_CUST_LABEL_QTY");
                                dt_result.Columns.Add("L3_ESD_FLAG");
                                dt_result.Columns.Add("L3_ESD_QTY");
                                dt_result.Columns.Add("L3_BUBBLE_FLAG");
                                dt_result.Columns.Add("L3_BUBBLE_QTY");
                                dt_result.Columns.Add("L3_CAUTION_FLAG");
                                dt_result.Columns.Add("L3_CAUTION_QTY");
                                dt_result.Columns.Add("L3_QTY_TAPE_LINE");
                                dt_result.Columns.Add("UNIQUE_ID");
                                dt_result.Columns.Add("STATUS");
                                dt_result.Columns.Add("CREATED_BY");
                                dt_result.Columns.Add("CREATED_BY_NAME");
                                dt_result.Columns.Add("CREATED_DATE");
                                dt_result.Columns.Add("UPDATED_BY");
                                dt_result.Columns.Add("UPDATED_BY_NAME");
                                dt_result.Columns.Add("UPDATED_DATE");
                                dt_result.Columns.Add("SPECIAL_MAT6");
                                dt_result.Columns.Add("SPECIAL_MAT7");
                                dt_result.Columns.Add("SPECIAL_MAT8");
                                dt_result.Columns.Add("SPECIAL_MAT9");
                                dt_result.Columns.Add("SPECIAL_MAT10");
                                dt_result.Columns.Add("SPECIAL_MAT6_QTY");
                                dt_result.Columns.Add("SPECIAL_MAT7_QTY");
                                dt_result.Columns.Add("SPECIAL_MAT8_QTY");
                                dt_result.Columns.Add("SPECIAL_MAT9_QTY");//26*5
                                dt_result.Columns.Add("SPECIAL_MAT10_QTY");
                                dt_result.Columns.Add("L1_UNIT_QTY");
                                dt_result.Columns.Add("L2_UNIT_QTY");
                                dt_result.Columns.Add("L2_PACK_QTY");
                                dt_result.Columns.Add("L3_UNIT_QTY");
                                dt_result.Columns.Add("L3_PACK_QTY");

                                //dt_resultAssy = dt_result;
                                //dt_resultRNTF = dt_result;
                                //dt_resultRNTO = dt_result;
                            }
                            catch (Exception ex)
                            {

                                Console.WriteLine(ex);
                                Console.ReadLine();
                            }




                            string Type = "";
                            //test convert data from many row to 1 row

                            Console.WriteLine("Check Type");
                            DataRow first = dtMain.Rows[0];
                            string Ptype = first["PACK_TYPE"].ToString();
                            string FType = first["FLOW_TYPE"].ToString();
                            Console.WriteLine(Ptype);
                            Console.WriteLine(FType);
                            if (Ptype == "TRAY" && FType == "A")
                            {
                                Type = "AssyTray";
                            }
                            if (Ptype == "REEL")
                            {
                                Type = "RNT";
                            }
                            Console.WriteLine(Type);
                            Console.WriteLine("Start!!!!");
                            Console.ReadLine();


                        
                            var dr = dt_result.NewRow();
                            //var dso = dt_resultRNTO.NewRow();
                            var ds = dt_result.NewRow();
                            //var drA = dt_resultAssy.NewRow();
                            //var drRO = dt_resultRNTO.NewRow();
                            //var drRF = dt_resultRNTF.NewRow();
                            var check = false;

                            int LayoutID = 0;
                            string LayoutName = "WI";
                            int zeroLayout = 9;
                            int loop = 1;
                            string A = "";
                            string B = "";
                            foreach (DataRow dataRow in dtMain.Rows)
                            {
                                if (loop == 1)
                                {
                                    switch (dataRow["PACK_TYPE"].ToString())
                                    {
                                        case "REEL": check = true; Type = "RNT"; break;
                                        case "TRAY": check = true; Type = "AssyTray"; break;
                                        default: if (check == false) { Type = "Another"; } break;
                                    }
                                }
                                //checkloop
                                if (Type == "AssyTray")
                                {
                                    if (A == dataRow["PACK_ID"].ToString())
                                    {
                                        loop = loop + 1;
                                    }
                                    
                                    A = dataRow["PACK_ID"].ToString();

                                    if (loop == 1)
                                    {
                                        dr = dt_result.NewRow();


                                        dr["WI_TYPE"] = "Generic";
                                        dr["DESCRIPTION"] = dataRow["PACK_DESCRIPTION"].ToString();
                                        dr["INSTRUC_OPTN"] = "Pack Out";
                                        dr["L1_UNIT_PER_TRAY"] = dataRow["UNIT"].ToString();
                                        dr["L1_QTY_TACK_TRAY_FLAG"] = "No";
                                        dr["L2_DRY_PACK_FLAG"] = "No";
                                        dr["L2_CACUUM_SEAL_FLAG"] = "No";
                                        dr["L2_CUST_LABEL_FLAG"] = "No";
                                        dr["L2_ESD_FLAG"] = "No";
                                        dr["L2_CAUTION_FLAG"] = "No";
                                        dr["L2_HIC_FLAG"] = "No";
                                        dr["L3_CUST_LABEL_FLAG"] = "No";
                                        dr["L3_ESD_FLAG"] = "No";
                                        dr["L3_BUBBLE_FLAG"] = "No";
                                        dr["L3_CAUTION_FLAG"] = "No";
                                        dr["L2_DESICCANT_FLAG"] = "No";
                                        dr["UNIQUE_ID"] = "";
                                        dr["CREATED_BY"] = "System";
                                        dr["CREATED_BY_NAME"] = "System";
                                        dr["CREATED_DATE"] = mytime;
                                        dr["UPDATED_BY"] = "System";
                                        dr["UPDATED_BY_NAME"] = "System";
                                        dr["UPDATED_DATE"] = mytime;
                                        ds["UNIQUE_ID"] = "0";
                                        ds["STATUS"] = "1";
                                        ds["PACKOUT_TYPE"] = "Tray";

                                    }
                                    if (dataRow["PACK_TYPE"].ToString() == "BAG")
                                    {                                        


                                            dr["L2_QTY_REEL_PER_BAG"] = dataRow["PACK_QTY"].ToString();//Loop5
                                        dr["L3_QTY_UNIT_PER_BOX"] = dataRow["PACK_QTY"].ToString();
                                    }
                                    if (dataRow["PACK_TYPE"].ToString() == "BOX")
                                    {
                                        dr["L3_QTY_BAG_PER_BOX"] = dataRow["PACK_QTY"].ToString();//Loop9
                                        LayoutID = LayoutID + 1;
                                        dr["WI_PACK_ID"] = LayoutName + LayoutID.ToString().PadLeft(zeroLayout, '0');
                                        dt_result.Rows.Add(dr);
                                        //dt_resultAssy.Rows.Add(dr);
                                        loop = 1;
                                        check = false;
                                    }

                                }
                                if (Type == "Another")
                                {
                                    loop = 1;
                                }
                                if (Type == "RNT")
                                {
                                    if (B == dataRow["PACK_ID"].ToString())
                                    {
                                        loop = loop + 1;
                                    }
                                   
                                    B = dataRow["PACK_ID"].ToString();
                                    if (loop == 1)
                                    {
                                        dr = dt_result.NewRow();
                                        ds = dt_result.NewRow();
                                        //Operation=dr Final=ds
                                        dr["WI_TYPE"] = "Generic";
                                        //dr["DESCRIPTION"] = dataRow["PACK_ID"].ToString() + " Operation";
                                        dr["DESCRIPTION"] = dataRow["PACK_DESCRIPTION"].ToString();

                                        dr["INSTRUC_OPTN"] = dataRow["METHOD"].ToString();
                                        dr["HTB"] = dataRow["HTB"].ToString();
                                        dr["UNIT_PER_REEL"] = dataRow["UNIT"].ToString();
                                        dr["UNIT_PLACEMENT"] = "Live bug";
                                        dr["LABEL_POSITION"] = "Sprocket hole";
                                        dr["UNIQUE_ID"] = "";
                                        dr["CREATED_BY"] = "System";
                                        dr["CREATED_BY_NAME"] = "System";
                                        dr["CREATED_DATE"] = mytime;
                                        dr["UPDATED_BY"] = "System";
                                        dr["UPDATED_BY_NAME"] = "System";
                                        dr["UPDATED_DATE"] = mytime;


                                        ds["WI_TYPE"] = "Generic";
                                        ds["DESCRIPTION"] = dataRow["PACK_DESCRIPTION"].ToString();
                                        ds["INSTRUC_OPTN"] = "Pack Out";
                                        ds["PACKOUT_TYPE"] = dataRow["METHOD"].ToString();
                                        ds["L1_UNIT_PER_REEL"] = dataRow["UNIT"].ToString();
                                        ds["L1_CUST_LABEL_FLAG"] = "No";
                                        ds["L1_ESD_FLAG"] = "No";
                                        ds["L1_PROTECTIVE_FLAG"] = "No";
                                        ds["L2_DRY_PACK_FLAG"] = "No";
                                        ds["L2_CACUUM_SEAL_FLAG"] = "No";
                                        ds["L2_CUST_LABEL_FLAG"] = "No";
                                        ds["L2_ESD_FLAG"] = "No";
                                        ds["L2_CAUTION_FLAG"] = "No";
                                        ds["L2_HIC_FLAG"] = "No";
                                        ds["L2_DESICCANT_FLAG"] = "No";
                                        ds["L3_CUST_LABEL_FLAG"] = "No";
                                        ds["L3_ESD_FLAG"] = "No";
                                        ds["L3_BUBBLE_FLAG"] = "No";
                                        ds["L3_CAUTION_FLAG"] = "No";
                                        ds["UNIQUE_ID"] = "";
                                        ds["CREATED_BY"] = "System";
                                        ds["CREATED_BY_NAME"] = "System";
                                        ds["CREATED_DATE"] = mytime;
                                        ds["UPDATED_BY"] = "System";
                                        ds["UPDATED_BY_NAME"] = "System";
                                        ds["UPDATED_DATE"] = mytime;





                                        //dt_result.Rows.Add(dr);
                                        //dt_result.Rows.Add(ds);
                                    }
                                    if (dataRow["PACK_TYPE"].ToString() == "BAG")
                                    {
                                        
                                        //    ds["L2_UNIT_PER_BAG"] = dataRow["UNIT"].ToString();
                                        //ds["L2_QTY_REEL_PER_BAG"] = dataRow["PACK_QTY"].ToString();

                                    }
                                    if (dataRow["PACK_TYPE"].ToString() == "BOX")
                                    {
                                        ds["L3_QTY_UNIT_PER_BOX"] = dataRow["UNIT"].ToString();
                                        ds["L3_QTY_REEL_PER_BOX"] = dataRow["PACK_QTY"].ToString();
                                        ds["L2_UNIT_PER_BAG"] = dataRow["UNIT"].ToString();
                                        ds["L2_QTY_REEL_PER_BAG"] = dataRow["PACK_QTY"].ToString();
                                    }
                                    if (dataRow["PACK_TYPE"].ToString() == "LEADER_MIN")
                                    {
                                        dr["LEADER_POCKET_MAX"] = dataRow["PACK_QTY"].ToString();
                                        dr["LEADER_POCKET_MIN"] = dataRow["PACK_QTY"].ToString();
                                    }
                                    if (dataRow["PACK_TYPE"].ToString() == "TRAILER_MIN")
                                    {
                                        dr["TRAILER_POCKET_MAX"] = dataRow["PACK_QTY"].ToString();
                                        dr["TRAILER_POCKET_MIN"] = dataRow["PACK_QTY"].ToString();

                                    }
                                    if (dataRow["PACK_TYPE"].ToString() == "QUADRANT")
                                    {
                                        switch (dataRow["STOCK_NO"].ToString())
                                        {
                                            case "QUAD_1": dr["PIN1_ORIENTATION"] = "Quadrant 1"; break;
                                            case "QUAD_2": dr["PIN1_ORIENTATION"] = "Quadrant 2"; break;
                                            case "QUAD_3": dr["PIN1_ORIENTATION"] = "Quadrant 3"; break;
                                            case "QUAD_4": dr["PIN1_ORIENTATION"] = "Quadrant 4"; break;

                                        }
                                        LayoutID = LayoutID + 1;
                                        dr["WI_PACK_ID"] = LayoutName + LayoutID.ToString().PadLeft(zeroLayout, '0');
                                        dt_result.Rows.Add(dr);
                                        LayoutID = LayoutID + 1;
                                        ds["WI_PACK_ID"] = LayoutName + LayoutID.ToString().PadLeft(zeroLayout, '0');
                                        dt_result.Rows.Add(ds);
                                        loop = 1;
                                        check = false;

                                    }
                                }
                            }
                            //if (lastSTAT == "AssyTray")
                            //{
                            //    dt_result.Rows.Add(dr);

                            //}
                            //else
                            //{
                            //    dt_result.Rows.Add(dr);
                            //    dt_result.Rows.Add(ds);
                            //}

                            Console.WriteLine("Finish?");
                            Console.ReadLine();
                            Console.WriteLine("Generate File Output");

                            string filename;
                            Console.Write("Enter File Name:");
                            filename = Convert.ToString(Console.ReadLine());

                            Console.ReadLine();

                            var filespathOutput = @"F:\UtacCoop\key-in data WI-Pack\OUTPUT";
                            try
                            {
                                using (var workbook = new XLWorkbook())

                                {

                                    var worksheet = workbook.Worksheets.Add(dt_result, "result_part_mark");

                                    var fullpath = filespathOutput + "\\" + filename + ".xlsx";

                                    //MessageBox.Show(fullpath);
                                    workbook.SaveAs(fullpath);


                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message.ToString());
                            }

                            Console.WriteLine("Your File Name Output in " + filespathOutput + "\\" + filename + ".xlsx");
                            Console.WriteLine("All Session Has Completed");
                            Console.ReadLine();
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Stage 2 ERROR bc = " + ex.Message);
                            Console.WriteLine("Press Enter to Close");
                            Console.ReadLine();
                        }

                    }

                }
            }
            catch (Exception ex)
            {

                Console.WriteLine("Stage 1 ERROR bc = " + ex.Message);
                Console.WriteLine("Press Enter to Close");
                Console.ReadLine();

            }
        }
    }
}
