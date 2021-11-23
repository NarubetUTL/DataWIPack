using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDataReader;
using System.IO;
using ClosedXML.Excel;

namespace KeyInDataWIPackWinApp
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }
        public string mytime = "";
        private DataTable GLOBAL_DataSource = new DataTable();
        private DataTable dt_result = new DataTable();
        public string ERRORget = ""; 
        #region get Input
        private void btnBrowse_Click(object sender, EventArgs e)
        {
            string filePath = string.Empty;
            string fileExt = string.Empty;
            OpenFileDialog file = new OpenFileDialog(); 
            if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK) 
            {
                filePath = file.FileName; 
                fileExt = Path.GetExtension(filePath); 
                if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                {
                    try
                    {
                        tbBrowse.Text = filePath;

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString());
                    }
                }
                else
                {
                    MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);  
                }

            }

        }
        #endregion

        #region Read
        private void btnRead_Click(object sender, EventArgs e)
        {
            try
            {
                btnStart.Enabled = true;
                var filePath = tbBrowse.Text;
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

                        GLOBAL_DataSource = dt_order;

                        dataGridViewInput.DataSource = GLOBAL_DataSource;
                    }
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        #endregion
        private void bynStart_Click(object sender, EventArgs e)
        {
            dt_result = new DataTable();
            try
            {

                #region setcolumn
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


                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                string Type = "";

                #endregion
                #region Convert
                var dr = dt_result.NewRow();
                var ds = dt_result.NewRow();
                var check = false;
                int loop = 1;
                string A = "";
                string B = "";
                string sideTye = "";
                var checkdtAdd = true;
                var checkpassOnce = false;
                var checksideType = false;
            int LevelCOunt = 0;
                string ItemSet = "";
                int errorCount = 0;
                string getErrorPoint = "";
            string PassID = "";
            string PassSEQ = "";
            string PassDateCre = "";
            int countrow = 0;
            int countItem = 0;
            int CountRAI = 0;
            int CountWAF = 0;
            int CountTNR = 0;
            int CountTRA = 0;
            int CountCAN = 0;

            //var falsecal = false;
            foreach (DataRow dataRow in GLOBAL_DataSource.Rows)
                {
                countrow = countrow + 1;
                if (dataRow["SEQ_NO"].ToString() == "1" )
                {
                    countItem = countItem + 1;
                }
                //if (dataRow["PACK_ID"].ToString() == "ATIU008" && dataRow["PACK_REV"].ToString() == "A")
                //{
                //    string aaaad = "";
                //}
                    checksideType = false;
                    #region Last Value 1
                    if ((A != dataRow["PACK_ID"].ToString() || ItemSet !=dataRow["PACK_REV"].ToString() ) )
                    {
                        if(sideTye != "TUBE" && loop != 1)
                        {
                            checksideType = true;
                        }
                        if(sideTye != "TUBE" && PassSEQ=="1" && dataRow["SEQ_NO"].ToString() == "1")
                    {
                        checksideType = true; 
                        checkpassOnce = true;

                    }
                }

                    if(sideTye == "TUBE" &&( B != dataRow["PACK_ID"].ToString() || ItemSet != dataRow["PACK_REV"].ToString() ))
                    {
                    if (sideTye == "TUBE" && loop != 1 && B == dataRow["PACK_ID"].ToString()&& ItemSet != dataRow["PACK_REV"].ToString())
                    {
                        checksideType = true;
                    }
                    if(sideTye == "TUBE" && B != dataRow["PACK_ID"].ToString() && loop != 1)
                    {
                        checksideType = true;

                    }
                    if (sideTye == "TUBE" && PassSEQ == "1" && dataRow["SEQ_NO"].ToString() == "1")
                    {
                        checksideType = true;
                        checkpassOnce = true;

                    }
                }

                    if (checksideType)
                    {
                        if (checkdtAdd == false && checkpassOnce == true && Type != "Another")
                        {
                            
                            switch (Type)
                            {
                                case "CAN":
                                    
                                    dt_result.Rows.Add(ds);
                                    dt_result.Rows.Add(dr);
                                CountCAN =CountCAN+ 1;
                                break;
                                case "TNR":
                                    
                                    dt_result.Rows.Add(ds);
                                    dt_result.Rows.Add(dr);CountTNR = CountTNR + 1
                                    ; break;
                                case "TRA":
                                   
                                    dt_result.Rows.Add(dr); 
                                    dt_result.Rows.Add(ds);CountTRA = CountTRA + 1; break;

                                case "WAF":
                                    
                                    
                                    dt_result.Rows.Add(dr);
                                    dt_result.Rows.Add(ds);CountWAF = CountWAF + 1; break;
                                case "RAI":
                                    if (sideTye == "FILM FRAME" )
                                    {
                                        dt_result.Rows.Add(dr);

                                        dt_result.Rows.Add(ds);
                                    CountWAF = CountWAF + 1;
                                }
                                    if (sideTye == "TUBE" )
                                    {
                                        
                                        dt_result.Rows.Add(ds);
                                        dt_result.Rows.Add(dr);
                                    CountRAI = CountRAI + 1;
                                }
                               
                                break;


                            }
                            loop = 1;
                            check = false;
                            checkdtAdd = true;
                            sideTye = "";
                        Type = "Another";

                    }
                }
                #endregion

                #region checktype
                if (loop == 1 && dataRow["SEQ_NO"].ToString() == "1") //use type for case
                {
                    if (PassID == dataRow["PACK_ID"].ToString() && ItemSet == dataRow["PACK_REV"].ToString() && dataRow["CREATE_DATE"].ToString() == PassDateCre)
                    {
                        Type = "Another";
                    }
                    else { 
                    switch (dataRow["METHOD"].ToString())
                    {
                        case "TNR": checkdtAdd = false; check = true; if (dataRow["PACK_TYPE"].ToString() == "REEL") { Type = "TNR"; } else { Type = "Another2"; }; break;
                        case "TRA": checkdtAdd = false; check = true; if (dataRow["PACK_TYPE"].ToString() == "TRAY") { Type = "TRA"; } else { Type = "Another2"; }; break;
                        case "WAF":
                            checkdtAdd = false; check = true; if (dataRow["PACK_TYPE"].ToString() == "WAFER BOX") { Type = "WAF"; }
                            else
                            {
                                Type = "Another2";
                            }; break;
                        case "RAI": checkdtAdd = false; check = true; if (dataRow["PACK_TYPE"].ToString() == "TUBE" || dataRow["PACK_TYPE"].ToString() == "FILM FRAME") { Type = "RAI"; } else { Type = "Another2"; }; break;
                        case "CAN": checkdtAdd = false; check = true; if (dataRow["PACK_TYPE"].ToString() == "CANISTER") { Type = "CAN"; } else { Type = "Another2"; }; break;

                        case "":
                            if (/*dataRow["HTB"].ToString() != "" &&*/ (PassID == dataRow["PACK_ID"].ToString() && ItemSet != dataRow["PACK_REV"].ToString()) || (PassID != dataRow["PACK_ID"].ToString()) || PassID == dataRow["PACK_ID"].ToString() && ItemSet == dataRow["PACK_REV"].ToString() && dataRow["CREATE_DATE"].ToString() != PassDateCre)
                            {
                                checkdtAdd = false; check = true;
                                if (dataRow["PACK_TYPE"].ToString() == "REEL" || dataRow["PACK_TYPE"].ToString() == "TRAY" || dataRow["PACK_TYPE"].ToString() == "TUBE" || dataRow["PACK_TYPE"].ToString() == "FILM FRAME" || dataRow["PACK_TYPE"].ToString() == "CANISTER")
                                {
                                    if (dataRow["PACK_TYPE"].ToString() == "REEL")
                                    {
                                        Type = "TNR";
                                    }
                                    if (dataRow["PACK_TYPE"].ToString() == "TRAY")
                                    { Type = "TRA"; }
                                    if (dataRow["PACK_TYPE"].ToString() == "WAFER BOX")
                                    { Type = "WAF"; }
                                    if (dataRow["PACK_TYPE"].ToString() == "TUBE" || dataRow["PACK_TYPE"].ToString() == "FILM FRAME")
                                    { Type = "RAI"; }
                                    if (dataRow["PACK_TYPE"].ToString() == "CANISTER")
                                    { Type = "CAN"; }
                                }

                                else { Type = "Another2"; }
                            }
                            else { checkdtAdd = false; if (check == false) { Type = "Another"; } }
                            break;
                        default: checkdtAdd = false; if (check == false) { Type = "Another"; } break;
                    }
                }
                    }
                    if(loop==1 && dataRow["SEQ_NO"].ToString() !="1" && Type=="Another" && check == false &&
                checkdtAdd == true &&
                sideTye == "" )
                {
                    if (PassID == dataRow["PACK_ID"].ToString() && ItemSet != dataRow["PACK_REV"].ToString())
                    {
                        Type = "Another2";
                    }
                    if (PassID != dataRow["PACK_ID"].ToString())
                    {
                        Type = "Another2";

                    }
                }
                    
                    
                    #endregion

                    #region CAN
                    //ถูก 88% 
                    if (Type == "CAN")
                    {
                        if (A == dataRow["PACK_ID"].ToString() &&dataRow["PACK_REV"].ToString()==ItemSet && checkdtAdd != true)
                        {
                            loop = loop + 1;
                            checkpassOnce = true;

                        }

                        A = dataRow["PACK_ID"].ToString();
                        if (loop == 1&&dataRow["PACK_TYPE"].ToString()== "CANISTER")
                        {
                            dr = dt_result.NewRow();//OPR
                            ds = dt_result.NewRow();//PackOUT
                            checkpassOnce = false;
                            checkdtAdd = false;
                            dr["WI_PACK_ID"] = dataRow["PACK_ID"].ToString() +"_"+ dataRow["PACK_REV"].ToString() + "_OPTN";
                            ds["WI_PACK_ID"] = dataRow["PACK_ID"].ToString()+"_"+dataRow["PACK_REV"].ToString();
                            dr["WI_TYPE"] = "Generic";

                            dr["DESCRIPTION"] = dataRow["PACK_DESCRIPTION"].ToString();

                            dr["INSTRUC_OPTN"] = "Canister";
                            if (dataRow["HTB"].ToString() == "IMMED"){
                                dr["HTB"] = "P";
                            }
                            else
                            {
                                dr["HTB"] = dataRow["HTB"].ToString();
                            }
                            dr["UNIT_PER_CAN"] = dataRow["UNIT"].ToString();

                            dr["CREATED_BY"] = "System";
                            dr["CREATED_BY_NAME"] = "System";
                            dr["CREATED_DATE"] = mytime;
                            dr["UPDATED_BY"] = "System";
                            dr["UPDATED_BY_NAME"] = "System";
                            dr["UPDATED_DATE"] = mytime;
                            dr["UNIQUE_ID"] = "0";
                            dr["STATUS"] = "1";

                            ds["WI_TYPE"] = "Generic";
                            ds["DESCRIPTION"] = dataRow["PACK_DESCRIPTION"].ToString();
                            ds["INSTRUC_OPTN"] = "Pack Out";
                            ds["PACKOUT_TYPE"] = "Canister";
                            ds["L1_UNIT_PER_CAN"] = dataRow["UNIT"].ToString();
                            ds["L1_CUST_LABEL_FLAG"] = "No";
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

                            ds["CREATED_BY"] = "System";
                            ds["CREATED_BY_NAME"] = "System";
                            ds["CREATED_DATE"] = mytime;
                            ds["UPDATED_BY"] = "System";
                            ds["UPDATED_BY_NAME"] = "System";
                            ds["UPDATED_DATE"] = mytime;
                            ds["UNIQUE_ID"] = "0";
                            ds["STATUS"] = "1";
                        }
                        if (dataRow["PACK_TYPE"].ToString() == "BOX")
                        {
                            ds["L3_QTY_UNIT_PER_BOX"] = dataRow["UNIT"].ToString();

                            //if(checkdtAdd != true) { 

                            dt_result.Rows.Add(ds);
                            dt_result.Rows.Add(dr);
                        CountCAN = CountCAN + 1;
                            loop = 1;
                            check = false;
                            checkdtAdd = true;
                        Type = "Another";

                        //}
                    }
                }
                    #endregion
                    #region RAI
                    //ถูก 88%
                    if (Type == "RAI")
                    {
                        if(dataRow["PACK_TYPE"].ToString() == "FILM FRAME" || sideTye== "FILM FRAME")
                        {
                            if (A == dataRow["PACK_ID"].ToString() && dataRow["PACK_REV"].ToString() == ItemSet && checkdtAdd != true && dataRow["SEQ_NO"].ToString() != "1")
                            {
                                loop = loop + 1;
                                checkpassOnce = true;

                            }

                            A = dataRow["PACK_ID"].ToString();
                            if (loop == 1)
                            {
                                sideTye = "FILM FRAME";
                                checkpassOnce = false;

                                dr = dt_result.NewRow();
                                ds = dt_result.NewRow();
                                dr["WI_PACK_ID"] = dataRow["PACK_ID"].ToString() + "_" + dataRow["PACK_REV"].ToString();
                                ds["WI_PACK_ID"] = dataRow["PACK_ID"].ToString() + "_" + dataRow["PACK_REV"].ToString() + "_OPTN";
                            dr["WI_TYPE"] = "Generic";
                                dr["DESCRIPTION"] = dataRow["PACK_DESCRIPTION"].ToString();
                                dr["INSTRUC_OPTN"] = "Pack Out";
                                dr["PACKOUT_TYPE"] = "Wafer";
                                dr["L1_UNIT_PER_WF_BOX"] = dataRow["UNIT"].ToString();
                                dr["L2_QTY_WF_BOX_PER_BAG"] = dataRow["PACK_QTY"].ToString();

                                dr["L1_CUST_LABEL_FLAG"] = "No";

                                dr["L2_DRY_PACK_FLAG"] = "No";
                                dr["L2_CACUUM_SEAL_FLAG"] = "No";
                                dr["L2_CUST_LABEL_FLAG"] = "No";
                                dr["L2_ESD_FLAG"] = "No";
                                dr["L2_CAUTION_FLAG"] = "No";
                                dr["L2_HIC_FLAG"] = "No";
                                dr["L2_DESICCANT_FLAG"] = "No";
                                dr["L3_CUST_LABEL_FLAG"] = "No";
                                dr["L3_ESD_FLAG"] = "No";
                                dr["L3_BUBBLE_FLAG"] = "No";
                                dr["L3_CAUTION_FLAG"] = "No";
                                dr["CREATED_BY"] = "System";
                                dr["CREATED_BY_NAME"] = "System";
                                dr["CREATED_DATE"] = mytime;
                                dr["UPDATED_BY"] = "System";
                                dr["UPDATED_BY_NAME"] = "System";
                                dr["UPDATED_DATE"] = mytime;
                                dr["UNIQUE_ID"] = "0";
                                dr["STATUS"] = "1";


                                ds["WI_TYPE"] = "Generic";
                                ds["DESCRIPTION"] = dataRow["PACK_DESCRIPTION"].ToString();
                                ds["INSTRUC_OPTN"] = "Wafer";
                                ds["UNIT_PER_WAFER_BOX"] = dataRow["UNIT"].ToString(); //unit per wafer box
                                if (dataRow["HTB"].ToString() == "IMMED")
                                {
                                    ds["HTB"] = "P";
                                }
                                else
                                {
                                    if (dataRow["HTB"].ToString() == "IMMED")
                                    {
                                        ds["HTB"] = "P";
                                    }
                                    else
                                    {
                                        ds["HTB"] = dataRow["HTB"].ToString();
                                    }
                                }

                                ds["CREATED_BY"] = "System";
                                ds["CREATED_BY_NAME"] = "System";
                                ds["CREATED_DATE"] = mytime;
                                ds["UPDATED_BY"] = "System";
                                ds["UPDATED_BY_NAME"] = "System";
                                ds["UPDATED_DATE"] = mytime;
                                ds["UNIQUE_ID"] = "0";
                                ds["STATUS"] = "1";
                            }
                            if (dataRow["PACK_TYPE"].ToString() == "BOX")
                            {
                                dr["L3_QTY_WF_BOX_PER_BOX"] = dataRow["PACK_QTY"].ToString();
                                dr["L3_QTY_UNIT_PER_BOX"] = dataRow["UNIT"].ToString();
                                
                                dt_result.Rows.Add(dr);
                                dt_result.Rows.Add(ds);
                            CountWAF = CountWAF + 1;
                            loop = 1;
                                check = false;
                                checkdtAdd = true;
                                sideTye = "";
                            Type = "Another";

                        }
                    }

                        if (dataRow["PACK_TYPE"].ToString() == "TUBE" || sideTye == "TUBE")
                        {
                            if (B == dataRow["PACK_ID"].ToString() && dataRow["PACK_REV"].ToString() == ItemSet && checkdtAdd != true && sideTye!="")
                            {
                                loop = loop + 1;
                                checkpassOnce = true;

                            }

                            B = dataRow["PACK_ID"].ToString();
                            if (loop == 1)
                            {
                                sideTye = "TUBE"; checkpassOnce = false;

                                checkdtAdd = false;

                                dr = dt_result.NewRow();//OPR
                                ds = dt_result.NewRow();//PackOUT

                                dr["WI_PACK_ID"] = dataRow["PACK_ID"].ToString() + "_" + dataRow["PACK_REV"].ToString() + "_OPTN";
                            ds["WI_PACK_ID"] = dataRow["PACK_ID"].ToString() + "_" + dataRow["PACK_REV"].ToString();
                            dr["WI_TYPE"] = "Generic";
                                dr["DESCRIPTION"] = dataRow["PACK_DESCRIPTION"].ToString();
                                dr["INSTRUC_OPTN"] = "Tube";

                                if (dataRow["HTB"].ToString() == "IMMED")
                                {
                                    dr["HTB"] = "P";
                                }
                                else
                                {
                                    dr["HTB"] = dataRow["HTB"].ToString();
                                }

                                dr["UNIT_PER_TUBE"] = dataRow["UNIT"].ToString();
                                dr["P1_FULL_TUBE_FOAM"] = "No";
                                dr["OP_P1_FULL_TUBE_FOAM"] = "No";
                                dr["P1_COMBINE_TUBE_FOAM"] = "No";
                                dr["OP_P1_COMBINE_TUBE_FOAM"] = "No";
                                dr["P1_PARTIAL_TUBE_FOAM"] = "No";
                                dr["OP_P1_PARTIAL_TUBE_FOAM"] = "No";

                                dr["CREATED_BY"] = "System";
                                dr["CREATED_BY_NAME"] = "System";
                                dr["CREATED_DATE"] = mytime;
                                dr["UPDATED_BY"] = "System";
                                dr["UPDATED_BY_NAME"] = "System";
                                dr["UPDATED_DATE"] = mytime;
                                dr["UNIQUE_ID"] = "0";
                                dr["STATUS"] = "1";


                                ds["WI_TYPE"] = "Generic";
                                ds["DESCRIPTION"] = dataRow["PACK_DESCRIPTION"].ToString();
                                ds["INSTRUC_OPTN"] = "Pack Out";
                                ds["PACKOUT_TYPE"] = "Tube";
                                ds["L1_UNIT_PER_TUBE"] = dataRow["UNIT"].ToString();
                                ds["L1_PINK_FOAM_FLAG"] = "No";
                                ds["L1_WRAP_RUBBER_FLAG"] = "No";
                                ds["L1_WRAP_BUBBLE_FLAG"] = "No";
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



                                ds["CREATED_BY"] = "System";
                                ds["CREATED_BY_NAME"] = "System";
                                ds["CREATED_DATE"] = mytime;
                                ds["UPDATED_BY"] = "System";
                                ds["UPDATED_BY_NAME"] = "System";
                                ds["UPDATED_DATE"] = mytime;
                                ds["UNIQUE_ID"] = "0";
                                ds["STATUS"] = "1";
                            }
                            if (dataRow["PACK_TYPE"].ToString() == "BAG")
                            {
                                ds["L2_QTY_TUBE_PER_BAG"] = dataRow["PACK_QTY"].ToString();
                                ds["L3_QTY_TUBE_PER_BOX"] = dataRow["PACK_QTY"].ToString();


                            }
                            if (dataRow["PACK_TYPE"].ToString() == "BOX")
                            {
                                ds["L3_QTY_BAG_PER_BOX"] = dataRow["PACK_QTY"].ToString();


                                

                                dt_result.Rows.Add(ds);
                                dt_result.Rows.Add(dr);
                            CountRAI = CountRAI + 1;
                            loop = 1;
                                check = false;
                                checkdtAdd = true;
                                sideTye = "";
                            Type = "Another";

                        }
                    }
                    }
                    #endregion
                    #region WAF
                    //ถูก 60% มีข้อมูลที่ไม่มีที่มา และข้อมูลที่ไม่ถูกตามหลัก หากเพิ่มตามหลัก Database จะไม่มีปัญหา

                    if (Type == "WAF")
                    {
                        if (A == dataRow["PACK_ID"].ToString() && dataRow["PACK_REV"].ToString() == ItemSet && checkdtAdd != true && dataRow["SEQ_NO"].ToString() != "1")
                        {
                            loop = loop + 1; checkpassOnce = true;

                        }

                        A = dataRow["PACK_ID"].ToString();
                        if (loop == 1 && dataRow["PACK_TYPE"].ToString() == "WAFER BOX")
                        {
                            dr = dt_result.NewRow();
                            ds = dt_result.NewRow();//OPT
                            checkpassOnce = false;
                            checkdtAdd = false;
                            ds["WI_PACK_ID"] = dataRow["PACK_ID"].ToString() +"_" + dataRow["PACK_REV"].ToString() + "_OPTN";
                        dr["WI_PACK_ID"] = dataRow["PACK_ID"].ToString() +"_" + dataRow["PACK_REV"].ToString() ;
                        dr["WI_TYPE"] = "Generic";
                            dr["DESCRIPTION"] = dataRow["PACK_DESCRIPTION"].ToString();
                            dr["INSTRUC_OPTN"] = "Pack Out";
                            dr["PACKOUT_TYPE"] = "Wafer";
                            dr["L1_UNIT_PER_WF_BOX"] = dataRow["UNIT"].ToString();//unit per wafer box

                            dr["L1_CUST_LABEL_FLAG"] = "No";

                            dr["L2_DRY_PACK_FLAG"] = "No";
                            dr["L2_CACUUM_SEAL_FLAG"] = "No";
                            dr["L2_CUST_LABEL_FLAG"] = "No";
                            dr["L2_ESD_FLAG"] = "No";
                            dr["L2_CAUTION_FLAG"] = "No";
                            dr["L2_HIC_FLAG"] = "No";
                            dr["L2_DESICCANT_FLAG"] = "No";

                            dr["L3_CUST_LABEL_FLAG"] = "No";
                            dr["L3_ESD_FLAG"] = "No";
                            dr["L3_BUBBLE_FLAG"] = "No";
                            dr["L3_CAUTION_FLAG"] = "No";


                            ds["WI_TYPE"] = "Generic";
                            ds["DESCRIPTION"] = dataRow["PACK_DESCRIPTION"].ToString();
                            ds["INSTRUC_OPTN"] = "Wafer";
                            ds["UNIT_PER_WAFER_BOX"] = dataRow["UNIT"].ToString(); //unit per wafer box
                            if (dataRow["HTB"].ToString() == "IMMED"){
                                ds["HTB"] = "P";
                            }
                            else
                            {
                                if (dataRow["HTB"].ToString() == "IMMED")
                                {
                                    ds["HTB"] = "P";
                                }
                                else
                                {
                                    ds["HTB"] = dataRow["HTB"].ToString();
                                }
                            }




                            dr["CREATED_BY"] = "System";
                            dr["CREATED_BY_NAME"] = "System";
                            dr["CREATED_DATE"] = mytime;
                            dr["UPDATED_BY"] = "System";
                            dr["UPDATED_BY_NAME"] = "System";
                            dr["UPDATED_DATE"] = mytime;
                            dr["UNIQUE_ID"] = "0";
                            dr["STATUS"] = "1";
                            
                            ds["CREATED_BY"] = "System";
                            ds["CREATED_BY_NAME"] = "System";
                            ds["CREATED_DATE"] = mytime;
                            ds["UPDATED_BY"] = "System";
                            ds["UPDATED_BY_NAME"] = "System";
                            ds["UPDATED_DATE"] = mytime;
                            ds["UNIQUE_ID"] = "0";
                            ds["STATUS"] = "1";
                        }
                       
                        if (dataRow["PACK_TYPE"].ToString() == "BAG")
                        {
                        if (dataRow["UNIT"].ToString() != "")
                        {
                            dr["L2_QTY_WF_BOX_PER_BAG"] = Convert.ToString(Convert.ToInt32(dataRow["UNIT"].ToString()) / Convert.ToInt32(dr["L1_UNIT_PER_WF_BOX"]));
                        }
                            //dr["L2_QTY_WF_BOX_PER_BAG"] = dataRow["PACK_QTY"].ToString();
                        }
                        if (dataRow["PACK_TYPE"].ToString() == "BOX")
                        {
                        if (dataRow["UNIT"].ToString() != "")
                        {
                            dr["L3_QTY_UNIT_PER_BOX"] = dataRow["UNIT"].ToString();//unit per box

                            dr["L3_QTY_WF_BOX_PER_BOX"] = Convert.ToString(Convert.ToInt32(dr["L3_QTY_UNIT_PER_BOX"]) / Convert.ToInt32(dr["L1_UNIT_PER_WF_BOX"]));
                        }
                            //dr["L3_QTY_WF_BOX_PER_BOX"] = dataRow["PACK_QTY"].ToString();//qty

                            //if (dr["L2_QTY_WF_BOX_PER_BAG"].ToString() != "")
                            //{
                            //    dr["L3_QTY_BAG_PER_BOX"] = Convert.ToString(Convert.ToInt32(dr["L3_QTY_WF_BOX_PER_BOX"]) / Convert.ToInt32(dr["L2_QTY_WF_BOX_PER_BAG"]));
                            //}


                            //dr["WI_PACK_ID"] = "ERROR" + " Method=" + dataRow["METHOD"].ToString() + " TypeT =" + dataRow["PACK_TYPE"].ToString(); For Except Wafer





                            dt_result.Rows.Add(dr);
                            dt_result.Rows.Add(ds);//Operation

                        CountWAF = CountWAF + 1;



                        loop = 1;
                        check = false;
                            checkdtAdd = true;
                        Type = "Another";

                    }
                }
                    #endregion
                    #region TRA
                    //ถูก 60% มีข้อมูลที่ไม่ทราบที่มา และ ข้อมูล OPRT ที่ไม่มีต้นแบบ
                    if (Type == "TRA")
                    {
                        if (A == dataRow["PACK_ID"].ToString() && dataRow["PACK_REV"].ToString() == ItemSet && checkdtAdd != true && dataRow["SEQ_NO"].ToString() != "1")
                        {
                            loop = loop + 1; checkpassOnce = true;

                        }

                        A = dataRow["PACK_ID"].ToString();

                        if (loop == 1 && dataRow["PACK_TYPE"].ToString() == "TRAY")
                        {
                        LevelCOunt = 0;
                            dr = dt_result.NewRow();
                            ds = dt_result.NewRow();

                            checkpassOnce = false;
                            checkdtAdd = false;
                        dr["WI_PACK_ID"] = dataRow["PACK_ID"].ToString() + "_" + dataRow["PACK_REV"].ToString();
                        ds["WI_PACK_ID"] = dataRow["PACK_ID"].ToString() + "_" + dataRow["PACK_REV"].ToString() + "_OPTN";


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
                            dr["CREATED_BY"] = "System";
                            dr["CREATED_BY_NAME"] = "System";
                            dr["CREATED_DATE"] = mytime;
                            dr["UPDATED_BY"] = "System";
                            dr["UPDATED_BY_NAME"] = "System";
                            dr["UPDATED_DATE"] = mytime;
                            dr["UNIQUE_ID"] = "0";
                            dr["STATUS"] = "1";
                            dr["PACKOUT_TYPE"] = "Tray";

                            ds["WI_TYPE"] = "Generic";
                            ds["DESCRIPTION"] = dataRow["PACK_DESCRIPTION"].ToString();
                            ds["INSTRUC_OPTN"] = "Tray";
                            if (dataRow["HTB"].ToString() == "IMMED")
                            {
                                ds["HTB"] = "P";
                            }
                            else
                            {
                                ds["HTB"] = dataRow["HTB"].ToString();
                            }
                            ds["UNIT_PER_TRAY"] = dataRow["UNIT"].ToString();
                            //??Value
                            ds["QTY_COVER_TOP_SIDE"] = "";
                            ds["QTY_COVER_BOTTOM_SIDE"] = "";
                            ds["QTY_STACK_TRAY"] = "";
                            ds["PIN1_ON_TRAY_IMG_PATH"] = "";
                            ds["PARTIAL_TRAY_DIRECTION"] = "";
                            ds["PARTIAL_TRAY_IMG_PATH"] = "";
                            ds["PIN1_ON_TRAY"] = "";
                            //
                            ds["CREATED_BY"] = "System";
                            ds["CREATED_BY_NAME"] = "System";
                            ds["CREATED_DATE"] = mytime;
                            ds["UPDATED_BY"] = "System";
                            ds["UPDATED_BY_NAME"] = "System";
                            ds["UPDATED_DATE"] = mytime;
                            ds["UNIQUE_ID"] = "0";
                            ds["STATUS"] = "1";
                            //ds["PACKOUT_TYPE"] = "Tray";
                        LevelCOunt = LevelCOunt + 1;
                        }
                        if (dataRow["PACK_TYPE"].ToString() == "BAG")
                        {
                        LevelCOunt = LevelCOunt + 1;

                        dr["L2_QTY_REEL_PER_BAG"] = dataRow["PACK_QTY"].ToString();//this is Tray per Bag
                        if (LevelCOunt == 3)
                        {
                            if (dr["L3_QTY_BAG_PER_BOX"].ToString() != "" && dr["L2_QTY_REEL_PER_BAG"].ToString() != "")
                            {
                                dr["L3_QTY_UNIT_PER_BOX"] = Convert.ToString(Convert.ToInt32(dr["L3_QTY_BAG_PER_BOX"]) * Convert.ToInt32(dr["L2_QTY_REEL_PER_BAG"]));
                            }
                            else
                            {
                                dr["L3_QTY_UNIT_PER_BOX"] = dr["L2_QTY_REEL_PER_BAG"].ToString();
                            }
                            dt_result.Rows.Add(dr);
                            dt_result.Rows.Add(ds);
                            CountTRA = CountTRA + 1;
                            loop = 1;
                            check = false;
                            checkdtAdd = true;
                            Type = "Another";

                        }

                        //dr["L3_QTY_UNIT_PER_BOX"] = dataRow["PACK_QTY"].ToString();
                    }
                        if (dataRow["PACK_TYPE"].ToString() == "BOX")
                        {
                        LevelCOunt = LevelCOunt + 1;

                        dr["L3_QTY_BAG_PER_BOX"] = dataRow["PACK_QTY"].ToString();
                            if (dataRow["UNIT"].ToString() != "")
                            {

                            if (dr["L3_QTY_BAG_PER_BOX"].ToString() != "" && dr["L2_QTY_REEL_PER_BAG"].ToString() != "")
                            {
                                dr["L3_QTY_UNIT_PER_BOX"] = Convert.ToString(Convert.ToInt32(dr["L3_QTY_BAG_PER_BOX"]) * Convert.ToInt32(dr["L2_QTY_REEL_PER_BAG"]));
                            }

                          
                        }
                            else
                            {
                                dr["L3_QTY_UNIT_PER_BOX"] = dr["L2_QTY_REEL_PER_BAG"].ToString();
                            }
                        //else
                        //{
                        //    dr["L3_QTY_UNIT_PER_BOX"] = dr["L2_QTY_REEL_PER_BAG"];
                        //}

                        //Loop9
                        if (LevelCOunt == 3)
                        {
                            dt_result.Rows.Add(dr);
                            dt_result.Rows.Add(ds);
                            CountTRA = CountTRA + 1;
                            loop = 1;
                            check = false;
                            checkdtAdd = true;
                            Type = "Another";

                        }
                    }
                        //if (dataRow["PACK_TYPE"].ToString() == "QUADRANT")
                        //{
                        //    switch (dataRow["STOCK_NO"].ToString())
                        //    {
                        //        case "QUAD_1": ds["PIN1_ON_TRAY"] = "Quadrant 1"; break;
                        //        case "QUAD_2": ds["PIN1_ON_TRAY"] = "Quadrant 2"; break;
                        //        case "QUAD_3": ds["PIN1_ON_TRAY"] = "Quadrant 3"; break;
                        //        case "QUAD_4": ds["PIN1_ON_TRAY"] = "Quadrant 4"; break;

                        //    }
                        //    dt_result.Rows.Add(dr);
                        //    dt_result.Rows.Add(ds);
                        //    loop = 1;
                        //    check = false;
                        //    checkdtAdd = true;
                        //}

                    }
                    
                    #endregion
                    #region TNR
                    //ถูก 90% ต้นแบบละเอียด
                    if (Type == "TNR")
                    {
                        if (A == dataRow["PACK_ID"].ToString() && dataRow["PACK_REV"].ToString() == ItemSet && checkdtAdd != true && dataRow["SEQ_NO"].ToString() !="1")
                        {
                            loop = loop + 1; checkpassOnce = true;

                        }

                        A = dataRow["PACK_ID"].ToString();
                        if (loop == 1 && dataRow["PACK_TYPE"].ToString() == "REEL")
                        {
                            checkpassOnce = false;
                            checkdtAdd = false;

                            dr = dt_result.NewRow();
                            ds = dt_result.NewRow();
                            dr["WI_PACK_ID"] = dataRow["PACK_ID"].ToString() + "_" + dataRow["PACK_REV"].ToString() + "_OPTN";
                        ds["WI_PACK_ID"] = dataRow["PACK_ID"].ToString() + "_" + dataRow["PACK_REV"].ToString();

                            dr["WI_TYPE"] = "Generic";
                            dr["DESCRIPTION"] = dataRow["PACK_DESCRIPTION"].ToString();
                        dr["INSTRUC_OPTN"] = "TNR";
                            if (dataRow["HTB"].ToString() == "IMMED")
                            {
                                dr["HTB"] = "P";
                            }
                            else
                            {
                                dr["HTB"] = dataRow["HTB"].ToString();
                            }
                            dr["UNIT_PER_REEL"] = dataRow["UNIT"].ToString();
                            dr["UNIT_PLACEMENT"] = "Live bug";
                            dr["LABEL_POSITION"] = "Sprocket hole";
                            dr["CREATED_BY"] = "System";
                            dr["CREATED_BY_NAME"] = "System";
                            dr["CREATED_DATE"] = mytime;
                            dr["UPDATED_BY"] = "System";
                            dr["UPDATED_BY_NAME"] = "System";
                            dr["UPDATED_DATE"] = mytime;
                            dr["UNIQUE_ID"] = "0";
                            dr["STATUS"] = "1";

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
                            ds["UNIQUE_ID"] = "0";
                            ds["STATUS"] = "1";

                        }
                        if (dataRow["PACK_TYPE"].ToString() == "BAG")
                        {
                            ds["L2_UNIT_PER_BAG"] = dataRow["UNIT"].ToString();
                            ds["L2_QTY_REEL_PER_BAG"] = dataRow["PACK_QTY"].ToString();
                        }
                        if (dataRow["PACK_TYPE"].ToString() == "BOX")
                        {
                            ds["L3_QTY_UNIT_PER_BOX"] = dataRow["UNIT"].ToString();
                            ds["L3_QTY_REEL_PER_BOX"] = dataRow["PACK_QTY"].ToString();
                            //ds["L2_UNIT_PER_BAG"] = dataRow["UNIT"].ToString();
                            //ds["L2_QTY_REEL_PER_BAG"] = dataRow["PACK_QTY"].ToString();
                        }
                        if(dataRow["PACK_TYPE"].ToString() == "LABEL" && dataRow["PACK_QTY"].ToString()!="")
                        {
                            ds["L3_CUST_LABEL_FLAG"] = "Yes";
                            ds["L3_CUST_LABEL_QTY"] = dataRow["PACK_QTY"].ToString();

                        }
                        if (dataRow["PACK_TYPE"].ToString() == "TAPE" && dataRow["SEQ_NO"].ToString() != "4")
                        {
                            ds["L3_QTY_TAPE_LINE"] = dataRow["PACK_QTY"].ToString();

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
                            

                            dt_result.Rows.Add(ds);
                            dt_result.Rows.Add(dr);
                        CountTNR = CountTNR + 1;
                            loop = 1;
                            check = false;
                            checkdtAdd = true;
                        Type = "Another";

                    }
                }
                    #endregion
                    if (Type == "Another")
                    {

                        loop = 1;

                    }
                    #region ERROR

                    if (Type == "Another2")
                    {
                        if (loop == 1 /*&& dataRow["PACK_REV"].ToString() != ItemSet && dataRow["PACK_ID"].ToString() != PassID*/)
                        {
                            dr = dt_result.NewRow();
                        if (dataRow["SEQ_NO"].ToString()!="1")
                        {
                            dr["WI_PACK_ID"] = "ERROR On " + dataRow["PACK_ID"].ToString() + "_" + dataRow["PACK_REV"].ToString() +" Seq="+ dataRow["SEQ_NO"].ToString() + " Method=" + dataRow["METHOD"].ToString() + " TypeT =" + dataRow["PACK_TYPE"].ToString();
                        }
                        else
                        {
                            dr["WI_PACK_ID"] = "ERROR On " + dataRow["PACK_ID"].ToString() + "_" + dataRow["PACK_REV"].ToString() + " Method=" + dataRow["METHOD"].ToString() + " TypeT =" + dataRow["PACK_TYPE"].ToString();
                        }
                        getErrorPoint = getErrorPoint  + dr["WI_PACK_ID"].ToString()+ "\n";
                            errorCount=errorCount + 1;
                            dt_result.Rows.Add(dr);
                            check = false;
                            checkdtAdd = true;
                        Type = "Another";
                            loop = 1;
                        }
                    }
                    #endregion
                    ItemSet = dataRow["PACK_REV"].ToString();
                PassID = dataRow["PACK_ID"].ToString();
                PassSEQ = dataRow["SEQ_NO"].ToString();
                PassDateCre = dataRow["CREATE_DATE"].ToString();
                }
                #region Last Value 2
                //last
                if (checkdtAdd == false && checkpassOnce == true && Type != "Another" && loop != 1)
                {

                    switch (Type)
                    {
                    case "CAN":

                        dt_result.Rows.Add(ds);
                        dt_result.Rows.Add(dr);
                        CountCAN = CountCAN + 1;
                        break;
                    case "TNR":

                        dt_result.Rows.Add(ds);
                        dt_result.Rows.Add(dr); CountTNR = CountTNR + 1
                         ; break;
                    case "TRA":

                        dt_result.Rows.Add(dr);
                        dt_result.Rows.Add(ds); CountTRA = CountTRA + 1; break;

                    case "WAF":


                        dt_result.Rows.Add(dr);
                        dt_result.Rows.Add(ds); CountWAF = CountWAF + 1; break;
                    case "RAI":
                        if (sideTye == "FILM FRAME")
                        {
                            dt_result.Rows.Add(dr);

                            dt_result.Rows.Add(ds);
                            CountWAF = CountWAF + 1;
                        }
                        if (sideTye == "TUBE")
                        {

                            dt_result.Rows.Add(ds);
                            dt_result.Rows.Add(dr);
                            CountRAI = CountRAI + 1;
                        }

                        break;


                }
                    
                }
                #endregion
                //forthe last error
                #endregion
                int resultROW = dt_result.Rows.Count;
                int ItemsSuccess = (resultROW - errorCount) / 2;
                string AllItems = Convert.ToString(ItemsSuccess + errorCount);
            ERRORget = "ALLItem ="+AllItems+"\n ItemsGenSuccess = "+ItemsSuccess.ToString()+"\n CAN ="+CountCAN.ToString()+"\n TNR =" + CountTNR.ToString()+"\n TRA =" + CountTRA.ToString()+"\n WAF =" + CountWAF.ToString() + "\n RAI =" + CountRAI.ToString()  + "\n Error value Count = " + errorCount.ToString() + "\n" + getErrorPoint.ToString();
            dataGridViewOutput.DataSource = dt_result;
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        #region Export

        private void btnExport_Click(object sender, EventArgs e)
        {
            string filename = tbFile.Text;
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    string[] files = Directory.GetFiles(fbd.SelectedPath);
                    try
                    {
                        using (var workbook = new XLWorkbook())

                        {

                            var worksheet = workbook.Worksheets.Add(dt_result, "WIPack");

                            var fullpath = @fbd.SelectedPath + "\\" + filename + ".xlsx";
                            workbook.SaveAs(fullpath);

                            using (FileStream fs = File.Create(@fbd.SelectedPath + "\\" + filename +" ERROR"+".txt"))
                            {
                                // Add some text to file    
                                Byte[] title = new UTF8Encoding(true).GetBytes(ERRORget);
                                fs.Write(title, 0, title.Length);
                                
                            }

                            MessageBox.Show("SAVE to " + fbd.SelectedPath);





                        }

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString());
                    }
                }
            }

        }
        #endregion

        private void MainForm_Load(object sender, EventArgs e)
        {
            //mytime = DateTime.Now.ToString("R");
            //mytime = mytime.Substring(5);
            //string[] timeS = mytime.Split(' ');
            //timeS[1] = timeS[1].ToUpper();
            //timeS[2] = Convert.ToString(Convert.ToInt32(timeS[2]) % 100);
            //var timeList = timeS.ToList();
            //timeList.Remove(timeS[3]);
            //timeList.Remove(timeS[4]);
            //mytime = String.Join("-", timeList);
            //MessageBox.Show(mytime + "   " + DateTime.Now.ToString("dd-MMM-yy").ToUpper());
            mytime = DateTime.Now.ToString("dd-MMM-yy").ToUpper();
            MessageBox.Show("Welcome to WI Pack Migration Data ,\nToday is "+mytime+" \n Have A Good Day.");
        }

        private void tbFile_TextChanged(object sender, EventArgs e)
        {
            if (tbFile.Text != ""||tbFile.Text!=" ")
            {
                btnExport.Enabled = true;
            }
        }
    }
}
