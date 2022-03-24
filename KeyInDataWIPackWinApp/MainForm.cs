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
using System.Reflection;
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
                        MessageBox.Show("Don't forget to check duplicate data before convert");
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
                    dt_result.Columns.Add("L1_HTB");
                    dt_result.Columns.Add("L2_HTB");
                    dt_result.Columns.Add("L3_HTB");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                string Type = "";
                #endregion

                #region Convert
                var PackOut = dt_result.NewRow();
                var OPTN = dt_result.NewRow();
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
                int CountBAG = 0;
                int Count2first = 0;

                foreach (DataRow dataRow in GLOBAL_DataSource.Rows) 
                {
                    countrow = countrow + 1;
                    if (dataRow["SEQ_NO"].ToString() == "1")
                    {
                        countItem = countItem + 1;
                    }
                    checksideType = false;

                    #region Last Value 1
                    if ((A != dataRow["PACK_ID"].ToString() || ItemSet != dataRow["PACK_REV"].ToString()))
                    {
                        if (sideTye != "TUBE" && loop != 1)
                        {
                            checksideType = true;
                        }
                        if (sideTye != "TUBE" && PassSEQ == "1" && dataRow["SEQ_NO"].ToString() == "1")
                        {
                            checksideType = true;
                            checkpassOnce = true;
                        }
                    }
                    if (sideTye == "TUBE" && (B != dataRow["PACK_ID"].ToString() || ItemSet != dataRow["PACK_REV"].ToString()))
                    {
                        if (sideTye == "TUBE" && loop != 1 && B == dataRow["PACK_ID"].ToString() && ItemSet != dataRow["PACK_REV"].ToString())
                        {
                            checksideType = true;
                        }
                        if (sideTye == "TUBE" && B != dataRow["PACK_ID"].ToString() && loop != 1)
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
                            dt_result.Rows.Add(PackOut);
                            dt_result.Rows.Add(OPTN);
                            switch (Type)
                            {
                                case "CAN":
                                    CountCAN = CountCAN + 1;
                                    break;
                                case "TNR":
                                    CountTNR = CountTNR + 1
                                     ; break;
                                case "TRA":
                                   CountTRA = CountTRA + 1; break;
                                case "WAF":
                                    CountWAF = CountWAF + 1; break;
                                case "RAI":
                                    if (sideTye == "FILM FRAME")
                                    {
                                        CountWAF = CountWAF + 1;
                                    }
                                    if (sideTye == "TUBE")
                                    {
                                        CountRAI = CountRAI + 1;
                                    }
                                    break;
                                case "BAG":
                                    CountBAG = CountBAG + 1; break;
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
                        else
                        {
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
                                case "BAG": checkdtAdd = false; check = true; if (dataRow["PACK_TYPE"].ToString() == "BAG") { Type = "BAG"; } else { Type = "Another2"; }; break;
                                case "":
                                    if (/*dataRow["HTB"].ToString() != "" &&*/ (PassID == dataRow["PACK_ID"].ToString() && ItemSet != dataRow["PACK_REV"].ToString()) 
                                        || (PassID != dataRow["PACK_ID"].ToString()) || PassID == dataRow["PACK_ID"].ToString() 
                                        && ItemSet == dataRow["PACK_REV"].ToString() && dataRow["CREATE_DATE"].ToString() != PassDateCre)
                                    {
                                        checkdtAdd = false; check = true;
                                        if (dataRow["PACK_TYPE"].ToString() == "REEL" || dataRow["PACK_TYPE"].ToString() == "TRAY" 
                                            || dataRow["PACK_TYPE"].ToString() == "TUBE" || dataRow["PACK_TYPE"].ToString() == "FILM FRAME" 
                                            || dataRow["PACK_TYPE"].ToString() == "CANISTER" || dataRow["PACK_TYPE"].ToString() == "BAG")
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
                                            if (dataRow["PACK_TYPE"].ToString() == "BAG")
                                            { Type = "BAG"; }
                                        }
                                        else { Type = "Another2"; }
                                    }
                                    else { checkdtAdd = false; if (check == false) { Type = "Another"; } }
                                    break;
                                default: checkdtAdd = false; if (check == false) { Type = "Another"; } break;
                            }
                        }
                    }
                    if (loop == 1 && dataRow["SEQ_NO"].ToString() != "1" && Type == "Another" && check == false && checkdtAdd == true && sideTye == "")
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

                    #region BAG
                    if (Type == "BAG")
                    {
                        if (A == dataRow["PACK_ID"].ToString() && dataRow["PACK_REV"].ToString() == ItemSet && checkdtAdd != true)
                        {
                            loop = loop + 1;
                            checkpassOnce = true;
                        }

                        A = dataRow["PACK_ID"].ToString();
                        if (loop == 1 && dataRow["PACK_TYPE"].ToString() == "BAG")
                        {
                            OPTN = dt_result.NewRow();//OPR
                            PackOut = dt_result.NewRow();//PackOUT
                            checkpassOnce = false;
                            checkdtAdd = false;
                            OPTN["WI_PACK_ID"] = dataRow["PACK_ID"].ToString() + "_" + dataRow["PACK_REV"].ToString() + "_OPTN";
                            PackOut["WI_PACK_ID"] = dataRow["PACK_ID"].ToString() + "_" + dataRow["PACK_REV"].ToString();
                            OPTN["WI_TYPE"] = "Generic";
                            OPTN["DESCRIPTION"] = dataRow["PACK_DESCRIPTION"].ToString();
                            OPTN["INSTRUC_OPTN"] = "Bag";
                            OPTN["HTB"] = dataRow["HTB"].ToString();
                            OPTN["UNIT_PER_BAG"] = dataRow["UNIT"].ToString();
                            OPTN["CREATED_BY"] = "System";
                            OPTN["CREATED_BY_NAME"] = "System";
                            OPTN["CREATED_DATE"] = mytime;
                            OPTN["UPDATED_BY"] = "System";
                            OPTN["UPDATED_BY_NAME"] = "System";
                            OPTN["UPDATED_DATE"] = mytime;
                            OPTN["UNIQUE_ID"] = "0";
                            OPTN["STATUS"] = "1";
                            OPTN["UNIT_PER_BAG"] = dataRow["UNIT"].ToString();
                            //PackOut["L1_HTB"] = dataRow["HTB"].ToString();//?
                            PackOut["WI_TYPE"] = "Generic";
                            PackOut["DESCRIPTION"] = dataRow["PACK_DESCRIPTION"].ToString();
                            PackOut["INSTRUC_OPTN"] = "Pack Out";
                            PackOut["PACKOUT_TYPE"] = "Bag";
                            PackOut["L1_UNIT_PER_BAG"] = dataRow["UNIT"].ToString();
                            PackOut["L1_CUST_LABEL_FLAG"] = "N";
                            PackOut["L1_ESD_FLAG"] = "N";
                            PackOut["L2_QTY_BAG_PER_BAG"] = dataRow["PACK_QTY"].ToString();
                            PackOut["L2_DRY_PACK_FLAG"] = "N";
                            PackOut["L2_CACUUM_SEAL_FLAG"] = "N";
                            PackOut["L2_CUST_LABEL_FLAG"] = "N";
                            PackOut["L2_ESD_FLAG"] = "N";
                            PackOut["L2_CAUTION_FLAG"] = "N";
                            PackOut["L2_HIC_FLAG"] = "No";
                            PackOut["L2_DESICCANT_FLAG"] = "N";
                            PackOut["L3_CUST_LABEL_FLAG"] = "N";
                            PackOut["L3_ESD_FLAG"] = "N";
                            PackOut["L3_BUBBLE_FLAG"] = "N";
                            PackOut["L3_CAUTION_FLAG"] = "N";
                            PackOut["L1_HTB"] = dataRow["HTB"].ToString();
                            PackOut["CREATED_BY"] = "System";
                            PackOut["CREATED_BY_NAME"] = "System";
                            PackOut["CREATED_DATE"] = mytime;
                            PackOut["UPDATED_BY"] = "System";
                            PackOut["UPDATED_BY_NAME"] = "System";
                            PackOut["UPDATED_DATE"] = mytime;
                            PackOut["UNIQUE_ID"] = "0";
                            PackOut["STATUS"] = "1";
                        }
                        if (dataRow["PACK_QTY"].ToString() != "" && dataRow["CAUTION"].ToString() == "Y")
                        {
                            switch (dataRow["PACK_LEVEL"].ToString())
                            {
                                case "IN1":
                                    PackOut["L2_CAUTION_FLAG"] = "Y"; PackOut["L2_CAUTION_QTY"] = dataRow["PACK_QTY"].ToString();
                                    break;
                                case "IN2":
                                    PackOut["L3_CAUTION_FLAG"] = "Y"; PackOut["L3_CAUTION_QTY"] = dataRow["PACK_QTY"].ToString();
                                    break;
                            }
                        }
                        if (dataRow["PACK_QTY"].ToString() == "" && dataRow["CAUTION"].ToString() == "Y")
                        {
                            switch (dataRow["PACK_LEVEL"].ToString())
                            {
                                case "IN1":
                                    PackOut["L2_CAUTION_FLAG"] = "Y"; PackOut["L2_CAUTION_QTY"] = "0";
                                    break;
                                case "IN2":
                                    PackOut["L3_CAUTION_FLAG"] = "Y"; PackOut["L3_CAUTION_QTY"] = "0";
                                    break;
                            }
                        }
                        if (dataRow["PACK_TYPE"].ToString() == "BOX")
                        {
                            PackOut["L3_QTY_BAG_PER_BOX"] = dataRow["PACK_QTY"].ToString();
                            PackOut["L3_QTY_UNIT_PER_BOX"] = dataRow["UNIT"].ToString();
                            //if(checkdtAdd != true) { 
                            dt_result.Rows.Add(PackOut);
                            dt_result.Rows.Add(OPTN);
                            CountBAG = CountBAG + 1;
                            loop = 1;
                            check = false;
                            checkdtAdd = true;
                            Type = "Another";
                            //}
                        }
                    }///อ่านต้องเปลี่ยนเป็นเช็คค่าตาม Column
                    #endregion

                    #region CAN
                    //ถูก 88% 
                    if (Type == "CAN")
                    {
                        if (A == dataRow["PACK_ID"].ToString() && dataRow["PACK_REV"].ToString() == ItemSet && checkdtAdd != true)
                        {
                            loop = loop + 1;
                            checkpassOnce = true;
                        }
                        A = dataRow["PACK_ID"].ToString();
                        if (loop == 1 && dataRow["PACK_TYPE"].ToString() == "CANISTER")
                        {
                            OPTN = dt_result.NewRow();//OPR
                            PackOut = dt_result.NewRow();//PackOUT
                            checkpassOnce = false;
                            checkdtAdd = false;
                            OPTN["WI_PACK_ID"] = dataRow["PACK_ID"].ToString() + "_" + dataRow["PACK_REV"].ToString() + "_OPTN";
                            PackOut["WI_PACK_ID"] = dataRow["PACK_ID"].ToString() + "_" + dataRow["PACK_REV"].ToString();
                            OPTN["WI_TYPE"] = "Generic";
                            OPTN["DESCRIPTION"] = dataRow["PACK_DESCRIPTION"].ToString();
                            OPTN["INSTRUC_OPTN"] = "Canister";
                            //if (dataRow["HTB"].ToString().Contains(">"))
                            //{
                            //    OPTN["HTB"] = dataRow["HTB"].ToString().Substring(dataRow["HTB"].ToString().IndexOf(">") + 1);
                            //    OPTN["HTB"] = "P";
                            //}
                            //else
                            //{
                            //    OPTN["HTB"] = dataRow["HTB"].ToString();
                            //}
                            OPTN["HTB"] = dataRow["HTB"].ToString();
                            OPTN["UNIT_PER_CAN"] = dataRow["UNIT"].ToString();
                            OPTN["CREATED_BY"] = "System";
                            OPTN["CREATED_BY_NAME"] = "System";
                            OPTN["CREATED_DATE"] = mytime;
                            OPTN["UPDATED_BY"] = "System";
                            OPTN["UPDATED_BY_NAME"] = "System";
                            OPTN["UPDATED_DATE"] = mytime;
                            OPTN["UNIQUE_ID"] = "0";
                            OPTN["STATUS"] = "1";
                            PackOut["L1_HTB"] = dataRow["HTB"].ToString();
                            PackOut["WI_TYPE"] = "Generic";
                            PackOut["DESCRIPTION"] = dataRow["PACK_DESCRIPTION"].ToString();
                            PackOut["INSTRUC_OPTN"] = "Pack Out";
                            PackOut["PACKOUT_TYPE"] = "Canister";
                            PackOut["L1_UNIT_PER_CAN"] = dataRow["UNIT"].ToString();
                            PackOut["L1_CUST_LABEL_FLAG"] = "No";
                            PackOut["L2_DRY_PACK_FLAG"] = "No";
                            PackOut["L2_CACUUM_SEAL_FLAG"] = "No";
                            PackOut["L2_CUST_LABEL_FLAG"] = "No";
                            PackOut["L2_ESD_FLAG"] = "N";
                            PackOut["L2_CAUTION_FLAG"] = "N";
                            PackOut["L2_HIC_FLAG"] = "No";
                            PackOut["L2_DESICCANT_FLAG"] = "No";
                            PackOut["L3_CUST_LABEL_FLAG"] = "No";
                            PackOut["L3_ESD_FLAG"] = "N";
                            PackOut["L3_BUBBLE_FLAG"] = "No";
                            PackOut["L3_CAUTION_FLAG"] = "N";
                            PackOut["L1_HTB"] = dataRow["HTB"].ToString();
                            PackOut["CREATED_BY"] = "System";
                            PackOut["CREATED_BY_NAME"] = "System";
                            PackOut["CREATED_DATE"] = mytime;
                            PackOut["UPDATED_BY"] = "System";
                            PackOut["UPDATED_BY_NAME"] = "System";
                            PackOut["UPDATED_DATE"] = mytime;
                            PackOut["UNIQUE_ID"] = "0";
                            PackOut["STATUS"] = "1";
                        }
                        if (dataRow["PACK_QTY"].ToString() != "" && dataRow["CAUTION"].ToString() == "Y")
                        {
                            switch (dataRow["PACK_LEVEL"].ToString())
                            {
                                case "IN1":
                                    PackOut["L2_CAUTION_FLAG"] = "Y"; PackOut["L2_CAUTION_QTY"] = dataRow["PACK_QTY"].ToString();
                                    break;
                                case "IN2":
                                    PackOut["L3_CAUTION_FLAG"] = "Y"; PackOut["L3_CAUTION_QTY"] = dataRow["PACK_QTY"].ToString();
                                    break;
                            }
                        }
                        if (dataRow["PACK_QTY"].ToString() == "" && dataRow["CAUTION"].ToString() == "Y")
                        {
                            switch (dataRow["PACK_LEVEL"].ToString())
                            {
                                case "IN1":
                                    PackOut["L2_CAUTION_FLAG"] = "Y"; PackOut["L2_CAUTION_QTY"] = "0";
                                    break;
                                case "IN2":
                                    PackOut["L3_CAUTION_FLAG"] = "Y"; PackOut["L3_CAUTION_QTY"] = "0";
                                    break;
                            }
                        }
                        if (dataRow["PACK_TYPE"].ToString() == "BOX")
                        {
                            PackOut["L3_QTY_UNIT_PER_BOX"] = dataRow["UNIT"].ToString();
                            //if(checkdtAdd != true) { 
                            dt_result.Rows.Add(PackOut);
                            dt_result.Rows.Add(OPTN);
                            CountCAN = CountCAN + 1;
                            loop = 1;
                            check = false;
                            checkdtAdd = true;
                            Type = "Another";
                            //}
                        }
                    }///อ่านต้องเปลี่ยนเป็นเช็คค่าตาม Column
                    #endregion

                    #region RAI
                    //ถูก 88%
                    if (Type == "RAI")
                    {
                        if (dataRow["PACK_TYPE"].ToString() == "FILM FRAME" || sideTye == "FILM FRAME")
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
                                PackOut = dt_result.NewRow();
                                OPTN = dt_result.NewRow();
                                PackOut["WI_PACK_ID"] = dataRow["PACK_ID"].ToString() + "_" + dataRow["PACK_REV"].ToString();
                                OPTN["WI_PACK_ID"] = dataRow["PACK_ID"].ToString() + "_" + dataRow["PACK_REV"].ToString() + "_OPTN";
                                PackOut["WI_TYPE"] = "Generic";
                                PackOut["DESCRIPTION"] = dataRow["PACK_DESCRIPTION"].ToString();
                                PackOut["INSTRUC_OPTN"] = "Pack Out";
                                PackOut["PACKOUT_TYPE"] = "Wafer";
                                PackOut["L1_UNIT_PER_WF_BOX"] = dataRow["UNIT"].ToString();
                                PackOut["L2_QTY_WF_BOX_PER_BAG"] = dataRow["PACK_QTY"].ToString();
                                PackOut["L1_CUST_LABEL_FLAG"] = "No";
                                PackOut["L2_DRY_PACK_FLAG"] = "No";
                                PackOut["L2_CACUUM_SEAL_FLAG"] = "No";
                                PackOut["L2_CUST_LABEL_FLAG"] = "No";
                                PackOut["L2_ESD_FLAG"] = "N";
                                PackOut["L2_CAUTION_FLAG"] = "N";
                                PackOut["L2_HIC_FLAG"] = "No";
                                PackOut["L2_DESICCANT_FLAG"] = "No";
                                PackOut["L3_CUST_LABEL_FLAG"] = "No";
                                PackOut["L3_ESD_FLAG"] = "N";
                                PackOut["L3_BUBBLE_FLAG"] = "No";
                                PackOut["L3_CAUTION_FLAG"] = "N";
                                PackOut["CREATED_BY"] = "System";
                                PackOut["CREATED_BY_NAME"] = "System";
                                PackOut["CREATED_DATE"] = mytime;
                                PackOut["UPDATED_BY"] = "System";
                                PackOut["UPDATED_BY_NAME"] = "System";
                                PackOut["UPDATED_DATE"] = mytime;
                                PackOut["UNIQUE_ID"] = "0";
                                PackOut["STATUS"] = "1";
                                OPTN["WI_TYPE"] = "Generic";
                                OPTN["DESCRIPTION"] = dataRow["PACK_DESCRIPTION"].ToString();
                                OPTN["INSTRUC_OPTN"] = "Wafer";
                                OPTN["UNIT_PER_WAFER_BOX"] = dataRow["UNIT"].ToString(); 
                                OPTN["HTB"] = dataRow["HTB"].ToString();
                                OPTN["CREATED_BY"] = "System";
                                OPTN["CREATED_BY_NAME"] = "System";
                                OPTN["CREATED_DATE"] = mytime;
                                OPTN["UPDATED_BY"] = "System";
                                OPTN["UPDATED_BY_NAME"] = "System";
                                OPTN["UPDATED_DATE"] = mytime;
                                OPTN["UNIQUE_ID"] = "0";
                                OPTN["STATUS"] = "1";
                                PackOut["L1_HTB"] = dataRow["HTB"].ToString();
                            }
                            if (dataRow["PACK_QTY"].ToString() != "" && dataRow["CAUTION"].ToString() == "Y")
                            {
                                switch (dataRow["PACK_LEVEL"].ToString())
                                {
                                    case "IN1":
                                        PackOut["L2_CAUTION_FLAG"] = "Y"; PackOut["L2_CAUTION_QTY"] = dataRow["PACK_QTY"].ToString();
                                        break;
                                    case "IN2":
                                        PackOut["L3_CAUTION_FLAG"] = "Y"; PackOut["L3_CAUTION_QTY"] = dataRow["PACK_QTY"].ToString();
                                        break;
                                }
                            }
                            if (dataRow["PACK_QTY"].ToString() == "" && dataRow["CAUTION"].ToString() == "Y")
                            {
                                switch (dataRow["PACK_LEVEL"].ToString())
                                {
                                    case "IN1":
                                        PackOut["L2_CAUTION_FLAG"] = "Y"; PackOut["L2_CAUTION_QTY"] = "0";
                                        break;
                                    case "IN2":
                                        PackOut["L3_CAUTION_FLAG"] = "Y"; PackOut["L3_CAUTION_QTY"] = "0";
                                        break;
                                }
                            }
                            if (dataRow["PACK_TYPE"].ToString() == "BOX")
                            {
                                PackOut["L2_QTY_WF_BOX_PER_BAG"] = dataRow["PACK_QTY"].ToString();
                                PackOut["L3_QTY_WF_BOX_PER_BOX"] = dataRow["PACK_QTY"].ToString();
                                PackOut["L3_QTY_UNIT_PER_BOX"] = dataRow["UNIT"].ToString();
                                dt_result.Rows.Add(PackOut);
                                dt_result.Rows.Add(OPTN);
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
                            string statoflevel = "";
                            if (B == dataRow["PACK_ID"].ToString() && dataRow["PACK_REV"].ToString() == ItemSet && checkdtAdd != true && sideTye != "")
                            {
                                loop = loop + 1;
                                checkpassOnce = true;
                            }
                            B = dataRow["PACK_ID"].ToString();
                            if (loop == 1)
                            {
                                sideTye = "TUBE"; checkpassOnce = false;
                                checkdtAdd = false;
                                OPTN = dt_result.NewRow();//OPR
                                PackOut = dt_result.NewRow();//PackOUT
                                OPTN["WI_PACK_ID"] = dataRow["PACK_ID"].ToString() + "_" + dataRow["PACK_REV"].ToString() + "_OPTN";
                                PackOut["WI_PACK_ID"] = dataRow["PACK_ID"].ToString() + "_" + dataRow["PACK_REV"].ToString();
                                OPTN["WI_TYPE"] = "Generic";
                                OPTN["DESCRIPTION"] = dataRow["PACK_DESCRIPTION"].ToString();
                                OPTN["INSTRUC_OPTN"] = "Tube";
                                OPTN["HTB"] = dataRow["HTB"].ToString();
                                OPTN["UNIT_PER_TUBE"] = dataRow["UNIT"].ToString();
                                OPTN["P1_FULL_TUBE_FOAM"] = "No";
                                OPTN["OP_P1_FULL_TUBE_FOAM"] = "No";
                                OPTN["P1_COMBINE_TUBE_FOAM"] = "No";
                                OPTN["OP_P1_COMBINE_TUBE_FOAM"] = "No";
                                OPTN["P1_PARTIAL_TUBE_FOAM"] = "No";
                                OPTN["OP_P1_PARTIAL_TUBE_FOAM"] = "No";
                                OPTN["CREATED_BY"] = "System";
                                OPTN["CREATED_BY_NAME"] = "System";
                                OPTN["CREATED_DATE"] = mytime;
                                OPTN["UPDATED_BY"] = "System";
                                OPTN["UPDATED_BY_NAME"] = "System";
                                OPTN["UPDATED_DATE"] = mytime;
                                OPTN["UNIQUE_ID"] = "0";
                                OPTN["STATUS"] = "1";
                                PackOut["WI_TYPE"] = "Generic";
                                PackOut["DESCRIPTION"] = dataRow["PACK_DESCRIPTION"].ToString();
                                PackOut["INSTRUC_OPTN"] = "Pack Out";
                                PackOut["PACKOUT_TYPE"] = "Tube";
                                PackOut["L1_UNIT_PER_TUBE"] = dataRow["UNIT"].ToString();
                                PackOut["L1_PINK_FOAM_FLAG"] = "No";
                                PackOut["L1_WRAP_RUBBER_FLAG"] = "No";
                                PackOut["L1_WRAP_BUBBLE_FLAG"] = "No";
                                PackOut["L2_DRY_PACK_FLAG"] = "No";
                                PackOut["L2_CACUUM_SEAL_FLAG"] = "No";
                                PackOut["L2_CUST_LABEL_FLAG"] = "No";
                                PackOut["L2_ESD_FLAG"] = "N";
                                PackOut["L2_CAUTION_FLAG"] = "N";
                                PackOut["L2_HIC_FLAG"] = "No";
                                PackOut["L2_DESICCANT_FLAG"] = "No";
                                PackOut["L3_CUST_LABEL_FLAG"] = "No";
                                PackOut["L3_ESD_FLAG"] = "N";
                                PackOut["L3_BUBBLE_FLAG"] = "No";
                                PackOut["L3_CAUTION_FLAG"] = "N";
                                PackOut["CREATED_BY"] = "System";
                                PackOut["CREATED_BY_NAME"] = "System";
                                PackOut["CREATED_DATE"] = mytime;
                                PackOut["UPDATED_BY"] = "System";
                                PackOut["UPDATED_BY_NAME"] = "System";
                                PackOut["UPDATED_DATE"] = mytime;
                                PackOut["UNIQUE_ID"] = "0";
                                PackOut["STATUS"] = "1";
                                PackOut["L1_HTB"] = dataRow["HTB"].ToString();
                            }
                            if (dataRow["PACK_LEVEL"].ToString() != "")
                            {
                                statoflevel = dataRow["PACK_LEVEL"].ToString();
                            }
                            if (dataRow["PACK_TYPE"].ToString() == "BAG")
                            {
                                PackOut["L2_QTY_TUBE_PER_BAG"] = dataRow["PACK_QTY"].ToString();
                            }
                            if (dataRow["PACK_QTY"].ToString() != "" && dataRow["CAUTION"].ToString() == "Y")
                            {
                                switch (dataRow["PACK_LEVEL"].ToString())
                                {
                                    case "IN1":
                                        PackOut["L2_CAUTION_FLAG"] = "Y"; PackOut["L2_CAUTION_QTY"] = dataRow["PACK_QTY"].ToString();
                                        break;
                                    case "IN2":
                                        PackOut["L3_CAUTION_FLAG"] = "Y"; PackOut["L3_CAUTION_QTY"] = dataRow["PACK_QTY"].ToString();
                                        break;
                                }
                            }
                            if (dataRow["PACK_QTY"].ToString() == "" && dataRow["CAUTION"].ToString() == "Y")
                            {
                                switch (dataRow["PACK_LEVEL"].ToString())
                                {
                                    case "IN1":
                                        PackOut["L2_CAUTION_FLAG"] = "Y"; PackOut["L2_CAUTION_QTY"] = "0";
                                        break;
                                    case "IN2":
                                        PackOut["L3_CAUTION_FLAG"] = "Y"; PackOut["L3_CAUTION_QTY"] = "0";
                                        break;
                                }
                            }
                            if (dataRow["PACK_TYPE"].ToString() == "BOX")
                            {
                                PackOut["L3_QTY_BAG_PER_BOX"] = dataRow["PACK_QTY"].ToString();
                                if (PackOut["L3_QTY_BAG_PER_BOX"] != null && PackOut["L2_QTY_TUBE_PER_BAG"] != null 
                                    && PackOut["L3_QTY_BAG_PER_BOX"].ToString() != "" && PackOut["L2_QTY_TUBE_PER_BAG"].ToString() != "")
                                {
                                    PackOut["L3_QTY_TUBE_PER_BOX"] = Convert.ToString(Convert.ToInt32(PackOut["L3_QTY_BAG_PER_BOX"]) * Convert.ToInt32(PackOut["L2_QTY_TUBE_PER_BAG"]));
                                }
                                if (statoflevel == "IN1")
                                {
                                    PackOut["L2_QTY_TUBE_PER_BAG"] = "1";
                                    statoflevel = "";
                                }
                                dt_result.Rows.Add(PackOut);
                                dt_result.Rows.Add(OPTN);
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
                            PackOut = dt_result.NewRow();
                            OPTN = dt_result.NewRow();//OPT
                            checkpassOnce = false;
                            checkdtAdd = false;
                            OPTN["WI_PACK_ID"] = dataRow["PACK_ID"].ToString() + "_" + dataRow["PACK_REV"].ToString() + "_OPTN";
                            PackOut["WI_PACK_ID"] = dataRow["PACK_ID"].ToString() + "_" + dataRow["PACK_REV"].ToString();
                            PackOut["WI_TYPE"] = "Generic";
                            PackOut["DESCRIPTION"] = dataRow["PACK_DESCRIPTION"].ToString();
                            PackOut["INSTRUC_OPTN"] = "Pack Out";
                            PackOut["PACKOUT_TYPE"] = "Wafer";
                            PackOut["L1_UNIT_PER_WF_BOX"] = dataRow["UNIT"].ToString();//unit per wafer box
                            PackOut["L1_CUST_LABEL_FLAG"] = "No";
                            PackOut["L2_DRY_PACK_FLAG"] = "No";
                            PackOut["L2_CACUUM_SEAL_FLAG"] = "No";
                            PackOut["L2_CUST_LABEL_FLAG"] = "No";
                            PackOut["L2_ESD_FLAG"] = "N";
                            PackOut["L2_CAUTION_FLAG"] = "N";
                            PackOut["L2_HIC_FLAG"] = "No";
                            PackOut["L2_DESICCANT_FLAG"] = "No";
                            PackOut["L3_CUST_LABEL_FLAG"] = "No";
                            PackOut["L3_ESD_FLAG"] = "N";
                            PackOut["L3_BUBBLE_FLAG"] = "No";
                            PackOut["L3_CAUTION_FLAG"] = "N";
                            OPTN["WI_TYPE"] = "Generic";
                            OPTN["DESCRIPTION"] = dataRow["PACK_DESCRIPTION"].ToString();
                            OPTN["INSTRUC_OPTN"] = "Wafer";
                            OPTN["UNIT_PER_WAFER_BOX"] = dataRow["UNIT"].ToString(); //unit per wafer box
                                                                                     //if (dataRow["HTB"].ToString() == "IMMED"){
                                                                                     //    ds["HTB"] = "P";
                                                                                     //}
                                                                                     //else
                                                                                     //{
                                                                                     //    if (dataRow["HTB"].ToString() == "IMMED")
                                                                                     //    {
                                                                                     //        ds["HTB"] = "P";
                                                                                     //    }
                                                                                     //    else
                                                                                     //    {
                                                                                     //        ds["HTB"] = dataRow["HTB"].ToString();
                                                                                     //    }
                                                                                     //}

                            OPTN["HTB"] = dataRow["HTB"].ToString();
                            PackOut["CREATED_BY"] = "System";
                            PackOut["CREATED_BY_NAME"] = "System";
                            PackOut["CREATED_DATE"] = mytime;
                            PackOut["UPDATED_BY"] = "System";
                            PackOut["UPDATED_BY_NAME"] = "System";
                            PackOut["UPDATED_DATE"] = mytime;
                            PackOut["UNIQUE_ID"] = "0";
                            PackOut["STATUS"] = "1";
                            OPTN["CREATED_BY"] = "System";
                            OPTN["CREATED_BY_NAME"] = "System";
                            OPTN["CREATED_DATE"] = mytime;
                            OPTN["UPDATED_BY"] = "System";
                            OPTN["UPDATED_BY_NAME"] = "System";
                            OPTN["UPDATED_DATE"] = mytime;
                            OPTN["UNIQUE_ID"] = "0";
                            OPTN["STATUS"] = "1";
                            PackOut["L1_HTB"] = dataRow["HTB"].ToString();
                        }
                        if (dataRow["PACK_TYPE"].ToString() == "BAG")
                        {
                            if (dataRow["UNIT"].ToString() != "")
                            {
                                PackOut["L2_QTY_WF_BOX_PER_BAG"] = Convert.ToString(Convert.ToInt32(dataRow["UNIT"].ToString()) / Convert.ToInt32(PackOut["L1_UNIT_PER_WF_BOX"]));
                            }
                            //PackOut["L2_QTY_WF_BOX_PER_BAG"] = dataRow["PACK_QTY"].ToString();
                        }
                        if (dataRow["PACK_QTY"].ToString() != "" && dataRow["CAUTION"].ToString() == "Y")
                        {
                            switch (dataRow["PACK_LEVEL"].ToString())
                            {
                                case "IN1":
                                    PackOut["L2_CAUTION_FLAG"] = "Y"; PackOut["L2_CAUTION_QTY"] = dataRow["PACK_QTY"].ToString();
                                    break;
                                case "IN2":
                                    PackOut["L3_CAUTION_FLAG"] = "Y"; PackOut["L3_CAUTION_QTY"] = dataRow["PACK_QTY"].ToString();
                                    break;
                            }
                        }
                        if (dataRow["PACK_QTY"].ToString() == "" && dataRow["CAUTION"].ToString() == "Y")
                        {
                            switch (dataRow["PACK_LEVEL"].ToString())
                            {
                                case "IN1":
                                    PackOut["L2_CAUTION_FLAG"] = "Y"; PackOut["L2_CAUTION_QTY"] = "0";
                                    break;
                                case "IN2":
                                    PackOut["L3_CAUTION_FLAG"] = "Y"; PackOut["L3_CAUTION_QTY"] = "0";
                                    break;
                            }
                        }
                        if (dataRow["PACK_TYPE"].ToString() == "BOX")
                        {
                            if (dataRow["UNIT"].ToString() != "")
                            {
                                PackOut["L3_QTY_UNIT_PER_BOX"] = dataRow["UNIT"].ToString();//unit per box
                                PackOut["L3_QTY_WF_BOX_PER_BOX"] = Convert.ToString(Convert.ToInt32(PackOut["L3_QTY_UNIT_PER_BOX"]) / Convert.ToInt32(PackOut["L1_UNIT_PER_WF_BOX"]));
                            }
                            //PackOut["L3_QTY_WF_BOX_PER_BOX"] = dataRow["PACK_QTY"].ToString();//qty
                            //if (PackOut["L2_QTY_WF_BOX_PER_BAG"].ToString() != "")
                            //{
                            //    PackOut["L3_QTY_BAG_PER_BOX"] = Convert.ToString(Convert.ToInt32(PackOut["L3_QTY_WF_BOX_PER_BOX"]) / Convert.ToInt32(PackOut["L2_QTY_WF_BOX_PER_BAG"]));
                            //}
                            //PackOut["WI_PACK_ID"] = "ERROR" + " Method=" + dataRow["METHOD"].ToString() + " PackType =" + dataRow["PACK_TYPE"].ToString(); For Except Wafer
                            dt_result.Rows.Add(PackOut);
                            dt_result.Rows.Add(OPTN);//Operation
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
                            PackOut = dt_result.NewRow();
                            OPTN = dt_result.NewRow();
                            checkpassOnce = false;
                            checkdtAdd = false;
                            PackOut["WI_PACK_ID"] = dataRow["PACK_ID"].ToString() + "_" + dataRow["PACK_REV"].ToString();
                            OPTN["WI_PACK_ID"] = dataRow["PACK_ID"].ToString() + "_" + dataRow["PACK_REV"].ToString() + "_OPTN";
                            PackOut["WI_TYPE"] = "Generic";
                            PackOut["DESCRIPTION"] = dataRow["PACK_DESCRIPTION"].ToString();
                            PackOut["INSTRUC_OPTN"] = "Pack Out";
                            PackOut["L1_UNIT_PER_TRAY"] = dataRow["UNIT"].ToString();
                            PackOut["L1_QTY_TACK_TRAY_FLAG"] = "No";
                            PackOut["L2_DRY_PACK_FLAG"] = "No";
                            PackOut["L2_CACUUM_SEAL_FLAG"] = "No";
                            PackOut["L2_CUST_LABEL_FLAG"] = "No";
                            PackOut["L2_ESD_FLAG"] = "N";
                            PackOut["L2_CAUTION_FLAG"] = "N";
                            PackOut["L2_HIC_FLAG"] = "No";
                            PackOut["L3_CUST_LABEL_FLAG"] = "No";
                            PackOut["L3_ESD_FLAG"] = "N";
                            PackOut["L3_BUBBLE_FLAG"] = "No";
                            PackOut["L3_CAUTION_FLAG"] = "N";
                            PackOut["L2_DESICCANT_FLAG"] = "No";
                            PackOut["CREATED_BY"] = "System";
                            PackOut["CREATED_BY_NAME"] = "System";
                            PackOut["CREATED_DATE"] = mytime;
                            PackOut["UPDATED_BY"] = "System";
                            PackOut["UPDATED_BY_NAME"] = "System";
                            PackOut["UPDATED_DATE"] = mytime;
                            PackOut["UNIQUE_ID"] = "0";
                            PackOut["STATUS"] = "1";
                            PackOut["PACKOUT_TYPE"] = "Tray";
                            OPTN["WI_TYPE"] = "Generic";
                            OPTN["DESCRIPTION"] = dataRow["PACK_DESCRIPTION"].ToString();
                            OPTN["INSTRUC_OPTN"] = "Tray";
                            //if (dataRow["HTB"].ToString() == "IMMED")
                            //{
                            //    ds["HTB"] = "P";
                            //}
                            //else
                            //{
                            //    ds["HTB"] = dataRow["HTB"].ToString();
                            //}
                            OPTN["HTB"] = dataRow["HTB"].ToString();
                            OPTN["UNIT_PER_TRAY"] = dataRow["UNIT"].ToString();
                            //??Value
                            OPTN["QTY_COVER_TOP_SIDE"] = "";
                            OPTN["QTY_COVER_BOTTOM_SIDE"] = "";
                            OPTN["QTY_STACK_TRAY"] = "";
                            OPTN["PIN1_ON_TRAY_IMG_PATH"] = "";
                            OPTN["PARTIAL_TRAY_DIRECTION"] = "";
                            OPTN["PARTIAL_TRAY_IMG_PATH"] = "";
                            OPTN["PIN1_ON_TRAY"] = "";
                            //
                            OPTN["CREATED_BY"] = "System";
                            OPTN["CREATED_BY_NAME"] = "System";
                            OPTN["CREATED_DATE"] = mytime;
                            OPTN["UPDATED_BY"] = "System";
                            OPTN["UPDATED_BY_NAME"] = "System";
                            OPTN["UPDATED_DATE"] = mytime;
                            OPTN["UNIQUE_ID"] = "0";
                            OPTN["STATUS"] = "1";
                            PackOut["L1_HTB"] = dataRow["HTB"].ToString();
                            //ds["PACKOUT_TYPE"] = "Tray";
                            LevelCOunt = LevelCOunt + 1;
                        }
                        if (dataRow["PACK_QTY"].ToString() != "" && dataRow["CAUTION"].ToString() == "Y")
                        {
                            switch (dataRow["PACK_LEVEL"].ToString())
                            {
                                case "IN1":
                                    PackOut["L2_CAUTION_FLAG"] = "Y"; PackOut["L2_CAUTION_QTY"] = dataRow["PACK_QTY"].ToString();
                                    break;
                                case "IN2":
                                    PackOut["L3_CAUTION_FLAG"] = "Y"; PackOut["L3_CAUTION_QTY"] = dataRow["PACK_QTY"].ToString();
                                    break;
                            }
                        }
                        if (dataRow["PACK_QTY"].ToString() == "" && dataRow["CAUTION"].ToString() == "Y")
                        {
                            switch (dataRow["PACK_LEVEL"].ToString())
                            {
                                case "IN1":
                                    PackOut["L2_CAUTION_FLAG"] = "Y"; PackOut["L2_CAUTION_QTY"] = "0";
                                    break;
                                case "IN2":
                                    PackOut["L3_CAUTION_FLAG"] = "Y"; PackOut["L3_CAUTION_QTY"] = "0";
                                    break;
                            }
                        }
                        if (dataRow["PACK_TYPE"].ToString() == "BAG")
                        {
                            LevelCOunt = LevelCOunt + 1;
                            PackOut["L2_QTY_REEL_PER_BAG"] = dataRow["PACK_QTY"].ToString();//this is Tray per Bag
                            if (LevelCOunt == 3)
                            {
                                if (PackOut["L3_QTY_BAG_PER_BOX"].ToString() != "" && PackOut["L2_QTY_REEL_PER_BAG"].ToString() != "")
                                {
                                    PackOut["L3_QTY_UNIT_PER_BOX"] = Convert.ToString(Convert.ToInt32(PackOut["L3_QTY_BAG_PER_BOX"]) * Convert.ToInt32(PackOut["L2_QTY_REEL_PER_BAG"]));
                                }
                                else
                                {
                                    PackOut["L3_QTY_UNIT_PER_BOX"] = PackOut["L2_QTY_REEL_PER_BAG"].ToString();
                                }
                                dt_result.Rows.Add(PackOut);
                                dt_result.Rows.Add(OPTN);
                                CountTRA = CountTRA + 1;
                                loop = 1;
                                check = false;
                                checkdtAdd = true;
                                Type = "Another";
                            }
                            //PackOut["L3_QTY_UNIT_PER_BOX"] = dataRow["PACK_QTY"].ToString();
                        }
                        if (dataRow["PACK_TYPE"].ToString() == "BOX")
                        {
                            LevelCOunt = LevelCOunt + 1;
                            PackOut["L3_QTY_BAG_PER_BOX"] = dataRow["PACK_QTY"].ToString();
                            if (dataRow["UNIT"].ToString() != "")
                            {
                                if (PackOut["L3_QTY_BAG_PER_BOX"].ToString() != "" && PackOut["L2_QTY_REEL_PER_BAG"].ToString() != "")
                                {
                                    PackOut["L3_QTY_UNIT_PER_BOX"] = Convert.ToString(Convert.ToInt32(PackOut["L3_QTY_BAG_PER_BOX"]) * Convert.ToInt32(PackOut["L2_QTY_REEL_PER_BAG"]));
                                }
                            }
                            else
                            {
                                PackOut["L3_QTY_UNIT_PER_BOX"] = PackOut["L2_QTY_REEL_PER_BAG"].ToString();
                            }
                            //else
                            //{
                            //    PackOut["L3_QTY_UNIT_PER_BOX"] = PackOut["L2_QTY_REEL_PER_BAG"];
                            //}
                            //Loop9
                            if (dataRow["PACK_QTY"].ToString() != "" && dataRow["CAUTION"].ToString() == "Y")
                            {
                                switch (dataRow["PACK_LEVEL"].ToString())
                                {
                                    case "IN1":
                                        PackOut["L2_CAUTION_FLAG"] = "Y"; PackOut["L2_CAUTION_QTY"] = dataRow["PACK_QTY"].ToString();
                                        break;
                                    case "IN2":
                                        PackOut["L3_CAUTION_FLAG"] = "Y"; PackOut["L3_CAUTION_QTY"] = dataRow["PACK_QTY"].ToString();
                                        break;
                                }
                            }
                            if (dataRow["PACK_QTY"].ToString() == "" && dataRow["CAUTION"].ToString() == "Y")
                            {
                                switch (dataRow["PACK_LEVEL"].ToString())
                                {
                                    case "IN1":
                                        PackOut["L2_CAUTION_FLAG"] = "Y"; PackOut["L2_CAUTION_QTY"] = "0";
                                        break;
                                    case "IN2":
                                        PackOut["L3_CAUTION_FLAG"] = "Y"; PackOut["L3_CAUTION_QTY"] = "0";
                                        break;
                                }
                            }
                            if (LevelCOunt == 3)
                            {
                                dt_result.Rows.Add(PackOut);
                                dt_result.Rows.Add(OPTN);
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
                        //        case "QUAD_1": OPTN["PIN1_ON_TRAY"] = "Quadrant 1"; break;
                        //        case "QUAD_2": OPTN["PIN1_ON_TRAY"] = "Quadrant 2"; break;
                        //        case "QUAD_3": OPTN["PIN1_ON_TRAY"] = "Quadrant 3"; break;
                        //        case "QUAD_4": OPTN["PIN1_ON_TRAY"] = "Quadrant 4"; break;

                        //    }
                        //    dt_result.Rows.Add(PackOut);
                        //    dt_result.Rows.Add(OPTN);
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
                        if (A == dataRow["PACK_ID"].ToString() && dataRow["PACK_REV"].ToString() == ItemSet && checkdtAdd != true && dataRow["SEQ_NO"].ToString() != "1")
                        {
                            loop = loop + 1; checkpassOnce = true;
                        }
                        A = dataRow["PACK_ID"].ToString();
                        if (loop == 1 && dataRow["PACK_TYPE"].ToString() == "REEL")
                        {
                            checkpassOnce = false;
                            checkdtAdd = false;
                            OPTN = dt_result.NewRow();
                            PackOut = dt_result.NewRow();
                            OPTN["WI_PACK_ID"] = dataRow["PACK_ID"].ToString() + "_" + dataRow["PACK_REV"].ToString() + "_OPTN";
                            PackOut["WI_PACK_ID"] = dataRow["PACK_ID"].ToString() + "_" + dataRow["PACK_REV"].ToString();
                            OPTN["WI_TYPE"] = "Generic";
                            OPTN["DESCRIPTION"] = dataRow["PACK_DESCRIPTION"].ToString();
                            OPTN["INSTRUC_OPTN"] = "TNR";
                            //if (dataRow["HTB"].ToString() == "IMMED")
                            //{
                            //    OPTN["HTB"] = "P";
                            //}
                            //else
                            //{
                            //    OPTN["HTB"] = dataRow["HTB"].ToString();
                            //}
                            OPTN["HTB"] = dataRow["HTB"].ToString();
                            OPTN["UNIT_PER_REEL"] = dataRow["UNIT"].ToString();
                            OPTN["UNIT_PLACEMENT"] = "Live bug";
                            OPTN["LABEL_POSITION"] = "Sprocket hole";
                            OPTN["CREATED_BY"] = "System";
                            OPTN["CREATED_BY_NAME"] = "System";
                            OPTN["CREATED_DATE"] = mytime;
                            OPTN["UPDATED_BY"] = "System";
                            OPTN["UPDATED_BY_NAME"] = "System";
                            OPTN["UPDATED_DATE"] = mytime;
                            OPTN["UNIQUE_ID"] = "0";
                            OPTN["STATUS"] = "1";
                            PackOut["WI_TYPE"] = "Generic";
                            PackOut["DESCRIPTION"] = dataRow["PACK_DESCRIPTION"].ToString();
                            PackOut["INSTRUC_OPTN"] = "Pack Out";
                            PackOut["PACKOUT_TYPE"] = dataRow["METHOD"].ToString();
                            PackOut["L1_UNIT_PER_REEL"] = dataRow["UNIT"].ToString();
                            PackOut["L1_CUST_LABEL_FLAG"] = "No";
                            PackOut["L1_ESD_FLAG"] = "N";
                            PackOut["L1_PROTECTIVE_FLAG"] = "No";
                            PackOut["L2_DRY_PACK_FLAG"] = "No";
                            PackOut["L2_CACUUM_SEAL_FLAG"] = "No";
                            PackOut["L2_CUST_LABEL_FLAG"] = "No";
                            PackOut["L2_ESD_FLAG"] = "N";
                            PackOut["L2_CAUTION_FLAG"] = "N";
                            PackOut["L2_HIC_FLAG"] = "No";
                            PackOut["L2_DESICCANT_FLAG"] = "No";
                            PackOut["L3_CUST_LABEL_FLAG"] = "No";
                            PackOut["L3_ESD_FLAG"] = "N";
                            PackOut["L3_BUBBLE_FLAG"] = "No";
                            PackOut["L3_CAUTION_FLAG"] = "N";
                            PackOut["UNIQUE_ID"] = "";
                            PackOut["CREATED_BY"] = "System";
                            PackOut["CREATED_BY_NAME"] = "System";
                            PackOut["CREATED_DATE"] = mytime;
                            PackOut["UPDATED_BY"] = "System";
                            PackOut["UPDATED_BY_NAME"] = "System";
                            PackOut["UPDATED_DATE"] = mytime;
                            PackOut["UNIQUE_ID"] = "0";
                            PackOut["STATUS"] = "1";
                            PackOut["L1_HTB"] = dataRow["HTB"].ToString();
                        }
                        if (dataRow["PACK_TYPE"].ToString() == "BAG")
                        {
                            PackOut["L2_UNIT_PER_BAG"] = dataRow["UNIT"].ToString();
                            PackOut["L2_QTY_REEL_PER_BAG"] = dataRow["PACK_QTY"].ToString();
                        }
                        if (dataRow["PACK_TYPE"].ToString() == "BOX")
                        {
                            PackOut["L3_QTY_UNIT_PER_BOX"] = dataRow["UNIT"].ToString();
                            PackOut["L3_QTY_REEL_PER_BOX"] = dataRow["PACK_QTY"].ToString();
                            //PackOut["L2_UNIT_PER_BAG"] = dataRow["UNIT"].ToString();
                            //PackOut["L2_QTY_REEL_PER_BAG"] = dataRow["PACK_QTY"].ToString();
                        }
                        if (dataRow["PACK_TYPE"].ToString() == "LABEL" && dataRow["PACK_QTY"].ToString() != "")
                        {
                            PackOut["L3_CUST_LABEL_FLAG"] = "Yes";
                            PackOut["L3_CUST_LABEL_QTY"] = dataRow["PACK_QTY"].ToString();
                        }
                        if (dataRow["PACK_TYPE"].ToString() == "TAPE" && dataRow["SEQ_NO"].ToString() != "4")
                        {
                            PackOut["L3_QTY_TAPE_LINE"] = dataRow["PACK_QTY"].ToString();
                        }
                        //if(dataRow["PACK_ID"].ToString() == "IFX003/A")
                        //{
                        //    MessageBox.Show("");
                        //}
                        if (dataRow["PACK_TYPE"].ToString() == "LEADER_MIN")
                        {
                            OPTN["LEADER_POCKET_MAX"] = dataRow["PACK_QTY"].ToString();
                            OPTN["LEADER_POCKET_MIN"] = dataRow["PACK_QTY"].ToString();
                        }
                        if (dataRow["PACK_TYPE"].ToString() == "TRAILER_MIN")
                        {
                            OPTN["TRAILER_POCKET_MAX"] = dataRow["PACK_QTY"].ToString();
                            OPTN["TRAILER_POCKET_MIN"] = dataRow["PACK_QTY"].ToString();
                            if (OPTN["PIN1_ORIENTATION"].ToString() != "")
                            {
                                dt_result.Rows.Add(PackOut);
                                dt_result.Rows.Add(OPTN);
                                CountTNR = CountTNR + 1;
                                loop = 1;
                                check = false;
                                checkdtAdd = true;
                                Type = "Another";
                            }
                        }
                        if (dataRow["PACK_QTY"].ToString() != "" && dataRow["CAUTION"].ToString() == "Y")
                        {
                            switch (dataRow["PACK_LEVEL"].ToString())
                            {
                                case "IN1":
                                    PackOut["L2_CAUTION_FLAG"] = "Y"; PackOut["L2_CAUTION_QTY"] = dataRow["PACK_QTY"].ToString();
                                    break;
                                case "IN2":
                                    PackOut["L3_CAUTION_FLAG"] = "Y"; PackOut["L3_CAUTION_QTY"] = dataRow["PACK_QTY"].ToString();
                                    break;
                            }
                        }
                        if (dataRow["PACK_QTY"].ToString() == "" && dataRow["CAUTION"].ToString() == "Y")
                        {
                            switch (dataRow["PACK_LEVEL"].ToString())
                            {
                                case "IN1":
                                    PackOut["L2_CAUTION_FLAG"] = "Y"; PackOut["L2_CAUTION_QTY"] = "0";
                                    break;
                                case "IN2":
                                    PackOut["L3_CAUTION_FLAG"] = "Y"; PackOut["L3_CAUTION_QTY"] = "0";
                                    break;
                            }
                        }
                        if (dataRow["PACK_TYPE"].ToString() == "QUADRANT")
                        {
                            switch (dataRow["STOCK_NO"].ToString())
                            {
                                case "QUAD_1": OPTN["PIN1_ORIENTATION"] = "Quadrant 1"; break;
                                case "QUAD_2": OPTN["PIN1_ORIENTATION"] = "Quadrant 2"; break;
                                case "QUAD_3": OPTN["PIN1_ORIENTATION"] = "Quadrant 3"; break;
                                case "QUAD_4": OPTN["PIN1_ORIENTATION"] = "Quadrant 4"; break;
                            }

                            if (OPTN["LEADER_POCKET_MAX"].ToString() != "")
                            {
                                dt_result.Rows.Add(PackOut);
                                dt_result.Rows.Add(OPTN);
                                CountTNR = CountTNR + 1;
                                loop = 1;
                                check = false;
                                checkdtAdd = true;
                                Type = "Another";
                            }
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
                            PackOut = dt_result.NewRow();
                            if (dataRow["SEQ_NO"].ToString() != "1")
                            {
                                PackOut["WI_PACK_ID"] = "ERROR On ";
                                PackOut["WI_TYPE"] = dataRow["PACK_ID"].ToString() + "_" + dataRow["PACK_REV"].ToString();
                                PackOut["ENGINEERING_CODE"] = " Seq=" + dataRow["SEQ_NO"].ToString() + " Method=" + dataRow["METHOD"].ToString() + " PackType =" + dataRow["PACK_TYPE"].ToString();
                                Count2first = Count2first + 1;
                            }
                            else
                            {
                                PackOut["WI_PACK_ID"] = "ERROR On ";
                                PackOut["WI_TYPE"] = dataRow["PACK_ID"].ToString() + "_" + dataRow["PACK_REV"].ToString();
                                PackOut["ENGINEERING_CODE"] = " Method=" + dataRow["METHOD"].ToString() + " PackType =" + dataRow["PACK_TYPE"].ToString();
                            }
                            getErrorPoint = getErrorPoint + PackOut["WI_PACK_ID"].ToString() + PackOut["WI_TYPE"].ToString() + PackOut["ENGINEERING_CODE"].ToString() + "\n";
                            errorCount = errorCount + 1;
                            dt_result.Rows.Add(PackOut);
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
                if (checkdtAdd == false && checkpassOnce == true && Type != "Another" && loop != 1)
                {
                    dt_result.Rows.Add(PackOut);
                    dt_result.Rows.Add(OPTN);
                    switch (Type)
                    {
                        case "CAN":
                            CountCAN = CountCAN + 1;
                            break;
                        case "TNR":
                            CountTNR = CountTNR + 1
                             ; break;
                        case "TRA":
                            CountTRA = CountTRA + 1; break;
                        case "WAF":
                            CountWAF = CountWAF + 1; break;
                        case "RAI":
                            if (sideTye == "FILM FRAME")
                            {
                                CountWAF = CountWAF + 1;
                            }
                            if (sideTye == "TUBE")
                            {
                                CountRAI = CountRAI + 1;
                            }
                            break;
                        case "BAG":
                            CountBAG = CountBAG + 1; break;
                    }
                }
                #endregion
                //forthe last error
                #endregion

                int resultROW = dt_result.Rows.Count;
                int ItemsSuccess = (resultROW - errorCount) / 2;
                string AllItems = Convert.ToString(ItemsSuccess + errorCount);
                string Allinput = Convert.ToString(countItem + Count2first);
                string Missingv = Convert.ToString((countItem + Count2first) - (ItemsSuccess + errorCount));
                ERRORget = "AllInput =" + Allinput + 
                    "\n ALLItemDetect =" + AllItems + 
                    "\n ItemsGenSuccess = " + ItemsSuccess.ToString() +
                    "\n CAN =" + CountCAN.ToString() + 
                    "\n TNR =" + CountTNR.ToString() + 
                    "\n TRA =" + CountTRA.ToString() + 
                    "\n WAF =" + CountWAF.ToString() +
                    "\n RAI =" + CountRAI.ToString() + 
                    "\n BAG =" + CountBAG.ToString() + 
                    "\n Missing Value = " + Missingv +
                    "\n Error value Count = " + errorCount.ToString() + 
                    "\n" + getErrorPoint.ToString();
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
                            using (FileStream fs = File.Create(@fbd.SelectedPath + "\\" + filename + " ERROR" + ".txt"))
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
            MessageBox.Show("Welcome to WI Pack Migration Data ,\nToday is " + mytime + " \n Have A Good Day.");
        }

        private void tbFile_TextChanged(object sender, EventArgs e)
        {
            if (tbFile.Text != "" || tbFile.Text != " ")
            {
                btnExport.Enabled = true;
            }
        }

        private void labelExample_Click(object sender, EventArgs e)
        {
            string RunningPath = AppDomain.CurrentDomain.BaseDirectory;
            string FileName = string.Format(Path.GetFullPath(Path.Combine(RunningPath, @"ExampleINPUT\Example Data.xlsx")));
            tbBrowse.Text = FileName;
        }

        private void labelHelp_Click(object sender, EventArgs e)
        {
            string RunningPath = AppDomain.CurrentDomain.BaseDirectory;
            string FileName = string.Format(Path.GetFullPath(Path.Combine(RunningPath, @"Manual\WI PACK Manual.pdf")));
            System.Diagnostics.Process.Start(FileName);

        }
    }
}
