using System;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Business_Case_Reader
{
    public partial class MainForm : Form
    {
        /*
         * Main Form and entry point
         */
        private string fileName = "";
        public MainForm()
        {
            //Initialize component and setup controls
            InitializeComponent();
            btn_browseBC.Enabled = true;
            bth_ReadSheet.Enabled = false;
            btn_ExportCSV.Enabled = false;

            //Prevent resize, remove minimize and whole screen button
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
        }
        private void btn_browseBC_Click(object sender, EventArgs e)
        {
            /*
             * Event handler for Business Case File analysis
             * User has to select excel macro file and this function is going to read it
             * Main logic: Loop all sheets and search for string FINANCIAL BUSINESS CASE in A1
             * In case 0 found - inform user to select correct file
             * In case 1 found - process directly with sheet analysis. Call function AnalizirajFajl
             * In case >1 found - inform user to select sheet for analysis
            */
            if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                //Get file name and write it into textbox
                fileName = openFileDialog.FileName;
                tb_FileName.Text = Path.GetFileName(openFileDialog.FileName);

                //Prepare visibility and enable/disable what has to be prepared
                pic_Loading.Visible = true;
                btn_browseBC.Enabled = true;
                bth_ReadSheet.Enabled = false;
                btn_ExportCSV.Enabled = false;
                
                //Refresh controles to be visible for user
                pic_Loading.Refresh();
                tb_FileName.Refresh();
                btn_browseBC.Refresh();
                bth_ReadSheet.Refresh();
                btn_ExportCSV.Refresh();


                // Analyze attached file
                AnalizirajFajl(fileName);

                //Setup controls
                pic_Loading.Visible = false;
                tb_FileName.Text = "";

                //Refresh visibility
                pic_Loading.Refresh();
                tb_FileName.Refresh();


            }
        }
        public static List<TranlationTable> ReturnTranslationTable()
        {
            /*
             * Result: build List that will be used in sheet analysis.
             * In current project scope, list has to be generated manually and controlled within the code.
             * It would be best to have list at some server and to call it.
             * It would allow list flexibility
             */
            List<TranlationTable> trTable = new List<TranlationTable>();

            trTable.Add(new TranlationTable() { id = 1, description = "LTA", exportValue = null, fieldForValue = "K17", textToCheck = "LTA:", fieldForCheck = "J17", year = null, isOK = false, isForTBL = true, TBLid = 1, isHeader =true, ExportPart=3 });
            trTable.Add(new TranlationTable() { id = 2, description = "CI", exportValue = null, fieldForValue = "G17", textToCheck = "CI:", fieldForCheck = "F17", year = null, isOK = false, isForTBL = true, TBLid = 2, isHeader = true, ExportPart = 3 });
            trTable.Add(new TranlationTable() { id = 3, description = "Fuel", exportValue = null, fieldForValue = "I17", textToCheck = "Fuel:", fieldForCheck = "H17", year = null, isOK = false, isForTBL = true, TBLid = 3, isHeader = true, ExportPart = 3 });
            trTable.Add(new TranlationTable() { id = 4, description = "Wage Economics", exportValue = null, fieldForValue = "E17", textToCheck = "Wage Economics:", fieldForCheck = "D17", year = null, isOK = false, isForTBL = true, TBLid = 4, isHeader = true, ExportPart = 3 });
            //trTable.Add(new TranlationTable() { id = 5, description = "New Bldg", exportValue = null, fieldForValue = "G17", textToCheck = "New Bldg:", fieldForCheck = "F17", year = null, isOK = false, isForTBL = true,TBLid = 5, isHeader=true, ExportPart=3 });
            trTable.Add(new TranlationTable() { id = 6, description = "WACC", exportValue = null, fieldForValue = "X13", textToCheck = "WACC:", fieldForCheck = "W12", year = null, isOK = false, isForTBL = true, TBLid = 6, isHeader = true, ExportPart = 3 });
            trTable.Add(new TranlationTable() { id = 7, description = "Material Savings", exportValue = null, fieldForValue = "E19", textToCheck = "Wage Economics:", fieldForCheck = "D17", year = null, isOK = false, isForTBL = true, TBLid = 7, isHeader = true, ExportPart = 3 });
            trTable.Add(new TranlationTable() { id = 8, description = "Material Ecconomics", exportValue = null, fieldForValue = "E25", textToCheck = "Wage Economics:", fieldForCheck = "D17", year = null, isOK = false, isForTBL = true, TBLid = 8, isHeader = true, ExportPart = 3 });
            trTable.Add(new TranlationTable() { id = 9, description = "Major Capex Items", exportValue = null, fieldForValue = "E12", textToCheck = "Major Capex Items:", fieldForCheck = "C12", year = null, isOK = false, isForTBL = true, TBLid = 9, isHeader = true, ExportPart = 3 });

            //Expenditure Outflow for Capitalized Items
            trTable.Add(new TranlationTable() { id = 10, description = "New Equipment/Molding", exportValue = null, fieldForValue = "35,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,X", textToCheck = "New Equipment/Molding:", fieldForCheck = "C35", year = null, isOK = false, isForTBL = false, TBLid = 0, isHeader = false, ExportPart = 3 });
            trTable.Add(new TranlationTable() { id = 11, description = "New Building and Land", exportValue = null, fieldForValue = "36,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,X", textToCheck = "New Building and Land:", fieldForCheck = "C36", year = null, isOK = false, isForTBL = false, TBLid = 0, isHeader = false, ExportPart = 3 });
            trTable.Add(new TranlationTable() { id = 12, description = "New JCI Owned Prototype Tooling", exportValue = null, fieldForValue = "37,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,X", textToCheck = "New JCI Owned Prototype Tooling:", fieldForCheck = "C37", year = null, isOK = false, isForTBL = false, TBLid = 0, isHeader = false, ExportPart = 3 });
            trTable.Add(new TranlationTable() { id = 13, description = "New JCI Owned Production Tooling", exportValue = null, fieldForValue = "38,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,X", textToCheck = "New JCI Owned Production Tooling:", fieldForCheck = "C38", year = null, isOK = false, isForTBL = false, TBLid = 0, isHeader = false, ExportPart = 3 });
            trTable.Add(new TranlationTable() { id = 14, description = "Engineering R&D and other Technical Services", exportValue = null, fieldForValue = "39,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,X", textToCheck = "Engineering R&D and other Technical Services:", fieldForCheck = "C39", year = null, isOK = false, isForTBL = false, TBLid = 0, isHeader = false, ExportPart = 3 });
            trTable.Add(new TranlationTable() { id = 15, description = "CapEx and capitalized ER&D / Other", exportValue = null, fieldForValue = "40,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,X", textToCheck = "CapEx and capitalized ER&D / Other:", fieldForCheck = "C40", year = null, isOK = false, isForTBL = false, TBLid = 0, isHeader = false, ExportPart = 3 });

            //Tooling Requirements (non-capitalized)
            trTable.Add(new TranlationTable() { id = 16, description = "Prototype Tooling Expenditures", exportValue = null, fieldForValue = "43,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,X", textToCheck = "Prototype Tooling Expenditures:", fieldForCheck = "C43", year = null, isOK = false, isForTBL = false, TBLid = 0, isHeader = false, ExportPart = 3 });
            trTable.Add(new TranlationTable() { id = 17, description = "Production Tooling Expenditures", exportValue = null, fieldForValue = "44,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,X", textToCheck = "Production Tooling Expenditures:", fieldForCheck = "C44", year = null, isOK = false, isForTBL = false, TBLid = 0, isHeader = false, ExportPart = 3 });
            trTable.Add(new TranlationTable() { id = 18, description = "Prototype Tooling Reimbursements", exportValue = null, fieldForValue = "45,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,X", textToCheck = "Prototype Tooling Reimbursements:", fieldForCheck = "C45", year = null, isOK = false, isForTBL = false, TBLid = 0, isHeader = false, ExportPart = 3 });
            trTable.Add(new TranlationTable() { id = 19, description = "Production Tooling Reimbursements", exportValue = null, fieldForValue = "46,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,X", textToCheck = "Production Tooling Reimbursements:", fieldForCheck = "C46", year = null, isOK = false, isForTBL = false, TBLid = 0, isHeader = false, ExportPart = 3 });
            trTable.Add(new TranlationTable() { id = 20, description = "Net Tooling Expenditures", exportValue = null, fieldForValue = "47,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,X", textToCheck = "Net Tooling Expenditures:", fieldForCheck = "C47", year = null, isOK = false, isForTBL = false, TBLid = 0, isHeader = false, ExportPart = 3 });

            //Working Capital Requirements
            trTable.Add(new TranlationTable() { id = 21, description = "Accounts Receivable", exportValue = null, fieldForValue = "50,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,X", textToCheck = "Accounts Receivable:", fieldForCheck = "C50", year = null, isOK = false, isForTBL = false, TBLid = 0, isHeader = false, ExportPart = 3 });
            trTable.Add(new TranlationTable() { id = 22, description = "Inventory", exportValue = null, fieldForValue = "51,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,X", textToCheck = "Inventory:", fieldForCheck = "C51", year = null, isOK = false, isForTBL = false, TBLid = 0, isHeader = false, ExportPart = 3 });
            trTable.Add(new TranlationTable() { id = 23, description = "Accounts Payable", exportValue = null, fieldForValue = "52,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,X", textToCheck = "Accounts Payable:", fieldForCheck = "C52", year = null, isOK = false, isForTBL = false, TBLid = 0, isHeader = false, ExportPart = 3 });
            trTable.Add(new TranlationTable() { id = 24, description = "Net Working Capital", exportValue = null, fieldForValue = "53,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,X", textToCheck = "Net Working Capital:", fieldForCheck = "C53", year = null, isOK = false, isForTBL = false, TBLid = 0, isHeader = false, ExportPart = 3 });

            //Investment Change Net Inflow/Outflow
            //trTable.Add(new TranlationTable() { id = 25, description = "Investment Change Net Inflow/Outflow", exportValue = null, fieldForValue = "55,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,X", textToCheck = "Investment Change Net Inflow/Outflow:", fieldForCheck = "C55", year = null, isOK = false, isForTBL = false, TBLid = 0, isHeader = false, ExportPart = 3 });

            //Customer Production Volume
            trTable.Add(new TranlationTable() { id = 26, description = "Customer Production Volume", exportValue = null, fieldForValue = "62,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,X", textToCheck = "Customer Production Volume:", fieldForCheck = "C62", year = null, isOK = false, isForTBL = false, TBLid = 0, isHeader = false, ExportPart = 3 });

            //Production Planning Volume  - TBD this contains formula. Check with Sven how to handle it
            trTable.Add(new TranlationTable() { id = 27, description = "Production Planning Volume", exportValue = null, fieldForValue = "64,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,X", textToCheck = "Volumes", fieldForCheck = "C58", year = null, isOK = false, isForTBL = false, TBLid = 0, isHeader = false, ExportPart = 3 });

            //Capacity Planning Volume - TBD Same as Production Planning volume
            trTable.Add(new TranlationTable() { id = 28, description = "Capacity Planning Volume", exportValue = null, fieldForValue = "66,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,X", textToCheck = "Volumes", fieldForCheck = "C58", year = null, isOK = false, isForTBL = false, TBLid = 0, isHeader = false, ExportPart = 3 });

            //Sales
            trTable.Add(new TranlationTable() { id = 29, description = "Sales (@ SOP Prices)", exportValue = null, fieldForValue = "69,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,X", textToCheck = "Sales (@ SOP Prices):", fieldForCheck = "C69", year = null, isOK = false, isForTBL = false, TBLid = 0, isHeader = false, ExportPart = 3 });
            trTable.Add(new TranlationTable() { id = 30, description = "Price Reductions / LTA", exportValue = null, fieldForValue = "70,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,X", textToCheck = "Price Reductions / LTA:", fieldForCheck = "C70", year = null, isOK = false, isForTBL = false, TBLid = 0, isHeader = false, ExportPart = 3 });
            trTable.Add(new TranlationTable() { id = 31, description = "Capitalized Engineering R&D and other Technical Services Depreciation", exportValue = null, fieldForValue = "71,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,X", textToCheck = "Capitalized Engineering R&D and other Technical Services Depreciation:", fieldForCheck = "C71", year = null, isOK = false, isForTBL = false, TBLid = 0, isHeader = false, ExportPart = 3 });
            trTable.Add(new TranlationTable() { id = 32, description = "New Prototype Tooling Amortization", exportValue = null, fieldForValue = "72,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,X", textToCheck = "New Prototype Tooling Amortization:", fieldForCheck = "C72", year = null, isOK = false, isForTBL = false, TBLid = 0, isHeader = false, ExportPart = 3 });
            trTable.Add(new TranlationTable() { id = 33, description = "New Production Tooling Amortization", exportValue = null, fieldForValue = "73,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,X", textToCheck = "New Production Tooling Amortization:", fieldForCheck = "C73", year = null, isOK = false, isForTBL = false, TBLid = 0, isHeader = false, ExportPart = 3 });
            trTable.Add(new TranlationTable() { id = 34, description = "Engineering R&D and other Technical Services Amortization", exportValue = null, fieldForValue = "74,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,X", textToCheck = "Engineering R&D and other Technical Services Amortization:", fieldForCheck = "C74", year = null, isOK = false, isForTBL = false, TBLid = 0, isHeader = false, ExportPart = 3 });
            trTable.Add(new TranlationTable() { id = 35, description = "Launch / Start-Up Amortization", exportValue = null, fieldForValue = "75,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,X", textToCheck = "Launch / Start-Up Amortization:", fieldForCheck = "C75", year = null, isOK = false, isForTBL = false, TBLid = 0, isHeader = false, ExportPart = 3 });
            trTable.Add(new TranlationTable() { id = 36, description = "Sales (after Price Reductions / LTA / Amortizations)", exportValue = null, fieldForValue = "76,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,X", textToCheck = "Sales (after Price Reductions / LTA / Amortizations):", fieldForCheck = "C76", year = null, isOK = false, isForTBL = false, TBLid = 0, isHeader = false, ExportPart = 3 });

            // Cost of Production
            trTable.Add(new TranlationTable() { id = 37, description = "Material Cost", exportValue = null, fieldForValue = "79,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,X", textToCheck = "Material Cost:", fieldForCheck = "C79", year = null, isOK = false, isForTBL = false, TBLid = 0, isHeader = false, ExportPart = 3 });
            trTable.Add(new TranlationTable() { id = 38, description = "Foreseen Material Cost Reductions", exportValue = null, fieldForValue = "80,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,X", textToCheck = "Foreseen Material Cost Reductions:", fieldForCheck = "C80", year = null, isOK = false, isForTBL = false, TBLid = 0, isHeader = false, ExportPart = 3 });
            trTable.Add(new TranlationTable() { id = 39, description = "Variable Conversion Cost(incl.Freight in)", exportValue = null, fieldForValue = "81,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,X", textToCheck = "Variable Conversion Cost (incl. Freight in):", fieldForCheck = "C81", year = null, isOK = false, isForTBL = false, TBLid = 0, isHeader = false, ExportPart = 3 });
            trTable.Add(new TranlationTable() { id = 40, description = "Foreseen Variable Conversion Cost Reductions", exportValue = null, fieldForValue = "82,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,X", textToCheck = "Foreseen Variable Conversion Cost Reductions:", fieldForCheck = "C82", year = null, isOK = false, isForTBL = false, TBLid = 0, isHeader = false, ExportPart = 3 });
            trTable.Add(new TranlationTable() { id = 41, description = "Outbound Costs", exportValue = null, fieldForValue = "83,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,X", textToCheck = "Outbound Costs:", fieldForCheck = "C83", year = null, isOK = false, isForTBL = false, TBLid = 0, isHeader = false, ExportPart = 3 });
            trTable.Add(new TranlationTable() { id = 42, description = "Foreseen Outbound Cost Reductions", exportValue = null, fieldForValue = "84,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,X", textToCheck = "Foreseen Outbound Cost Reductions:", fieldForCheck = "C84", year = null, isOK = false, isForTBL = false, TBLid = 0, isHeader = false, ExportPart = 3 });
            trTable.Add(new TranlationTable() { id = 43, description = "Net Tooling Expenditure", exportValue = null, fieldForValue = "85,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,X", textToCheck = "Net Tooling Expenditure:", fieldForCheck = "C85", year = null, isOK = false, isForTBL = false, TBLid = 0, isHeader = false, ExportPart = 3 });
            trTable.Add(new TranlationTable() { id = 44, description = "Other Costs net of Cost Reductions", exportValue = null, fieldForValue = "86,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,X", textToCheck = "Other Costs net of Cost Reductions:", fieldForCheck = "C86", year = null, isOK = false, isForTBL = false, TBLid = 0, isHeader = false, ExportPart = 3 });
            trTable.Add(new TranlationTable() { id = 45, description = "Variable Costs", exportValue = null, fieldForValue = "87,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,X", textToCheck = "Variable Costs:", fieldForCheck = "C87", year = null, isOK = false, isForTBL = false, TBLid = 0, isHeader = false, ExportPart = 3 });

            //Contribution Margin
            trTable.Add(new TranlationTable() { id = 46, description = "Contribution  Margin", exportValue = null, fieldForValue = "88,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,X", textToCheck = "Contribution  Margin:", fieldForCheck = "C88", year = null, isOK = false, isForTBL = false, TBLid = 0, isHeader = false, ExportPart = 3 });
            trTable.Add(new TranlationTable() { id = 47, description = "Contribution  Margin %", exportValue = null, fieldForValue = "89,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,X", textToCheck = "Contribution  Margin %:", fieldForCheck = "C89", year = null, isOK = false, isForTBL = false, TBLid = 10, isHeader = false, ExportPart = 3 });

            //Cost of Development
            trTable.Add(new TranlationTable() { id = 47, description = "Non - capitalized Engineering R&D and other Technical Services", exportValue = null, fieldForValue = "100,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,X", textToCheck = "Non-capitalized Engineering R&D and other Technical Services:", fieldForCheck = "C100", year = null, isOK = false, isForTBL = false, TBLid = 0, isHeader = false, ExportPart = 3 });
            trTable.Add(new TranlationTable() { id = 48, description = "Non - capitalized Engineering R&D and other Technical Services Reimbursements", exportValue = null, fieldForValue = "101,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,X", textToCheck = "Non-capitalized Engineering R&D and other Technical Services Reimbursements:", fieldForCheck = "C101", year = null, isOK = false, isForTBL = false, TBLid = 0, isHeader = false, ExportPart = 3 });
            trTable.Add(new TranlationTable() { id = 49, description = "Selling, General and Administrative(SG & A)", exportValue = null, fieldForValue = "102,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,X", textToCheck = "Selling, General and Administrative (SG&A):", fieldForCheck = "C102", year = null, isOK = false, isForTBL = false, TBLid = 0, isHeader = false, ExportPart = 3 });
            trTable.Add(new TranlationTable() { id = 50, description = "Additional SG&A resulting from this Program being associated with a 'Global Program'", exportValue = null, fieldForValue = "103,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,X", textToCheck = "Additional SG&A resulting from this Program being associated with a 'Global Program':", fieldForCheck = "C103", year = null, isOK = false, isForTBL = false, TBLid = 0, isHeader = false, ExportPart = 3 });
            trTable.Add(new TranlationTable() { id = 51, description = "Launch / Start - Up Costs", exportValue = null, fieldForValue = "104,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,X", textToCheck = "Launch / Start-Up Costs:", fieldForCheck = "C104", year = null, isOK = false, isForTBL = false, TBLid = 0, isHeader = false, ExportPart = 3 });
            trTable.Add(new TranlationTable() { id = 52, description = "Launch / Start - Up Cost Reimbursements", exportValue = null, fieldForValue = "105,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,X", textToCheck = "Launch / Start-Up Cost Reimbursements:", fieldForCheck = "C105", year = null, isOK = false, isForTBL = false, TBLid = 0, isHeader = false, ExportPart = 3 });
            trTable.Add(new TranlationTable() { id = 53, description = "Other Costs", exportValue = null, fieldForValue = "106,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,X", textToCheck = "Other Costs:", fieldForCheck = "C106", year = null, isOK = false, isForTBL = false, TBLid = 0, isHeader = false, ExportPart = 3 });
            trTable.Add(new TranlationTable() { id = 54, description = "Cost of Development", exportValue = null, fieldForValue = "107,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,X", textToCheck = "Cost of Development:", fieldForCheck = "C107", year = null, isOK = false, isForTBL = false, TBLid = 0, isHeader = false, ExportPart = 3 });

            //SINC
            trTable.Add(new TranlationTable() { id = 55, description = "SINC", exportValue = null, fieldForValue = "110,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,X", textToCheck = "SINC:", fieldForCheck = "C110", year = null, isOK = false, isForTBL = false, TBLid = 0, isHeader = false, ExportPart = 3 });


            return trTable;
        }
        void AnalizirajFajl(string fileName)
        {
            /*
             *Result: This function is going to try to read provided file. File has to be closed in order to OpenXML works.
             *        Function is looping all sheets and searches string FINANCIAL BUSINESS CASE in A1 and based on that it 
             *        has three possibilities. In case 0 sheets found, open second tab and inform user to select correct file
             *        In case 1 sheet found call function ReadSheet() which will analyse sheet
             *        In case more than 1 found, populate listbox with those sheets names and inform user to select one of them
             */

            //Delete all items from previous analysis
            lb_SheetNames.Items.Clear(); 

            //Try to read file. In case it is open Exception will occure. Catch it and inform user to close file. General exception places as well.
            try
            {
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
                {
                    
                    // Retrieve a reference to the workbook part.
                    WorkbookPart wbPart = document.WorkbookPart;

                    //Prepare working enviroment
                    pic_Loading.Visible = true;
                    var results = GetAllWorksheets(fileName);
                    Sheet rememberSheet = null;
                    int brojac = 0;

                    //Main loop
                    foreach (Sheet item in results)
                    {
                        string StaPiseUA1 = ProcitajNaSheetu(wbPart, item, "A1");
                        if (StaPiseUA1 == "FINANCIAL BUSINESS CASE")
                        {
                            lb_SheetNames.Items.Add(item.Name);
                            rememberSheet = item;
                            brojac++;
                        }
                    }

                    //Three alternatives now: brojac is 0, 1 or >1
                    if (brojac == 0) //No sheets with FBC found. Can not process further.
                    {
                        lbl_Info.Text = "Attached File is missing Sheet \n\nFINANCIAL BUSINESS CASE \nPlease select correct file!!!";
                        lbl_Info.ForeColor = System.Drawing.Color.Red;
                        tabovi.SelectedIndex = 1;

                    } else if(brojac == 1) //This is what it is supposed to be. Process dirrectly to analyze code
                    {
                        lbl_Info.Text = "FINANCIAL BUSINESS CASE sheet recognized!!!";
                        lbl_Info.ForeColor = System.Drawing.Color.White;
                        ReadSheet(wbPart, rememberSheet); //Analyze code

                    } else if (brojac >= 2) //Multiple FBC sheet. User needs to select correct one and then call analyzing code
                    {
                        lbl_Info.Text = "Multiple FINANCIAL BUSINESS CASE sheeets. \n\nPlease select correct and press button Read Sheet!!!";
                        lbl_Info.ForeColor = System.Drawing.Color.White;
                        tabovi.SelectedIndex = 1;
                    }

                }
            }
            catch(IOException)
            {
                MessageBox.Show("Please close file: " + fileName + "!!!");
            }

            catch (Exception ec)
            {
                MessageBox.Show("Something went wrong!!!");
            }
        }
        public void ReadSheet(WorkbookPart wbPart, Sheet item)
        {
            /*
             * Reads provided sheet in provided workbook based on logic provided from function ReturnTranslationTable()
             * Reports out result into third tab as table formated data stream which can be exported in csv or txt file
             */


            //Setup variables necessary for algorithm logic
            bool mainCheck = true;
            List<TranlationTable> translatedTable = ReturnTranslationTable();
            List<ExportTable> exTable = new List<ExportTable>();

            string PocetnaGodina = ProcitajNaSheetu(wbPart, item, "D31");
            PocetnaGodina = PocetnaGodina.Substring((PocetnaGodina.Length) - 4, 4);
            string KrajnjaGodina = ProcitajNaSheetu(wbPart, item, "V31");

            //Loop Translation Table and match agains sheet data
            foreach (TranlationTable lineObject in translatedTable)
            {
                //Check if sheet possition matches expected template positions.
                if (ProcitajNaSheetu(wbPart, item, lineObject.fieldForCheck).Trim() == lineObject.textToCheck)
                {
                    //Update table isOk variable and find value at sheet
                    lineObject.isOK = true;

                    if (lineObject.isHeader) //Header has to have only one value
                    {
                        lineObject.exportValue = ProcitajNaSheetu(wbPart, item, lineObject.fieldForValue);

                        //Add updated item to the result list
                        exTable.Add(new ExportTable()
                        {
                            id = lineObject.id,
                            description = lineObject.description,
                            exportValue = lineObject.exportValue,
                            year = lineObject.year,
                            isOK = lineObject.isOK,
                            isForTBL = lineObject.isForTBL,
                            TBLid = lineObject.TBLid,
                            isHeader = lineObject.isHeader,
                            ExportPart = lineObject.ExportPart
                        });
                    } else if(!lineObject.isHeader)
                    {
                        //Loop each line and then add items
                        string source = lineObject.fieldForValue;
                        string[] stringSeparators = new string[] { "," };
                        string[] result;
                        result = source.Split(stringSeparators, StringSplitOptions.None);

                        for(int i = 1; i <= 20; i++)
                        {
                            //Year returns sometimes formula + value. I just need value
                            string godina = ProcitajNaSheetu(wbPart, item, result[i] + "31");
                            if (godina.Length > 4)
                            {
                                godina = godina.Substring(godina.Length - 4);
                                //Last one has Total as value and it should have 5 characters
                                if (godina == "otal")
                                {
                                    godina = "Total";
                                }
                            }

                            //Add only if value <> 0
                            string vrijednost = ProcitajNaSheetu(wbPart, item, result[i] + result[0]);
                            int duzina = vrijednost.Length;

                            //Some values has formula SUM(something) in it. Remove that part. Leave only value
                            if (vrijednost.Length>12 && vrijednost.Substring(0, 3) == "SUM")
                            {
                                int gdjeJeZagrada = vrijednost.IndexOf(")");
                                vrijednost = vrijednost.Substring(gdjeJeZagrada + 1);
                            }

                            //Some values are small like 1.124545454E-12. It should be 0
                            if (vrijednost.IndexOf("E-") > 0)
                            {
                                vrijednost = "0";
                            }

                            //We operate here with huge amounts. No need for decimal places. No need to convert value to decimal because there is a huge risk for Exceptions
                            if (vrijednost.IndexOf(".") > 0)
                            {
                                int pozicijaZareza = vrijednost.IndexOf(".");
                                vrijednost = vrijednost.Substring(0, pozicijaZareza);
                            }
                            

                            if (vrijednost != "0" && vrijednost != "" && vrijednost != null)
                            {
                                //Add updated item to the result list
                                exTable.Add(new ExportTable()
                                {
                                    id = lineObject.id,
                                    description = lineObject.description,
                                    exportValue = vrijednost,
                                    year = godina,
                                    isOK = lineObject.isOK,
                                    isForTBL = lineObject.isForTBL,
                                    TBLid = lineObject.TBLid,
                                    isHeader = lineObject.isHeader,
                                    ExportPart = lineObject.ExportPart
                                });
                            }
                            
                        } 
                    }
                    
                    
                }
                else
                {
                    //In case expected data is not placed at correct possition
                    mainCheck = false;
                }

                
            }
             
            //All values should be updated and checked against all values. If any of them is not correct show that to the user
            if (!mainCheck)
            {
                //TBD - what to do here? Inform user only or block export???
                MessageBox.Show("All values can not be checked!!!");
            }

            //use binding source to send Export Table to the datagridview so that user can see result
            BindingSource binding = new BindingSource();
            binding.DataSource = exTable;
            dataGridView.DataSource = binding;

            //Format datagrid columns
            dataGridView.AllowUserToResizeColumns = false;

            dataGridView.Columns["id"].Width = 30;
            dataGridView.Columns["id"].ReadOnly = true;

            dataGridView.Columns["description"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridView.Columns["description"].ReadOnly = true;

            dataGridView.Columns["year"].Width = 50;
            dataGridView.Columns["year"].ReadOnly = true;

            dataGridView.Columns["isForTBL"].Width = 50;
            dataGridView.Columns["isForTBL"].ReadOnly = true;

            dataGridView.Columns["isOK"].Width = 50;
            dataGridView.Columns["isOK"].ReadOnly = true;

            //Select third tab
            btn_ExportCSV.Enabled = true;
            tabovi.SelectedIndex = 2;
        }
        public static string ProcitajNaSheetu(WorkbookPart wbPart, Sheet theSheet, string whereToLook)
        {
            /*
             * Result - returns content from provided Sheet and Cell. 
             */

            //Prepare working variables
            string SadrzajUPolju = "";

            //Get worksheet
            WorksheetPart wsPart = (WorksheetPart)(wbPart.GetPartById(theSheet.Id));

            //Get cell
            Cell theCell = wsPart.Worksheet.Descendants<Cell>().Where(c => c.CellReference == whereToLook).FirstOrDefault();

            // If the cell does not exist, return an empty string.
            if (theCell != null)
            {
                SadrzajUPolju = theCell.InnerText;

                // If the cell represents an integer number return it. 
                // For dates, this code returns the serialized value that 
                // represents the date. The code handles strings and 
                // Booleans individually. For shared strings, the code 
                // looks up the corresponding value in the shared string 
                // table. For Booleans, the code converts the value into 
                // the words TRUE or FALSE.
                if (theCell.DataType != null)
                {
                    switch (theCell.DataType.Value)
                    {
                        case CellValues.SharedString:

                            // For shared strings, look up the value in the
                            // shared strings table.
                            var stringTable =
                                wbPart.GetPartsOfType<SharedStringTablePart>()
                                .FirstOrDefault();

                            // If the shared string table is missing, something 
                            // is wrong. Return the index that is in
                            // the cell. Otherwise, look up the correct text in 
                            // the table.
                            if (stringTable != null)
                            {
                                SadrzajUPolju = stringTable.SharedStringTable.ElementAt(int.Parse(SadrzajUPolju)).InnerText;
                            }
                            break;

                        case CellValues.Boolean:
                            switch (SadrzajUPolju)
                            {
                                case "0":
                                    SadrzajUPolju = "FALSE";
                                    break;
                                default:
                                    SadrzajUPolju = "TRUE";
                                    break;
                            }
                            break;
                    }
                }
            }
            return SadrzajUPolju;
        }
        public static  Sheets GetAllWorksheets(string fileName)
        {
            /*
             * Result - loops all sheets in the file and returns them
             */

            //Start as null
            Sheets theSheets = null;

            //Get sheets
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart wbPart = document.WorkbookPart;
                theSheets = wbPart.Workbook.Sheets;
            }

            //Return sheets
            return theSheets;
        }
        private void lb_SheetNames_SelectedIndexChanged(object sender, EventArgs e)
        {
            /*
             * Result - Read Sheet button is disabled by default. Enable it when user select list item
             */
            bth_ReadSheet.Enabled = true;
        }
        private void bth_ReadSheet_Click(object sender, EventArgs e)
        {
            /*
             * Result - Check if user has selected sheet. Find that sheet and read data
             */
            string userSelection = lb_SheetNames.GetItemText(lb_SheetNames.SelectedItem);
            if (userSelection == "")
            {
                MessageBox.Show("Please select Financial Business Case Sheet!");
            }
            else
            {
                try
                {
                    using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
                    {

                        // Retrieve a reference to the workbook part.
                        WorkbookPart wbPart = document.WorkbookPart;
                        var results = GetAllWorksheets(fileName);

                        //Main loop
                        foreach (Sheet item in results)
                        {
                            if (item.Name == userSelection)
                            {
                                ReadSheet(wbPart, item); //Analyze code
                            }
                        }
                    }

                }
                catch (IOException)
                {
                    MessageBox.Show("Please close file: " + fileName + "!!!");
                }

                catch (Exception)
                {
                    MessageBox.Show("Something went wrong!!!");
                }
            }
        }
        private void btn_ExportCSV_Click(object sender, EventArgs e)
        {
            /*
             * Calls function which exports datagrid
             */

            //TBD 2 - logic for export. If there are false in column isOK, should I export or ask user to update its template???
            SaveToCSV(dataGridView);
        }
        private void SaveToCSV(DataGridView DGV)
        {
            /*
             * Result - it exports datagrid in csv file. It raises output directory for user to save it somewhere.
             */

            //Prepare working variables
            string filename = "";
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "CSV (*.csv)|*.csv";
            sfd.FileName = "Output.csv"; //TBD 3 - File name can be connected with hash code.

            //Show dialog box and wait for OK
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                //Notify user
                MessageBox.Show("Data will be exported and you will be notified when it is ready.");

                //Try to replace file if already exists. Otherwise save new file
                if (File.Exists(filename))
                {
                    try
                    {
                        File.Delete(filename);
                    }
                    catch (IOException ex)
                    {
                        MessageBox.Show("It wasn't possible to write the data to the disk." + ex.Message);
                    }
                }

                //prepare variables for datagrid reading
                int columnCount = DGV.ColumnCount;
                string columnNames = "";
                string[] output = new string[DGV.RowCount + 1];

                //loop all columns and collect names for first row
                for (int i = 0; i < columnCount; i++)
                {
                    columnNames += DGV.Columns[i].Name.ToString() + ",";
                }
                output[0] += columnNames;

                //loop all data
                for (int i = 1; (i - 1) < DGV.RowCount; i++)
                {
                    for (int j = 0; j < columnCount; j++)
                    {
                        string enterValue;
                        if(DGV.Rows[i - 1].Cells[j].Value == null)
                        {
                            enterValue = "";
                        } else
                        {
                            enterValue = DGV.Rows[i - 1].Cells[j].Value.ToString();
                        }
                        output[i] += enterValue + ",";
                    }
                }

                //Send everything to previously created file
                System.IO.File.WriteAllLines(sfd.FileName, output, System.Text.Encoding.UTF8);

                //Notify user
                MessageBox.Show("Your file was generated and its ready for use.");
            }
        }
    }    
}
