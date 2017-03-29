using System;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;
using System.Collections.Generic;
using System.IO;

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
            //bth_ReadSheet.Enabled = false;
            //btn_ExportCSV.Enabled = false;
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
                //bth_ReadSheet.Enabled = false;
                //btn_ExportCSV.Enabled = false;
                
                //Refresh controles to be visible for user
                pic_Loading.Refresh();
                tb_FileName.Refresh();
                btn_browseBC.Refresh();
                //bth_ReadSheet.Refresh();
                //btn_ExportCSV.Refresh();


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

            trTable.Add(new TranlationTable() { id = 1, description = "LTA", exportValue = null, fieldForValue = "K17", textToCheck = "LTA:", fieldForCheck = "J17", year = null, isOK = false, isForTBL = true, isHeader=true, ExportPart=3 });
            trTable.Add(new TranlationTable() { id = 2, description = "CI", exportValue = null, fieldForValue = "G17", textToCheck = "CI:", fieldForCheck = "F17", year = null, isOK = false, isForTBL = true, isHeader = true, ExportPart = 3 });
            trTable.Add(new TranlationTable() { id = 3, description = "Fuel", exportValue = null, fieldForValue = "I17", textToCheck = "Fuel:", fieldForCheck = "H17", year = null, isOK = false, isForTBL = true, isHeader = true, ExportPart = 3 });
            trTable.Add(new TranlationTable() { id = 4, description = "Wage Economics", exportValue = null, fieldForValue = "E17", textToCheck = "Wage Economics:", fieldForCheck = "D17", year = null, isOK = false, isForTBL = true, isHeader = true, ExportPart = 3 });
            //trTable.Add(new TranlationTable() { id = 5, description = "New Bldg", exportValue = null, fieldForValue = "G17", textToCheck = "CI:", fieldForCheck = "F17", year = null, isOK = false, isForTBL = true, isHeader=true, ExportPart=3 });
            trTable.Add(new TranlationTable() { id = 6, description = "WACC", exportValue = null, fieldForValue = "X13", textToCheck = "WACC:", fieldForCheck = "W12", year = null, isOK = false, isForTBL = true, isHeader = true, ExportPart = 3 });
            trTable.Add(new TranlationTable() { id = 7, description = "Material Savings", exportValue = null, fieldForValue = "E19", textToCheck = "Wage Economics:", fieldForCheck = "D17", year = null, isOK = false, isForTBL = true, isHeader = true, ExportPart = 3 });
            trTable.Add(new TranlationTable() { id = 8, description = "Material Ecconomics", exportValue = null, fieldForValue = "E25", textToCheck = "Wage Economics:", fieldForCheck = "D17", year = null, isOK = false, isForTBL = true, isHeader = true, ExportPart = 3 });
            trTable.Add(new TranlationTable() { id = 9, description = "Major Capex Items", exportValue = null, fieldForValue = "E12", textToCheck = "Major Capex Items:", fieldForCheck = "C12", year = null, isOK = false, isForTBL = true, isHeader = true, ExportPart = 3 });
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

            catch (Exception)
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

            //Loop Translation Table and match agains sheet data
            foreach (TranlationTable lineObject in translatedTable)
            {
                //Check if sheet possition matches expected template positions.
                if (ProcitajNaSheetu(wbPart, item, lineObject.fieldForCheck).Trim() == lineObject.textToCheck)
                {
                    //Update table isOk variable and find value at sheet
                    lineObject.isOK = true;
                    lineObject.exportValue = ProcitajNaSheetu(wbPart, item, lineObject.fieldForValue);
                }
                else
                {
                    //In case expected data is not placed at correct possition
                    mainCheck = false;
                }

                //Add updated item to the result list
                exTable.Add(new ExportTable() { id = lineObject.id,
                                                description = lineObject.description,
                                                exportValue = lineObject.exportValue,
                                                year = lineObject.year,
                                                isOK = lineObject.isOK,
                                                isForTBL = lineObject.isForTBL,
                                                isHeader=lineObject.isHeader,
                                                ExportPart = lineObject.ExportPart
                });
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
    }    
}
