using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using OfficeOpenXml;         // This is for EPPlus - Creating spreadsheets
using OfficeOpenXml.Style;   // This is for EPPlus - Creating spreadsheets
using ADODB;

namespace WriteToXLSX
{
    public partial class Form1 : Form
    {
        string strXLSXfile = "";
        string strDataFile = "";
        bool bComplete = false;
        private ADODB.Recordset rsCellData = new ADODB.Recordset();

        public Form1()
        {
            InitializeComponent();

            /* Get the arguments */

            string[] args = Environment.GetCommandLineArgs();
            if (args.Length != 2)
            {
                MessageBox.Show("No command line argument provided. " + args[0]);
                return;
            }

            /* Break them up since we know the expected format */

            string[] argsSplit = args[1].Split('|');
            if (argsSplit.Length != 2)
            {
                MessageBox.Show("Not all arguments were provided.");
                return;
            }

            /* Get the input XLSX file and make sure it is good */

            strXLSXfile = argsSplit[0].Trim();
            if (strXLSXfile == "")
            {
                MessageBox.Show("No input XLSX file provided.");
                return;
            }
            if (File.Exists(strXLSXfile) == false)
            {
                MessageBox.Show("Target file provided could not be found." + Environment.NewLine + strXLSXfile);
                return;
            }

            /* Get the output file and make sure it is good */

            strDataFile = argsSplit[1].Trim();
            if (strDataFile == "")
            {
                MessageBox.Show("No output XLSX file provided.");
                return;
            }
            if (File.Exists(strDataFile) == false)
            {
                MessageBox.Show("Data file provided doesn't exist." + Environment.NewLine + strDataFile);
                File.Delete(strDataFile);
            }

            WriteData();

        }

        private void WriteData()
        {

            /* Read up the data to be written to the spreadsheet from the text file */
            
            rsCellData.Fields.Append("Cell", DataTypeEnum.adVarChar, 10);
            rsCellData.Fields.Append("Value", DataTypeEnum.adVarChar, 512);
            rsCellData.Open();

            string line;
            int counter = 0;

            System.IO.StreamReader file = new System.IO.StreamReader(strDataFile);
            while ((line = file.ReadLine()) != null)
            {
                if (line == "")
                {
                    MessageBox.Show("Bad data format at line " + counter.ToString() + " No data on this line. Exiting.");
                    return;
                }
                counter++;

                int iIndex = line.IndexOf('|');
                if (iIndex == -1)
                {
                    MessageBox.Show("Bad data format at line " + counter.ToString() + " No Pipe delimiter found in line. Exiting.");
                    return;
                }

                rsCellData.AddNew();
                rsCellData.Fields["Cell"].Value = line.Substring(0, iIndex);
                iIndex++;
                rsCellData.Fields["Value"].Value = line.Substring(iIndex, line.Length - iIndex);
                rsCellData.Update();
            }

            if (rsCellData.RecordCount > 0)
            {
                rsCellData.MoveFirst();
            }

            /* Open and write the XlSX file. */
            
            FileInfo TargetFile = new FileInfo(strXLSXfile);
            
            try
            {
                using (ExcelPackage package = new ExcelPackage(TargetFile))
                {

                    /* Get the work book in the file */

                    ExcelWorkbook workBook = package.Workbook;
                    if (workBook != null)
                    {
                        if (workBook.Worksheets.Count < 0)
                        {
                            MessageBox.Show("There are no worksheets in the workbook.");
                            package.Dispose();
                            workBook.Dispose();
                            return;
                        }

                        /* Get the first worksheet */

                        //ExcelWorksheet Worksheet = workBook.Worksheets.First();
                        var worksheet = package.Workbook.Worksheets[1];

                        /* Loop throug the recordset and write the data to the worksheet */
                        string strCell = "";
                        for (; !rsCellData.EOF; rsCellData.MoveNext())
                        {
                            strCell = rsCellData.Fields["Cell"].Value;
                            worksheet.Cells[strCell].Value = rsCellData.Fields["Value"].Value;
                        }
                        worksheet.Cells["C13:C47"].Style.WrapText = true;
                    }
                    /* Save the workbook */
                    try
                    {
                        package.Save();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error saving the ATT West North Regional JPA Invoice report." + Environment.NewLine + ex);
                        return;
                    }
                    package.Dispose();
                }
            }
            catch (Exception Ex)
            {
                MessageBox.Show("Error opening spreadsheet. Is it already open? Close it and try again." + Environment.NewLine + Ex.Message);
                return;
            }
            rsCellData.Close();
            bComplete = true;
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            if (bComplete == true)
            {
                Application.Exit();
            }
        }
    }
}
