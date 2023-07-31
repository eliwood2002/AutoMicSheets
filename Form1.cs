using AutoMicSheets.Properties;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace AutoMicSheets
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private string idCheck(string size, string sch)
        {
            string id;
            sch = sch.ToUpper();
            
            switch(size)
            {
                case "2":
                    switch (sch)
                    {
                        case "STD":
                            id = "2.067";
                            break;
                        case "XH":
                            id = "1.939";
                            break;
                        case "XXH":
                            id = "1.503";
                            break;
                        case "40":
                            id = "2.067";
                            break;
                        case "80":
                            id = "1.939";
                            break;
                        case "160":
                            id = "1.689";
                            break;
                        default:
                            id = "0";
                            break;
                    }
                    break;
                case "3":
                    switch (sch)
                    {
                        case "STD":
                            id = "3.068";
                            break;
                        case "XH":
                            id = "2.900";
                            break;
                        case "XXH":
                            id = "2.300";
                            break;
                        case "40":
                            id = "3.068";
                            break;
                        case "80":
                            id = "2.900";
                            break;
                        case "160":
                            id = ".957";
                            break;
                        default:
                            id = "0";
                            break;
                    }
                    break;
                case "4":
                    switch (sch)
                    {
                        case "STD":
                            id = "4.026";
                            break;
                        case "XH":
                            id = "3.826";
                            break;
                        case "XXH":
                            id = "3.152";
                            break;
                        case "40":
                            id = "4.026";
                            break;
                        case "80":
                            id = "3.826";
                            break;
                        case "120":
                            id = "3.624";
                            break;
                        case "160":
                            id = "3.438";
                            break;
                        default:
                            id = "0";
                            break;
                    }
                    break;
                case "6":
                    switch (sch)
                    {
                        case "STD":
                            id = "6.065";
                            break;
                        case "XH":
                            id = "5.761";
                            break;
                        case "XXH":
                            id = "4.897";
                            break;
                        case "40":
                            id = "6.065";
                            break;
                        case "80":
                            id = "5.761";
                            break;
                        case "120":
                            id = "5.501";
                            break;
                        case "160":
                            id = "5.189";
                            break;
                        default:
                            id = "0";
                            break;
                    }
                    break;
                case "8":
                    switch (sch)
                    {
                        case "STD":
                            id = "7.981";
                            break;
                        case "XH":
                            id = "7.625";
                            break;
                        case "XXH":
                            id = "6.875";
                            break;
                        case "20":
                            id = "8.125";
                            break;
                        case "30":
                            id = "8.071";
                            break;
                        case "40":
                            id = "7.981";
                            break;
                        case "60":
                            id = "7.813";
                            break;
                        case "80":
                            id = "7.625";
                            break;
                        case "100":
                            id = "7.439";
                            break;
                        case "120":
                            id = "7.189";
                            break;
                        case "140":
                            id = "7.001";
                            break;
                        case "160":
                            id = "6.813";
                            break;
                        default:
                            id = "0";
                            break;
                    }
                    break;
                case "10":
                    switch (sch)
                    {
                        case "STD":
                            id = "10.020";
                            break;
                        case "XH":
                            id = "9.750";
                            break;
                        case "20":
                            id = "10.250";
                            break;
                        case "30":
                            id = "10.136";
                            break;
                        case "40":
                            id = "10.020";
                            break;
                        case "60":
                            id = "9.750";
                            break;
                        case "80":
                            id = "9.564";
                            break;
                        case "100":
                            id = "9.314";
                            break;
                        case "120":
                            id = "9.064";
                            break;
                        case "140":
                            id = "8.750";
                            break;
                        case "160":
                            id = "8.500";
                            break;
                        default:
                            id = "0";
                            break;
                    }
                    break;
                case "12":
                    switch (sch)
                    {
                        case "STD":
                            id = "12.000";
                            break;
                        case "XH":
                            id = "11.750";
                            break;
                        case "20":
                            id = "12.250";
                            break;
                        case "30":
                            id = "12.090";
                            break;
                        case "40":
                            id = "11.938";
                            break;
                        case "60":
                            id = "11.626";
                            break;
                        case "80":
                            id = "11.376";
                            break;
                        case "100":
                            id = "11.064";
                            break;
                        case "120":
                            id = "10.750";
                            break;
                        case "140":
                            id = "10.500";
                            break;
                        case "160":
                            id = "10.126";
                            break;
                        default:
                            id = "0";
                            break;
                    }
                    break;
                case "14":
                    switch (sch)
                    {
                        case "STD":
                            id = "13.250";
                            break;
                        case "XH":
                            id = "13.000";
                            break;
                        case "10":
                            id = "13.500";
                            break;
                        case "20":
                            id = "13.376";
                            break;
                        case "30":
                            id = "13.250";
                            break;
                        case "40":
                            id = "13.124";
                            break;
                        case "60":
                            id = "12.814";
                            break;
                        case "80":
                            id = "12.500";
                            break;
                        case "100":
                            id = "12.126";
                            break;
                        case "120":
                            id = "11.814";
                            break;
                        case "140":
                            id = "11.500";
                            break;
                        case "160":
                            id = "11.188";
                            break;
                        default:
                            id = "0";
                            break;
                    }
                    break;
                case "16":
                    switch (sch)
                    {
                        case "STD":
                            id = "15.250";
                            break;
                        case "XH":
                            id = "15.000";
                            break;
                        case "10":
                            id = "15.500";
                            break;
                        case "20":
                            id = "15.376";
                            break;
                        case "30":
                            id = "15.250";
                            break;
                        case "40":
                            id = "15.000";
                            break;
                        case "60":
                            id = "14.688";
                            break;
                        case "80":
                            id = "14.314";
                            break;
                        case "100":
                            id = "13.938";
                            break;
                        case "120":
                            id = "13.564";
                            break;
                        case "140":
                            id = "13.124";
                            break;
                        case "160":
                            id = "12.814";
                            break;
                        default:
                            id = "0";
                            break;
                    }
                    break;
                case "18":
                    switch (sch)
                    {
                        case "STD":
                            id = "17.250";
                            break;
                        case "XH":
                            id = "17.000";
                            break;
                        case "10":
                            id = "17.500";
                            break;
                        case "20":
                            id = "17.376";
                            break;
                        case "30":
                            id = "17.124";
                            break;
                        case "40":
                            id = "16.876";
                            break;
                        case "60":
                            id = "16.500";
                            break;
                        case "80":
                            id = "16.126";
                            break;
                        case "100":
                            id = "15.688";
                            break;
                        case "120":
                            id = "15.250";
                            break;
                        case "140":
                            id = "14.876";
                            break;
                        case "160":
                            id = "14.438";
                            break;
                        default:
                            id = "0";
                            break;
                    }
                    break;
                case "20":
                    switch (sch)
                    {
                        case "STD":
                            id = "19.250";
                            break;
                        case "XH":
                            id = "19.000";
                            break;
                        case "10":
                            id = "19.500";
                            break;
                        case "20":
                            id = "19.250";
                            break;
                        case "30":
                            id = "19.000";
                            break;
                        case "40":
                            id = "18.814";
                            break;
                        case "60":
                            id = "18.376";
                            break;
                        case "80":
                            id = "17.938";
                            break;
                        case "100":
                            id = "17.438";
                            break;
                        case "120":
                            id = "17.000";
                            break;
                        case "140":
                            id = "16.500";
                            break;
                        case "160":
                            id = "16.064";
                            break;
                        default:
                            id = "0";
                            break;
                    }
                    break;
                default:
                    id = "0";
                    break;
            }
            return(id);

        }
        private void createButton_Click(object sender, EventArgs e)
        {
            //checking if the button is clicked
            string templateName;
            if (lsCheckBox.Checked)
            {
                templateName = "LSMicShtTemplate.xlsm";
            }
            else
            {
                templateName = "MicShtTemplate.xlsm";
            }
            //setting the file path for the template being used.
            string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, templateName);

            //setting id so it runs through the function once instead of for every mic sheet being made.
            string pID = idCheck(sizeTextBox.Text, schTextBox.Text);

            //declaring a count for the looping of each mic sheet
            int count = 1;
            while (count <= int.Parse(numTubesTextBox.Text)) {

                /*opening excel template to modify*/
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                Workbook wb;
                Worksheet ws;

                //opening the worksheet to modify
                wb = excel.Workbooks.Open(filePath);
                ws = wb.Worksheets[1];

                /*top section of mic sheet*/

                //job number cell
                Range jobNum = ws.Range["B12"];
                jobNum.Value = jobNumTextBox.Text;

                //Customer cell
                Range custName = ws.Range["B14"];
                if(lsCheckBox.Checked)
                {
                    custTextBox.Text = "L&S Proline";
                }
                custName.Value = custTextBox.Text.ToUpper();

                //PO cell
                Range poNum = ws.Range["B16"];
                poNum.Value = poTextBox.Text.ToUpper();

                //getting tube number and file name
                string tubeNum = string.Format("{0}-{1:D3}", jobNumTextBox.Text, count);

                //Tube Number cell
                Range tubeNumCell = ws.Range["G14"];
                tubeNumCell.Value = tubeNum;

                /*bottom section of mic sheet*/

                //Meter Manufacturer
                Range mfgName = ws.Range["B60"];
                mfgName.Value = mfgComboBox.Text;

                //Type of Meter
                Range meterTypeName = ws.Range["D60"];
                meterTypeName.Value = meterTypeComboBox.Text;

                //Orientation of meter
                Range orient = ws.Range["K59"];
                orient.Value = orientComboBox.Text;

                //ANSI of Meter
                Range ansi = ws.Range["I60"];
                ansi.Value = ansiComboBox.Text;

                /*setting size and schedule of pipe so you can also set Published ID*/

                //size of pipe 
                Range sizeNum = ws.Range["B61"];
                sizeNum.Value = sizeTextBox.Text;

                //Sch of pipe
                Range schNum = ws.Range["D61"];
                schNum.Value = schTextBox.Text;

                //Set the Fitting Published ID
                Range idNum = ws.Range["G61"];
                idNum.Value = pID;

                Range idNum2 = ws.Range["K64"];
                idNum2.Value = pID;

                //TODO: Set to where you want to save.
                string yearFolder = jobNumTextBox.Text.Substring(0, 2);

                string savePath = "S:\\Circle B\\Shop Users\\Calibration group\\20" + yearFolder +" Job Mics\\"+ jobNumTextBox.Text + "\\";

                if (!Directory.Exists(savePath))
                {
                    Directory.CreateDirectory(savePath);
                }
                ws.Protect("cbmf", Contents: true);
                wb.SaveAs(savePath + tubeNum + " mic sht.xlsm");
                wb.Close();

                count++;
            }

            this.Close();
        }
    }
}

