using SS;
using System;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;

namespace SpreadsheetGUI
{
    /// <summary>
    /// This class is to show the data graph of its corresponding spreadsheet
    /// </summary>
    public partial class DataGraph : Form
    {
        private Spreadsheet _spreadsheet;

        /// <summary>
        /// Initializes DataGraph form, and retrieve Spreadsheet object
        /// </summary>
        /// <param name="spreadsheet"></param>
        public DataGraph(Spreadsheet spreadsheet)
        {
            InitializeComponent();
            _spreadsheet = spreadsheet;
        }

        /// <summary>
        /// Loads dataGraphChart from the data of Spreadsheet object
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DataGraph_Load(object sender, EventArgs e)
        {
            List<string> list = _spreadsheet.GetNamesOfAllNonemptyCells().ToList();

            try
            {
                // create series' names from the values of B1, C1, D1
                dataGraphChart.Series.Add(_spreadsheet.GetCellValue("B1").ToString());
                if (list.Contains("C1"))
                    dataGraphChart.Series.Add(_spreadsheet.GetCellValue("C1").ToString());
                if (list.Contains("D1"))
                    dataGraphChart.Series.Add(_spreadsheet.GetCellValue("D1").ToString());
                // loop through Spreadsheet, and update the corresponding data to the series
                foreach (string name in list)
                {
                    if (name != "B1" && name != "C1" && name != "D1")
                    {
                        string xVal = _spreadsheet.GetCellValue("A" + name[1]).ToString();
                        string yVal = _spreadsheet.GetCellValue(name).ToString();

                        switch (name[0])
                        {
                            case 'B':
                                dataGraphChart.Series[0].Points.AddXY(xVal, yVal);
                                break;
                            case 'C':
                                dataGraphChart.Series[1].Points.AddXY(xVal, yVal);
                                break;
                            case 'D':
                                dataGraphChart.Series[2].Points.AddXY(xVal, yVal);
                                break;
                        }
                    }
                }
            }
            catch (Exception e1)
            {
                MessageBox.Show(e1.Message, "Error");
                this.Close();
            }
        }
    }
}
