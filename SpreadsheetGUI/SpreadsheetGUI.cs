using SpreadsheetUtilities;
using SS;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace SpreadsheetGUI
{
    /// <summary>
    /// This class creates a GUI windows form for Spreadsheet
    /// </summary>
    public partial class SpreadsheetGUI : Form
    {
        private Spreadsheet spreadsheet; // the core Spreadsheet object
        private int _row, _col; // row and col of the spreadsheet panel
        private string currentOpenFileName; // keep track of the current open file

        /// <summary>
        /// Initializes Spreadsheet object, currentOpenFileName and registers the needed methods to SelectionChanged Delegate
        /// </summary>
        public SpreadsheetGUI()
        {
            InitializeComponent();
            spreadsheet = new Spreadsheet(IsValid, s => s.ToUpper(), "ps6");
            currentOpenFileName = string.Empty;
            spreadsheetPanel1.SelectionChanged += DisplayCellNameContentsValue;
        }


        private void SpreadsheetGUI_Load(object sender, EventArgs e)
        {
            contentsTextBox.Select(); // the contentsTextBox gets focus when the form loads
            DisplayCellNameContentsValue(spreadsheetPanel1); // to display the default cell name in cellNameTextBox
        }

        /// <summary>
        /// Puts cell name, contents, and value to the textboxes.
        /// </summary>
        /// <param name="ssp"></param>
        private void DisplayCellNameContentsValue(SpreadsheetPanel ssp)
        {
            ssp.GetSelection(out _col, out _row); // retrieve the col# and row# from mouse click
            string cellName = ConvertToCellName(_row, _col); // convert them to the cellname
            cellNameTextBox.Text = cellName;
            contentsTextBox.Text = spreadsheet.GetCellContents(cellName).ToString();
            valueTextBox.Text = ReturnCellValue(cellName);
        }

        /// <summary>
        /// Handles New menu
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void NewToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Tell the application context to run the form on the same
            // thread as the other forms.
            SpreadsheetApplicationContext.getAppContext().RunForm(new SpreadsheetGUI());
        }

        /// <summary>
        /// Handles changeButton, when clicking this button, valueTextBox is updated
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ChangeButton_Click(object sender, EventArgs e)
        {
            UpdateValue();
        }

        /// <summary>
        /// When pressing Enter in ContentsTextBox, valueTextBox is updated
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ContentsTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                UpdateValue();
                // gets rid of Ding sound
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }

        /// <summary>
        /// Handles Close menu
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CloseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        /// <summary>
        /// Before form is closed, prompt user to save if data is modified
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SpreadsheetGUI_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (AskSave() == DialogResult.Yes) // prompt user to save if data is modified
            {
                SaveSS(); // show save dialog
                // if save dialog is closed without saving, closing form is canceled
                if (spreadsheet.Changed == true) e.Cancel = true;
            }
        }

        /// <summary>
        /// Handles Load menu
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void LoadToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (AskSave() == DialogResult.Yes) // prompt user to save if data is modified
            {
                SaveSS();
            }
            else
            {
                var fileName = string.Empty;

                // create a file dialog
                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.Filter = "sprd files (*.sprd)|*.sprd|All files (*.*)|*.*"; // indicates file types
                    openFileDialog.FilterIndex = 1; // set the default file type to .sprd
                    openFileDialog.RestoreDirectory = true; // open the previous directory
                    openFileDialog.Multiselect = false; // disallow multiselect
                    openFileDialog.AddExtension = true; // auto-add .sprd
                    if (openFileDialog.ShowDialog() == DialogResult.OK) // when clicking Load
                    {
                        // Get the path of specified file
                        fileName = openFileDialog.FileName;
                    }
                }
                if (fileName == string.Empty) return;
                try
                {
                    spreadsheetPanel1.Clear(); // clear all the values on the spreadsheet panel before loading
                    currentOpenFileName = fileName;
                    spreadsheet = new Spreadsheet(fileName, IsValid, s => s.ToUpper(), "ps6");
                    foreach (string name in spreadsheet.GetNamesOfAllNonemptyCells())
                    {
                        FillValueToCell(name); // fill the panel with new values
                    }
                }
                catch (SpreadsheetReadWriteException e1)
                {
                    MessageBox.Show(e1.Message, "Error");
                }
            }
        }

        /// <summary>
        /// Handles Save menu
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SaveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // if the current spreadsheet hasn't been saved to an existing file,
            // prompt user to save; otherwise, save directly to the existing file
            if (currentOpenFileName == string.Empty)
                SaveSS();
            else
                spreadsheet.Save(currentOpenFileName);
        }

        /// <summary>
        /// Handles Save as menu, this always prompt user to save file
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SaveAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveSS();
        }

        /// <summary>
        /// Handles Generate Data Graph menu. This generates 3 series based on 3 columns B1, C1, D1.
        /// The cells' values below B1, C1, D1 must be numbers.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void GenerateGraphToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // check the cells' values below B1, C1, D1 whether they are numbers
            foreach (string name in spreadsheet.GetNamesOfAllNonemptyCells())
            {
                if ((name[0] == 'B' || name[0] == 'C' || name[0] == 'D')
                    && name != "B1" && name != "C1" && name != "D1")
                {
                    double num;
                    if (!double.TryParse(ReturnCellValue(name), out num))
                    {
                        MessageBox.Show("The table's format is illegal. Please refer to Help.", "Error");
                        return;
                    }
                }
            }
            // open DataGraph form to show the graph, and pass Spreadsheet object to DataGraph form
            DataGraph datagraph = new DataGraph(spreadsheet);
            datagraph.Show();
        }

        /// <summary>
        /// Converts row# and col# to the corresponding cell name
        /// </summary>
        /// <param></param>
        /// <returns></returns>
        private string ConvertToCellName(int row, int col)
        {
            row++;
            int dividend = col + 1;
            int modulo = (dividend - 1) % 26;
            string cellName = Convert.ToChar(65 + modulo).ToString() + row.ToString();
            return cellName;
        }

        /// <summary>
        /// Checks if a variable starts with a character and ends with a number from 1-99
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        private bool IsValid(string s)
        {
            String varPattern = @"^[a-zA-Z][1-9][0-9]?$";
            if (!Regex.IsMatch(s, varPattern)) return false;
            else return true;
        }

        /// <summary>
        /// Fills value in a cell
        /// </summary>
        /// <param name="name"></param>
        private void FillValueToCell(string name)
        {
            int col = char.ToUpper(name[0]) - 'A'; // convert col# to its corresponding character
            int row = name[1] - 49;
            string value = ReturnCellValue(name); // retrieve value from Spreadsheet object
            spreadsheetPanel1.SetValue(col, row, value); // update to spreadsheet panel
        }

        /// <summary>
        /// Prompts user to save if the current spreadsheet is modified
        /// </summary>
        private DialogResult AskSave()
        {
            if (spreadsheet.Changed == true)
            {
                MessageBoxButtons buttons = MessageBoxButtons.YesNo;
                DialogResult result = MessageBox.Show("Data has been modified. Would you like to save?", "Save?", buttons);
                return result;
            }
            return DialogResult.No;
        }


        /// <summary>
        /// This methods show a save dialog for user to save the file
        /// </summary>
        private void SaveSS()
        {
            var fileName = string.Empty;
            // create a SaveFileDialog object
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "sprd files (*.sprd)|*.sprd|All files (*.*)|*.*";
                saveFileDialog.FilterIndex = 1;
                saveFileDialog.RestoreDirectory = true;
                saveFileDialog.AddExtension = true;
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    // Get the path of specified file
                    fileName = saveFileDialog.FileName;
                }
            }
            if (fileName == string.Empty) return;
            try
            {
                spreadsheet.Save(fileName); // save the core spreadsheet object to the file
                currentOpenFileName = fileName;
            }
            catch (SpreadsheetReadWriteException e1)
            {
                MessageBox.Show(e1.Message, "Error");
            }
        }

        /// <summary>
        /// Retrieves the value from a cell from the spreadsheet object
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        private string ReturnCellValue(string name)
        {
            if (spreadsheet.GetCellValue(name) is FormulaError) // value can be a FormulaError
            {
                FormulaError error = (FormulaError)spreadsheet.GetCellValue(name);
                return error.Reason;
            }
            return spreadsheet.GetCellValue(name).ToString();
        }

        /// <summary>
        /// Handles Help menu
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void HelpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string help = string.Format("To update the value of a cell, select a cell, enter the contents, " +
                "and press Enter or click Change.{0}" +
                "To create a new spreadsheet, go to File/New{0}" +
                "To save the current spreadsheet, go to File/Save{0}" +
                "To save the current spreadsheet with a different name, go to File/Save as...{0}" +
                "Additional feature: " +
                "Generating a data graph for the current spreadsheet. " +
                "To generate, go to Tools/Generate data graph{0}" +
                "Constraints: this feature can generate only 3 series based on the columns B, C, and D. " +
                "The series' names are the values of B1, C1, and D1. " +
                "The values of the cells below B1, C1, D1 must be numbers.", Environment.NewLine);
            MessageBox.Show(help, "Help", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>
        /// Updates value of the current cell, and updates the values of other dependent cells
        /// </summary>
        private void UpdateValue()
        {
            try
            {
                IList<string> list = new List<string>();
                string cellName = ConvertToCellName(_row, _col);
                list = spreadsheet.SetContentsOfCell(cellName, contentsTextBox.Text);
                valueTextBox.Text = ReturnCellValue(cellName);
                foreach (string name in list)
                {
                    FillValueToCell(name);
                }
            }
            catch (FormulaFormatException e1)
            {
                MessageBox.Show(e1.Message, "Error");
            }
            catch (CircularException e2)
            {
                MessageBox.Show(e2.Message, "Error");
            }
        }
    }
}
