using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

namespace PhoneNumberSorter
{
    public partial class Form1 : Form
    {
        //Create Variables
        string DELETABLE_FILE_NAME;
        string COMPARABLE_FILE_NAME;

        public Form1()
        {
            InitializeComponent();

            //Adds directions to lblDirections at the top of the app when app opens
            lblDirections.Text = "This is a program that sorts through and compares two given Excel sheet " +
                "files that each contain a column list of phone numbers. Below, click each 'Browse' button to " +
                "upload a file to the corresponding box. The first box being the file that you want to delete from, " +
                "and the second box being the file that is for comparison purposes only. \n \nAfter your files are " +
                "selected, you can input an area code for any phone numbers (including their rows) you would like to " +
                "keep; these will not be deleted from the first sheet. You may also leave the textbox empty if this " +
                "option is not needed. Keep in mind that any area code entered must be three-digits long. Letters and " +
                "two-digit numbers will not be accepted. \n \nWhen 'Parse' is clicked, the program will search for any " +
                "phone number differences between the two given Excel files and delete those specific numbers from the " +
                "first given sheet as long as it does not begin with the area code provided. If an area code is not given, " +
                "then all differening phone numbers will be removed from the first sheet despite area code. This new Excel " +
                "sheet can then be stored to your computer via a popup. \n \nTo exit, either click the 'x' in the upper " +
                "right corner or click on 'exit'.";
        }

        /// <summary>
        /// Allows user to choose an excel file to delete numbers from
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnDelete_Click(object sender, EventArgs e)
        {
            GetFileName(true);
        }

        /// <summary>
        /// Allows user to choose an excel file to compare the first list to
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnCompare_Click(object sender, EventArgs e)
        {
            GetFileName(false);
        }

        /// <summary>
        /// Sets tbDelete and tbCompare text; saves corresponding file names to
        /// variables
        /// </summary>
        /// <param name="sheetToggle"> Represents which list to save and display
        /// file information for; true for DELETABLE_LIST (first given file) and false for
        /// COMPARABLE_LIST (second given file)</param>
        private void GetFileName(bool sheetToggle)
        {
            //Open file dialog
            using (OpenFileDialog openFile = new OpenFileDialog())
            {
                openFile.InitialDirectory = "c:\\"; 
                openFile.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm"; // Restricts user to uploading only excel files
                openFile.RestoreDirectory = true;

                if (openFile.ShowDialog() == DialogResult.OK)
                {
                    // First display file path in tbDelete for user to see
                    if (sheetToggle)
                    {
                        tbDelete.Text = openFile.SafeFileName;

                        //Set DELETABLE_FILE_NAME
                        DELETABLE_FILE_NAME = openFile.FileName;
                    }
                    else
                    {
                        tbCompare.Text = openFile.SafeFileName;

                        //Set DELETABLE_FILE_NAME
                        COMPARABLE_FILE_NAME = openFile.FileName;
                    }
                }
            }
        }

        /// <summary>
        /// Compares the first list to the second and deletes any differing numbers
        /// from list 1 while retaining any numbers beginning with the given area code. 
        /// Produces a new list in a .txt file for the user to save
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnParse_Click(object sender, EventArgs e)
        {
            // Save contents of tbAreaCode
           string areaCode = tbAreaCode.Text; //A string so that each number can be accessed by index

            //Make sure there are files selected
            if (!String.IsNullOrEmpty(tbDelete.Text) && !String.IsNullOrEmpty(tbCompare.Text))
            {
                //Variables
                int rowCount; //Excel rows
                int columnCount; // Excel columns

                //Pull data from second given file
                List<long> comparableList = SheetToArray();

                //Separate list data
                //string[] deletableLines = SHEET_CONTENTS_DELETE.Split('\n');
                //string[] comparableLines = SHEET_CONTENTS_COMPARE.Split('\n');

                // Store line data (numbers) into arrays
                /* LineDataToArray(DELETEABLE_LIST, deletableLines);
                 LineDataToArray(COMPARABLE_LIST, comparableLines); */

                 // Make sure area code is either three digits or empty, returns true
                 //if area code is acceptable and false if not
                 if(CheckAreaCode(areaCode))
                 {
                    //Initialize Excel application objects
                    /*Excel.Application xlApp;
                    Excel.Workbook xlWorkbook;
                    Excel.Worksheet xlWorksheet;
                    Excel.Range xlRange;

                    // Retrieve file contents
                    xlApp = new Excel.Application();
                    xlWorkbook = xlApp.Workbooks.Open(DELETABLE_FILE_NAME, 0, false);
                    xlWorksheet = xlWorkbook.Worksheets[1]; //Starts at 1 for Excel sheets
                    xlRange = xlWorksheet.UsedRange; //Holds the int range of all available columns/rows in document
                    rowCount = xlRange.Rows.Count;
                    columnCount = xlRange.Columns.Count;

                    //Compare numbers from first given list to seconf given list(comparableList)
                    for (int r = 4; r <= rowCount; r++)
                    {
                        long number = Convert.ToInt64(xlRange.Cells[r, 1].Value2);
                        if (!comparableList.Contains(number) &&
                            number / 10000000 != Convert.ToInt16(areaCode))
                        {
                            ((Range)xlWorksheet.Rows[r]).Delete(XlDeleteShiftDirection.xlShiftUp);
                            xlWorkbook.Save();
                        }
                    }

                    MessageBox.Show("Done!");*/

                    //CompareListsWithAreaCode(DELETEABLE_LIST, COMPARABLE_LIST, Convert.ToInt16(areaCode));

                    //Provide user with a file to save
                    //SaveNewFile(DELETEABLE_LIST);

                     //Clear textboxes
                     //ClearTextBoxes();
                 }
                 /*else
                 {
                     if(String.IsNullOrWhiteSpace(areaCode))
                     {
                         CompareListsWithNoAreaCode(DELETEABLE_LIST, COMPARABLE_LIST);

                         //Provide user with a file to save
                         SaveNewFile(DELETEABLE_LIST);

                         //Clear textboxes
                         ClearTextBoxes();
                     }
                 } */
            }
            else
            {
                MessageBox.Show("Please make sure two files are selected",
                    "Oh, no!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        /// <summary>
        /// Stores the given SECOND (Comparable) Excel Sheet phone numbers into 
        /// an array
        /// </summary>
        private List<long> SheetToArray()
        {
            //Variables       
            List<long> comparableList = new List<long>();
            int rowCount; //Excel rows

            //Initialize Excel application objects
            /*Excel.Application xlApp;
            Excel.Workbook xlWorkbook;
            Excel.Worksheet xlWorksheet;
            Excel.Range xlRange;   
            
            // Retrieve file contents
            xlApp = new Excel.Application();
            xlWorkbook = xlApp.Workbooks.Open(COMPARABLE_FILE_NAME);
            xlWorksheet = xlWorkbook.Worksheets[1]; //Starts at 1 for Excel sheets
            xlRange = xlWorksheet.UsedRange; //Holds the int range of all available columns/rows in document
            rowCount = xlRange.Rows.Count;

            //Look at each row and delete by number if does not meet requirements
            for (int r = 4; r <= rowCount; r++) //Starting at 4 to move past infornational rows; should write code to look for automatically; error with COMPARABLE
            {
                long number = Convert.ToInt64(xlRange.Cells[r, 1].Value2); //Only looking at first column which contains the phone number
                comparableList.Add(number);
            }
                    
            //De-initialize Excel objects
            xlWorkbook.Close(true, null, null);
            xlApp.Quit();
            Marshal.ReleaseComObject(xlWorksheet);
            Marshal.ReleaseComObject(xlWorkbook);
            Marshal.ReleaseComObject(xlApp); */

            //Return list
            return comparableList;
        }
        
        /// <summary>
        /// Makes sure area code is only used when it is valid: three numeric characters
        /// or nothing at all
        /// </summary>
        /// <param name="areaCode"></param>
        /// <returns></returns>
        private static bool CheckAreaCode(string areaCode)
        {
            //Compare lists and delete differing numbers from first list, saving numbers
            //with given area code
            if (areaCode.Length == 3) // Making sure there is a 3-character numeric area code to work with
            {
                if (LookForNumbers(areaCode))
                {
                    return true;
                }
                else // If area code ISN'T all numbers, give a warning
                {
                    MessageBox.Show("Only numeric area codes are accepted",
                                            "Oh, no!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
            else
            {
                //If the code entered IS NEITHER 3 characters nor 0, give a warning
                if (areaCode.Length == 2 || areaCode.Length == 1)
                {
                    MessageBox.Show("Any area code entered must be a total of three digits",
                        "Oh, no!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                }
            }

            return false;
        }

        /// <summary>
        /// Checks a string of three characters to make sure they are all
        /// numeric values
        /// </summary>
        /// <param name="areaCode"></param>
        /// <returns>True if every character is a number</returns>
        private static bool LookForNumbers(string areaCode)
        {
            //Variables
            bool index0isNumber = false;
            bool index1isNumber = false;
            bool index2isNumber = false;

            for(int i = 0; i < 3; i++)
            {
                if(areaCode[i] == '1' ||
                    areaCode[i] == '2' ||
                    areaCode[i] == '3' ||
                    areaCode[i] == '4' ||
                    areaCode[i] == '5' || 
                    areaCode[i] == '6' ||
                    areaCode[i] == '7' ||
                    areaCode[i] == '8' ||
                    areaCode[i] == '9' ||
                    areaCode[i] == '0')
                {
                    if(i == 0)
                    {
                        index0isNumber = true;
                    }
                    else if(i == 1)
                    {
                        index1isNumber = true;
                    }
                    else
                    {
                        index2isNumber = true;
                    }
                }
            }

            return index0isNumber && index1isNumber && index2isNumber;
        }

        /// <summary>
        /// Compares the deletableList to the comparableList and removes
        /// differing numbers that DO NOT beging with the given area code
        /// </summary>
        /// <param name="deletableList"></param>
        /// <param name="comparableList"></param>
        /// <param name="areaCode"></param>
        private static void CompareListsWithAreaCode(List<long> deletableList, List<long> comparableList, int areaCode)
        {
            for (int i = deletableList.Count - 1; i >= 0; i--)
            {
                //If deletable list has a number different than comparable                
                if (!(comparableList.Contains(deletableList[i])))
                {
                    // and DOES NOT start with the area code, delete it
                    if (deletableList[i] / 10000000 != Convert.ToInt16(areaCode))
                    {
                        deletableList.Remove(deletableList[i]);
                    }
                }
            }
        }

        /// <summary>
        /// Compares the deletableList to the comparableList and removes
        /// differing numbers (no area code involved)
        /// </summary>
        /// <param name="deletableList"></param>
        /// <param name="comparableList"></param>
        private static void CompareListsWithNoAreaCode(List<long> deletableList, List<long> comparableList)
        {
            for (int i = deletableList.Count - 1; i >= 0; i--)
            {              
                if (!(comparableList.Contains(deletableList[i])))
                {
                    deletableList.Remove(deletableList[i]);
                }
            }
        }

        /// <summary>
        /// Pops open a dialog for the user to save a newly written
        /// file of numbers to their computer
        /// </summary>
        /// <param name="deletableList"></param>
        private static void SaveNewFile(List<long> deletableList)
        {
            SaveFileDialog saveFile = new SaveFileDialog();

            saveFile.FileName = "newListOfNumbers.txt";
            saveFile.Filter = "txt files|*.txt";

            if (saveFile.ShowDialog() == DialogResult.OK)
            {
                StreamWriter streamWriter = new StreamWriter(saveFile.OpenFile());

                for (int i = 0; i < deletableList.Count; i++)
                {
                    streamWriter.WriteLine(deletableList[i]);
                }

                streamWriter.Dispose();
                streamWriter.Close();
            }
        }

        /// <summary>
        /// Removes all text from textboxes
        /// </summary>
        private void ClearTextBoxes()
        {
            tbDelete.Clear();
            tbCompare.Clear();
            tbAreaCode.Clear();
        }
        
        /// <summary>
        /// Allows user to exit the application
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnExit_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
