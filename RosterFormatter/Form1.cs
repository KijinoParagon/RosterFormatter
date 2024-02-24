
using System.Diagnostics;
using System.Text;
using Microsoft.WindowsAPICodePack.Dialogs;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using Microsoft.Office.Interop.Excel;


namespace RosterFormatter
{
    public partial class Form1 : Form
    {
        FileStream prefs;
        string outputLoc;
        string inputLoc;
        DateTime today;
        string dateString;
        public Form1()
        {
            today = DateTime.Now;
            dateString += today.ToString("MM/dd/YY");
            //See if the prefs file exists
            //If it does, get the preferred output location
            if (File.Exists("prefs.conf"))
            {
                prefs = File.Open("prefs.conf", FileMode.Open);
                outputLoc = "";
                byte[] b = new byte[1024];
                UTF8Encoding temp = new UTF8Encoding(true);
                while (prefs.Read(b, 0, b.Length) > 0)
                {
                    Console.WriteLine(temp.GetString(b));
                    outputLoc = temp.GetString(b);
                }
                prefs.Close();

            }
            //Otherwise, create the file so that we can store that location
            else
            {
                File.Create("prefs.conf");
            }
            InitializeComponent();
        }

        private string OpenFile(string path)
        {
                OpenFileDialog dialog = new OpenFileDialog();

            dialog.Filter = "csv files (*.csv)|*.csv|All files (*.*)|*.*";

            dialog.InitialDirectory = "C:\\Users\\";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                return dialog.FileName;
            }

            return null;

        }

        private string SaveFile(string path)
        {
            SaveFileDialog dialog = new SaveFileDialog();

            dialog.Filter = "xlsx files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            dialog.AddExtension = true;
            dialog.DefaultExt = ".xlsx";
            dialog.FileName = today.ToString("MM-dd-yyyy");

            dialog.InitialDirectory = "C:\\Users\\";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                return dialog.FileName;
            }

            return null;

        }
        private void btn_open_Click(object sender, EventArgs e)
        {
            inputLoc = OpenFile(outputLoc);
            txt_input.Text = inputLoc;
        }

        private void btn_output_Click(object sender, EventArgs e)
        {
            outputLoc = SaveFile(outputLoc);
            txt_output.Text = outputLoc;
        }
        

        

       

        private void btn_format_Click(object sender, EventArgs e)
        {
            //Create Excel workbook
            if(!File.Exists(outputLoc))
            {
                Excel.Application oXL;
                Excel._Workbook oWB;
                Excel._Worksheet oSheet;
                Excel.Range oRng;

                try
                {
                    //Start Excel and get Application object.
                    oXL = new Excel.Application();
                    oXL.Visible = true;

                    //Get a new workbook.
                    //oWB = (Excel._Workbook)(oXL.Workbooks.Add(Missing.Value));
                    oWB = (Excel._Workbook)(oXL.Workbooks.Open(inputLoc));
                    oSheet = (Excel._Worksheet)oWB.ActiveSheet;


                    /*
                    //Add table headers going cell by cell.
                    oSheet.Cells[1, 1] = "First Name";
                    oSheet.Cells[1, 2] = "Last Name";
                    oSheet.Cells[1, 3] = "Full Name";
                    oSheet.Cells[1, 4] = "Salary";

                    //Format A1:D1 as bold, vertical alignment = center.
                    oSheet.get_Range("A1", "D1").Font.Bold = true;
                    oSheet.get_Range("A1", "D1").VerticalAlignment =
                    Excel.XlVAlign.xlVAlignCenter;

                    

                    // Create an array to multiple values at once.
                    string[,] saNames = new string[5, 2];

                    saNames[0, 0] = "John";
                    saNames[0, 1] = "Smith";
                    saNames[1, 0] = "Tom";
                    saNames[1, 1] = "Brown";
                    saNames[2, 0] = "Sue";
                    saNames[2, 1] = "Thomas";
                    saNames[3, 0] = "Jane";
                    saNames[3, 1] = "Jones";
                    saNames[4, 0] = "Adam";
                    saNames[4, 1] = "Johnson";

                    //Fill A2:B6 with an array of values (First and Last Names).
                    oSheet.get_Range("A2", "B6").Value2 = saNames;



                    //AutoFit columns A:D.
                    oRng = oSheet.get_Range("A1", "D1");
                    oRng.EntireColumn.AutoFit();

                    */
                    //Make sure Excel is visible and give the user control
                    //of Microsoft Excel's lifetime.
                    oSheet.SaveAs2(outputLoc, XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, ReadOnlyRecommended: false);
                    oSheet.get_Range("I1", "J1").EntireColumn.Delete();
                    oSheet.get_Range("F1", "F2").EntireColumn.Delete();
                    oRng = oSheet.get_Range("A1", "G1");
                    oRng.EntireColumn.Font.Size = 12;
                    oRng.EntireColumn.AutoFit();
                    //oRng.EntireColumn.HorizontalAlignment = HorizontalAlignment.Left;
                    oRng = oSheet.get_Range("A1", "H1");
                    oRng.EntireColumn.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                    oSheet.get_Range("A1", "A2").EntireColumn.ColumnWidth += 5;
                    oSheet.get_Range("B1", "B2").EntireColumn.ColumnWidth += 5;
                    oSheet.get_Range("C1", "C2").EntireColumn.ColumnWidth += 5;
                    oSheet.get_Range("D1", "D2").EntireColumn.ColumnWidth += 5;
                    oSheet.get_Range("E1", "E2").EntireColumn.ColumnWidth += 5;
                    oSheet.get_Range("F1", "F2").EntireColumn.ColumnWidth += 5;
                    oSheet.get_Range("G1", "G2").EntireColumn.ColumnWidth += 5;
                    oSheet.get_Range("H1", "H2").EntireColumn.ColumnWidth += 5;

                    /*
                    oSheet.get_Range("A1", "A2").EntireColumn.HorizontalAlignment = HorizontalAlignment.Left;
                    oSheet.get_Range("B1", "B2").EntireColumn.HorizontalAlignment = HorizontalAlignment.Left;
                    oSheet.get_Range("C1", "C2").EntireColumn.HorizontalAlignment = HorizontalAlignment.Left;
                    oSheet.get_Range("D1", "D2").EntireColumn.HorizontalAlignment = HorizontalAlignment.Left;
                    oSheet.get_Range("E1", "E2").EntireColumn.HorizontalAlignment = HorizontalAlignment.Left;
                    oSheet.get_Range("F1", "F2").EntireColumn.HorizontalAlignment = HorizontalAlignment.Left;
                    oSheet.get_Range("G1", "G2").EntireColumn.HorizontalAlignment = HorizontalAlignment.Left;
                    oSheet.get_Range("H1", "H2").EntireColumn.HorizontalAlignment = HorizontalAlignment.Left;
                    */
                    oSheet.Cells[1, 8] = " ";

                    oXL.Visible = true;
                    oSheet.PageSetup.BottomMargin = 3;
                    oSheet.PageSetup.TopMargin = 3;
                    oSheet.PageSetup.LeftMargin = 3;
                    oSheet.PageSetup.RightMargin = 3;

                    oSheet.PageSetup.Orientation = XlPageOrientation.xlLandscape;
                    oSheet.PageSetup.Zoom = false;
                    oSheet.PageSetup.FitToPagesWide = 1;
                    //oSheet.SaveAs2(outputLoc, XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, ReadOnlyRecommended: false);
                    oWB.Save();
                    oWB.PrintPreview();
                    oXL.UserControl = true;
                }
                catch (Exception theException)
                {
                    String errorMessage;
                    errorMessage = "Error: ";
                    errorMessage = String.Concat(errorMessage, theException.Message);
                    errorMessage = String.Concat(errorMessage, " Line: ");
                    errorMessage = String.Concat(errorMessage, theException.Source);

                    MessageBox.Show(errorMessage, "Error");
                }

            }
        }
    }
}
