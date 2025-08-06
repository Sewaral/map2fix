using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
namespace sewartesting
{
    using Excel = Microsoft.Office.Interop.Excel;
    public partial class loginstudent : Form
    {
        public loginstudent()
        {
            InitializeComponent();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            signupasstudent signupForm = new signupasstudent();
            signupForm.Show();
        }

        private void loginstudent_Load(object sender, EventArgs e)
        {

        }
        private void button1_Click(object sender, EventArgs e)
        {
            string username = maskedTextBox1.Text.Trim();
            string password = maskedTextBox2.Text.Trim();

            if (string.IsNullOrEmpty(username) || string.IsNullOrEmpty(password))
            {
                MessageBox.Show("Please enter both username and password.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string solutionPath = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName;
            string folderPath = Path.Combine(solutionPath, "data");
            string signupPath = Path.Combine(folderPath, "signup1.xlsx");
            string loginLogPath = Path.Combine(folderPath, "logasstudent.xlsx");

            if (!File.Exists(signupPath))
            {
                MessageBox.Show("No registered users found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;

            try
            {
                excelApp = new Excel.Application();
                workbook = excelApp.Workbooks.Open(signupPath);
                worksheet = (Excel.Worksheet)workbook.Sheets[1];

                int rowCount = worksheet.UsedRange.Rows.Count;
                bool loginSuccessful = false;

                for (int row = 2; row <= rowCount; row++)
                {
                    string storedUsername = worksheet.Cells[row, 1].Value2?.ToString();
                    string storedPassword = worksheet.Cells[row, 3].Value2?.ToString();

                    if (storedUsername == username && storedPassword == password)
                    {
                        loginSuccessful = true;
                        break;
                    }
                }

                workbook.Close(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);

                if (loginSuccessful)
                {
                    MessageBox.Show("Login successful!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    // עכשיו שומרים את ההתחברות לתוך loginforstudent.xlsx
                    Excel.Workbook loginWorkbook = null;
                    Excel.Worksheet loginSheet = null;
                    bool loginFileExists = File.Exists(loginLogPath);

                    if (loginFileExists)
                    {
                        loginWorkbook = excelApp.Workbooks.Open(loginLogPath);
                        loginSheet = (Excel.Worksheet)loginWorkbook.Sheets[1];
                    }
                    else
                    {
                        loginWorkbook = excelApp.Workbooks.Add();
                        loginSheet = (Excel.Worksheet)loginWorkbook.Sheets[1];
                        loginSheet.Cells[1, 1] = "Username";
                        loginSheet.Cells[1, 2] = "Password";
                        loginSheet.Cells[1, 3] = "Login Time";
                    }

                    int lastRow = loginSheet.UsedRange.Rows.Count + 1;
                    loginSheet.Cells[lastRow, 1] = username;
                    loginSheet.Cells[lastRow, 2] = password;
                    loginSheet.Cells[lastRow, 3] = DateTime.Now.ToString("g");

                    // שמירה לקובץ
                    loginWorkbook.Save();

                    loginWorkbook.Close(false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(loginSheet);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(loginWorkbook);
                    studentlist signupForm = new studentlist();
                    signupForm.Show();
                }
                else
                {
                    MessageBox.Show("Invalid username or password.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (excelApp != null)
                {
                    excelApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                }
            }
            
        }

        
    }

}


