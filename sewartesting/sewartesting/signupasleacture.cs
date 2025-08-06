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

    public partial class signupasleacture : Form
    {
        public signupasleacture()
        {
            InitializeComponent();
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string name = textBox1.Text.Trim();
            string email = textBox2.Text.Trim();
            string password = textBox3.Text.Trim();
            string id = textBox4.Text.Trim();

            // Validation - Empty
            if (string.IsNullOrEmpty(name) || string.IsNullOrEmpty(email) || string.IsNullOrEmpty(password) || string.IsNullOrEmpty(id))
            {
                MessageBox.Show("Please fill all fields.");
                return;
            }

            // Validation - Name (username)
            if (name.Length < 6 || name.Length > 8 || !System.Text.RegularExpressions.Regex.IsMatch(name, @"^[a-zA-Z0-9]+$"))
            {
                MessageBox.Show("Username must be 6–8 characters long and contain only English letters and digits.");
                return;
            }

            // Validation - Password
            if (password.Length < 8 || password.Length > 10 ||
                !password.Any(char.IsDigit) ||
                !password.Any(char.IsLetter) ||
                !password.Any(ch => "!@#$%^&*()_+-=[]{}|;:',.<>?/".Contains(ch)))
            {
                MessageBox.Show("Password must be 8–10 characters long and contain at least one letter, one digit, and one special character.");
                return;
            }

            // Validation - Email
            if (!System.Text.RegularExpressions.Regex.IsMatch(email, @"^[^@\s]+@[^@\s]+\.[^@\s]+$"))
            {
                MessageBox.Show("Please enter a valid email address.");
                return;
            }

            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;

            try
            {
                string solutionPath = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName;
                string folderPath = Path.Combine(solutionPath, "data");

                if (!Directory.Exists(folderPath))
                    Directory.CreateDirectory(folderPath);

                string filePath = Path.Combine(folderPath, "signupasleacture.xlsx");

                excelApp = new Excel.Application();
                excelApp.Visible = false;

                bool fileExists = File.Exists(filePath);
                workbook = fileExists ? excelApp.Workbooks.Open(filePath) : excelApp.Workbooks.Add();
                worksheet = (Excel.Worksheet)workbook.Sheets[1];

                if (!fileExists || worksheet.Cells[1, 1].Value == null)
                {
                    worksheet.Cells[1, 1] = "Name";
                    worksheet.Cells[1, 2] = "Email";
                    worksheet.Cells[1, 3] = "Password";
                    worksheet.Cells[1, 4] = "IDLECTURE";
                }

                int lastRow = worksheet.UsedRange.Rows.Count;

                for (int i = 2; i <= lastRow; i++)
                {
                    string existingName = Convert.ToString((worksheet.Cells[i, 1] as Excel.Range).Value);

                    string existingEmail = Convert.ToString((worksheet.Cells[i, 2] as Excel.Range).Value);
                    string existingID = Convert.ToString((worksheet.Cells[i, 4] as Excel.Range).Value);
                    if (!string.IsNullOrEmpty(existingName) && name.Equals(existingName, StringComparison.OrdinalIgnoreCase))
                    {
                        MessageBox.Show("This username is already taken.");
                        return;
                    }

                    if (!string.IsNullOrEmpty(existingEmail) && email.Equals(existingEmail, StringComparison.OrdinalIgnoreCase))
                    {
                        MessageBox.Show("This email is already registered.");
                        return;
                    }

                    if (!string.IsNullOrEmpty(existingID) && id.Equals(existingID))
                    {
                        MessageBox.Show("This student ID is already registered.");
                        return;
                    }
                }

                int newRow = lastRow + 1;
                worksheet.Cells[newRow, 1] = name;
                worksheet.Cells[newRow, 2] = email;
                worksheet.Cells[newRow, 3] = password;
                worksheet.Cells[newRow, 4] = id;

                workbook.Save();
                MessageBox.Show("Data saved successfully!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
            finally
            {
                if (worksheet != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                if (workbook != null)
                {
                    workbook.Close(false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                }
                if (excelApp != null)
                {
                    excelApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                }
            }
            loginleacture signupForm = new loginleacture();
            signupForm.Show();
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
