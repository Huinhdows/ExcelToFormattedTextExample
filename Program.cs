using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DriverLicenseProcessing
{
    class Program
    {
        // STAThread (single threaded attribute) declaration is required in order to utilize System.Windows.Forms.OpenFileDialog
        [STAThread]
        static void Main(string[] args)
        {
            OpenFileDialog fileBrowser = new OpenFileDialog();
            DialogResult result = fileBrowser.ShowDialog();
            if (result == DialogResult.OK)
            {
                DataTable dtFileContents = ImportExceltoDatatable(fileBrowser.FileName);
                List<DriverRecord> driverRecords = GetDriverRecords(dtFileContents);
                string filePath = WriteDriverRecordsToFile(driverRecords);
                SendFileAsEmailAttachment(filePath);
                Console.WriteLine("Formatted driver license data outputted to "); 
                Console.WriteLine(filePath.Replace(@"\\SomeServer\SomeDirectory", @"S:\SomeDirectory"));
                Console.WriteLine("Press any key to exit this tool");
                Console.ReadKey();
            }
            else
            {
                Console.WriteLine("Invalid file selected");
                Console.WriteLine("Press any key to exit this tool");
                Console.ReadKey();
            }
        }

        // Converts the provided Excel spreadsheet to a DataTable structure
        private static DataTable ImportExceltoDatatable(string filepath)
        {
            DataSet dsFileContents = new DataSet();
            string excelConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filepath + ";Extended Properties=\"Excel 12.0;HDR=YES;\"";
            using (OleDbConnection connection = new OleDbConnection(excelConnectionString))
            {
                connection.Open();
                DataTable dtSchema = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                string sheet1 = dtSchema.Rows[0].Field<string>("TABLE_NAME");
                string sqlQuery = String.Format("SELECT * FROM [{0}]", sheet1);
                OleDbDataAdapter dataAdapter = new OleDbDataAdapter(sqlQuery, connection);
                dataAdapter.Fill(dsFileContents);
            }
            return dsFileContents.Tables[0];
        }

        // Creates a List of DriverRecord objects with correctly formatted data for processing
        private static List<DriverRecord> GetDriverRecords(DataTable dtFileContents)
        {
            List<DriverRecord> driverRecords = new List<DriverRecord>();
            foreach (DataRow row in dtFileContents.Rows)
            {
                DriverRecord newRecord = new DriverRecord();
                newRecord.License = row["LICENSE"].ToString().PadRight(8, ' ');
                newRecord.Department = row["DEPARTMENT"].ToString().Substring(0, 6);
                string formattedPayroll = "00" + row["PAYROLL"].ToString().Substring(0,4).Trim();
                newRecord.Payroll = formattedPayroll;
                string formattedLastName = row["LAST NAME"].ToString().ToUpper().PadRight(15, ' ');
                if (formattedLastName.Length > 15)
                    formattedLastName = formattedLastName.Substring(0, 15);
                newRecord.LastName = formattedLastName;
                driverRecords.Add(newRecord);
            }
            return driverRecords;
        }

        // Writes properly formatted DriverRecord text data to a .txt file in S:\Tech Strategy & Support\Development Team\Documents\DriversLicences\
        private static string WriteDriverRecordsToFile(List<DriverRecord> driverRecords)
        {
            string fileDirectory = String.Format(@"\\SomeServer\SomeDirectory", DateTime.Today.ToString("MMMM yyyy"));
            string fileName = String.Format("FileName{0}.txt", DateTime.Today.ToString("yyyyMMdd"));
            string filePath = fileDirectory + fileName;
            Directory.CreateDirectory(fileDirectory);
            File.Create(filePath).Dispose();
            TextWriter writter = new StreamWriter(filePath);
            foreach (DriverRecord record in driverRecords)
            {
                string newLine = String.Format("{0}{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}{12}{13}{14}", 
                    record.TransportationCode, record.License, record.LastNameAbbreviation, 
                    record.SubAccountNumber, record.PGPrefix, record.Department, 
                    record.Payroll, record.LastName, record.Address, 
                    record.DispositionCode, record.Filler, record.PurposeCode, 
                    record.AccountNumber, record.SecurityCode, record.DeptCode);
                writter.WriteLine(newLine);
            }
            writter.Close();
            return filePath;
        }

        // Sends formatted .txt file as an email attachment to dustin.huinh@pgworks.com
        private static bool SendFileAsEmailAttachment(string filePath)
        {
            string toAddress = "email.address@domain.com";
            string subject = String.Format("Driver's License file - {0}", DateTime.Today.ToString("MMMM yyyy"));
            string body = String.Format("Driver's License file for {0} has been generated and ready to be sent to the City. File is attached.", DateTime.Today.ToString("MMMM yyyy"));
            return Email.SendEmailWithAttachment(toAddress, subject, body, filePath);
        }
    }
}
