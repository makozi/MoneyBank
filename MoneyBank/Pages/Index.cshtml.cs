using System;
using System.Collections.Generic;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Http;
using ExcelDataReader;
using System.IO;
using System.Data;
using System.Text;
using System.Data.SqlClient;
using System.Globalization;
using System.Threading.Tasks;

namespace MoneyBank.Pages
{
    public class IndexModel : PageModel
    {
        private readonly string connectionString = "Data Source=TJ-LP-001-0046;Initial Catalog=transactions;Integrated Security=True"; // Replace with your MSSQL connection string

        public List<TransactionRecord> TransactionRecords { get; set; }

        public void OnGet()
        {
            // Initialization logic for the GET request, if needed.
        }

        public async Task<IActionResult> OnPost(IFormFile excelFile)
        {
            if (excelFile == null || excelFile.Length == 0)
            {
                ModelState.AddModelError("ExcelFile", "Please select a file.");
                return Page();
            }

            var extension = Path.GetExtension(excelFile.FileName);
            if (extension.ToLower() != ".xlsx")
            {
                ModelState.AddModelError("ExcelFile", "Invalid file format. Please select an Excel file (.xlsx).");
                return Page();
            }

            // Register encoding provider to handle encoding issues
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            using (var stream = excelFile.OpenReadStream())
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet(new ExcelDataSetConfiguration
                    {
                        ConfigureDataTable = _ => new ExcelDataTableConfiguration
                        {
                            UseHeaderRow = true
                        }
                    });

                    var dataTable = result.Tables[0];

                    if (dataTable != null)
                    {
                        TransactionRecords = new List<TransactionRecord>();

                        for (int rowIndex = 0; rowIndex < dataTable.Rows.Count; rowIndex++)
                        {
                            var row = dataTable.Rows[rowIndex];

                            TransactionRecord record = new TransactionRecord();
                            record.SerialNumber = Convert.ToInt32(row[0]);
                            record.DebitAccount = row[1].ToString();
                            record.CreditAccount = row[2].ToString();

                            // Safely attempt to convert to decimal
                            if (decimal.TryParse(row[3].ToString(), NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out decimal amount))
                            {
                                record.Amount = amount;
                            }
                            else
                            {
                                ModelState.AddModelError("ExcelFile", $"Invalid amount in row {rowIndex + 1}");
                                continue;
                            }

                            record.Narration = row[4].ToString();

                            ValidateTransaction(record);

                            if (record.IsValid)
                            {
                                TransactionRecords.Add(record);

                                // Insert valid data into the MSSQL database
                                InsertIntoDatabase(record);
                            }
                        }
                    }
                }
            }

            return Page();
        }

        private bool ValidateTransaction(TransactionRecord record)
        {
            if (record.DebitAccount.Length == 10 && record.CreditAccount.Length == 10
                && record.Amount > 0 && record.Amount <= 100000.00m)
            {
                record.IsValid = true;
                return true;
            }

            record.IsValid = false;
            return false;
        }

        private void InsertIntoDatabase(TransactionRecord record)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                using (SqlCommand command = connection.CreateCommand())
                {
                    command.CommandText = "INSERT INTO BulkTransaction(SrlNumber, DebitAccount, CreditAccount, Amount, Narration) " +
                                          "VALUES (@SerialNumber, @DebitAccount, @CreditAccount, @Amount, @Narration)";

                    command.Parameters.Add(new SqlParameter("@SerialNumber", SqlDbType.Int) { Value = record.SerialNumber });
                    command.Parameters.Add(new SqlParameter("@DebitAccount", SqlDbType.NVarChar, 10) { Value = record.DebitAccount });
                    command.Parameters.Add(new SqlParameter("@CreditAccount", SqlDbType.NVarChar, 10) { Value = record.CreditAccount });
                    command.Parameters.Add(new SqlParameter("@Amount", SqlDbType.Decimal) { Value = record.Amount });
                    command.Parameters.Add(new SqlParameter("@Narration", SqlDbType.NVarChar) { Value = record.Narration });

                    command.ExecuteNonQuery();
                }
            }
        }
    }

    public class TransactionRecord
    {
        public int SerialNumber { get; set; }
        public string DebitAccount { get; set; }
        public string CreditAccount { get; set; }
        public decimal Amount { get; set; }
        public string Narration { get; set; }
        public bool IsValid { get; set; }
    }
}
