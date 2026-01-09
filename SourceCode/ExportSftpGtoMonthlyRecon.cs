using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Microsoft.Data.SqlClient;
using ClosedXML.Excel;

class ExportMonthlyRecon
{
    static void Main()
    {
        // Read connection string from environment variable
        string connectionString =
            Environment.GetEnvironmentVariable("DB_CONNECTION_STRING");

        if (string.IsNullOrEmpty(connectionString))
        {
            Console.WriteLine("Database connection string not found.");
            return;
        }

        string fileName =
            "SFTP_GTO_" + DateTime.Now.AddMonths(-1).ToString("MMM_yyyy") + ".xlsx";

        string excelPath = Path.Combine("Exports", fileName);

        List<string> outlets = new List<string>
        {
            "OUTLET1", "OUTLET2", "OUTLET3"
        };

        using SqlConnection con = new SqlConnection(connectionString);
        con.Open();

        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Monthly_Recon");

        int startColumn = 1;

        foreach (string outlet in outlets)
        {
            using SqlCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "usp_GetMonthlySftpGtoReport";
            cmd.Parameters.Add("@OutletCode", SqlDbType.VarChar).Value = outlet;

            using SqlDataReader dr = cmd.ExecuteReader();

            int row = 1;
            while (dr.Read())
            {
                ws.Cell(row, startColumn).Value = dr["DocDate"];
                ws.Cell(row, startColumn + 1).Value = dr["SFTP_Amount"];
                ws.Cell(row, startColumn + 2).Value = dr["GTO_Amount"];
                ws.Cell(row, startColumn + 3).Value =
                    Convert.ToDecimal(dr["SFTP_Amount"]) -
                    Convert.ToDecimal(dr["GTO_Amount"]);

                row++;
            }

            startColumn += 5;
        }

        Directory.CreateDirectory("Exports");
        wb.SaveAs(excelPath);

        Console.WriteLine("Excel report generated successfully.");
    }
}

