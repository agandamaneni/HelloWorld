using MGI.Common.Logging.Impl.NLog;
//using MongoDB.Bson;
//using MongoDB.Driver;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Diagnostics;
using DocumentFormat.OpenXml;
using System.Data.SqlClient;

namespace HelloWorld
{
    class Program
    {
        static string connetionString = @"Data Source=localhost\sqlexpress;Initial Catalog=LoanMapsDev;User ID=sa;Password=Opteamix@123";
        static int ClientId = 2;
        static int CreatedBy = 22;
        static int IsDeleted = 0;
        static string NoteType = "Appraiser";

        static void Main(string[] args)
        {
            using (SpreadsheetDocument spreadsheetDocument =
                    SpreadsheetDocument.Open(@"E:\appraiser_data1.xlsx", false))
            {
                WorkbookPart workBookPart = spreadsheetDocument.WorkbookPart;
                foreach (Sheet s in workBookPart.Workbook.Descendants<Sheet>())
                {
                    List<List<string>> sheetData = new List<List<string>>();
                    WorksheetPart wsPart = workBookPart.GetPartById(s.Id) as WorksheetPart;
                    Debug.WriteLine("Worksheet {1}:{2} - id({0}) {3}", s.Id, s.SheetId, s.Name,
                        wsPart == null ? "NOT FOUND!" : "found.");

                    if (wsPart == null)
                    {
                        continue;
                    }

                    Row[] rows = wsPart.Worksheet.Descendants<Row>().ToArray();
                    foreach (Row row in wsPart.Worksheet.Descendants<Row>())
                    {
                        List<string> rowData = new List<string>();
                        string value;

                        foreach (Cell c in row.Elements<Cell>())
                        {
                            value = GetCellValue(c);
                            rowData.Add(value);
                        }

                        sheetData.Add(rowData);
                    }
                    if (s.Name == "Appraisers")
                    {
                        InsertAppraiserSheet(sheetData);
                    }
                    if (s.Name == "Appraiser_Countys")
                    {
                        InsertAppraiserCountys(sheetData);
                    }
                }
            }

            using (SpreadsheetDocument spreadsheetDocument =
                SpreadsheetDocument.Open(@"E:\Title_data.xlsx", false))
            {
                WorkbookPart workBookPart = spreadsheetDocument.WorkbookPart;
                foreach (Sheet s in workBookPart.Workbook.Descendants<Sheet>())
                {
                    List<List<string>> sheetData = new List<List<string>>();
                    WorksheetPart wsPart = workBookPart.GetPartById(s.Id) as WorksheetPart;
                    Debug.WriteLine("Worksheet {1}:{2} - id({0}) {3}", s.Id, s.SheetId, s.Name,
                        wsPart == null ? "NOT FOUND!" : "found.");

                    if (wsPart == null)
                    {
                        continue;
                    }

                    Row[] rows = wsPart.Worksheet.Descendants<Row>().ToArray();
                    foreach (Row row in wsPart.Worksheet.Descendants<Row>())
                    {
                        List<string> rowData = new List<string>();
                        string value;

                        foreach (Cell c in row.Elements<Cell>())
                        {
                            value = GetCellValue(c);
                            rowData.Add(value);
                        }

                        sheetData.Add(rowData);
                    }
                    InsertTitleCompany(sheetData);
                }
            }
        }

        private static void WriteToText(string text)
        {
            // Write the string to a file.
            System.IO.StreamWriter file = File.AppendText("c:\\megastarlogfile.txt");
            file.WriteLine(text);
            file.WriteLine("\r\n");

            file.Close();
        }

        public static string GetCellValue(Cell cell)
        {
            if (cell == null)
                return string.Empty;
            if (cell.DataType == null)
                return cell.InnerText;

            string value = cell.InnerText;
            switch (cell.DataType.Value)
            {
                case CellValues.SharedString:
                    // For shared strings, look up the value in the shared strings table.
                    // Get worksheet from cell
                    OpenXmlElement parent = cell.Parent;
                    while (parent.Parent != null && parent.Parent != parent
                            && string.Compare(parent.LocalName, "worksheet", true) != 0)
                    {
                        parent = parent.Parent;
                    }
                    if (string.Compare(parent.LocalName, "worksheet", true) != 0)
                    {
                        throw new Exception("Unable to find parent worksheet.");
                    }

                    Worksheet ws = parent as Worksheet;
                    SpreadsheetDocument ssDoc = ws.WorksheetPart.OpenXmlPackage as SpreadsheetDocument;
                    SharedStringTablePart sstPart = ssDoc.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

                    // lookup value in shared string table
                    if (sstPart != null && sstPart.SharedStringTable != null)
                    {
                        value = sstPart.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                    }
                    break;

                //this case within a case is copied from msdn. 
                case CellValues.Boolean:
                    switch (value)
                    {
                        case "0":
                            value = "FALSE";
                            break;
                        default:
                            value = "TRUE";
                            break;
                    }
                    break;
            }
            value = string.IsNullOrWhiteSpace(value) ? null : value;
            return value;
        }

        public static void InsertAppraiserSheet(List<List<string>> sheetData)
        {
            // columns for sheet FirstName = 0, LastName = 1, NickName = 2, AppEmail = 3, CellNumber = 4, 
            // BussPhNum = 5, LicenseNum = 6, CarrierName = 7, CoverageAmount = 8, Expiration = 9, CONV = 10, FHA = 11, Notes = 12
            SqlCommand commandAppraiser;
            SqlCommand commandAdminAppraiserOrder;
            SqlCommand commandAppraiserEOCoverage;
            SqlCommand commandAppraiserLoanNotes;
            using (SqlConnection connection = new SqlConnection(connetionString))
            {
                int rows = sheetData.Count();
                int appraiserCount = 0;
                for (int i = 1; i < rows; i++)
                {
                    connection.Close();
                    connection.Open();

                    int appraiserId = GetAppraiserId(sheetData[i][0], sheetData[i][1], sheetData[i][3]);
                    if (appraiserId > 0)
                    {
                        connection.Close();
                        continue;
                    }
                    // insert into Appraiser
                    string insertAppraiser = string.Format("INSERT INTO Appraiser(FirstName, LastName, ClientId, Createdby, CreatedDate, IsDeleted) VALUES(@FirstName, @LastName, @ClientId, @Createdby, @CreatedDate, @IsDeleted);SELECT CAST(scope_identity() AS int)");
                    commandAppraiser = new SqlCommand(insertAppraiser, connection);
                    commandAppraiser.Parameters.AddWithValue("@FirstName", sheetData[i][0]);
                    commandAppraiser.Parameters.AddWithValue("@LastName", sheetData[i][1]);
                    commandAppraiser.Parameters.AddWithValue("@ClientId", ClientId);
                    commandAppraiser.Parameters.AddWithValue("@Createdby", CreatedBy);
                    commandAppraiser.Parameters.AddWithValue("@CreatedDate", DateTime.Now);
                    commandAppraiser.Parameters.AddWithValue("@IsDeleted", IsDeleted);
                    try
                    {
                        appraiserCount = (int)commandAppraiser.ExecuteScalar();
                    }
                    catch(Exception ex)
                    {
                        WriteToText(string.Format("Appraiser Name : {0} {1}, Exception : {2}", sheetData[i][0], sheetData[i][1], ex.ToString()));
                        continue;
                    }
                    commandAppraiser.Dispose();
                    connection.Close();

                    connection.Open();
                    // insert into AdminAppraiserOrder
                    string insertAdminAppraiserOrder = string.Format("INSERT INTO AdminAppraiserOrder(AppraiserId, Email, ClientId, Createdby, CreatedDate, IsDeleted, MortgageTypes) VALUES(@AppraiserId, @Email, @ClientId, @Createdby, @CreatedDate, @IsDeleted, @MortgageTypes)");
                    commandAdminAppraiserOrder = new SqlCommand(insertAdminAppraiserOrder, connection);
                    commandAdminAppraiserOrder.Parameters.AddWithValue("@AppraiserId", appraiserCount);
                    commandAdminAppraiserOrder.Parameters.AddWithValue("@Email", sheetData[i][3]);
                    commandAdminAppraiserOrder.Parameters.AddWithValue("@ClientId", ClientId);
                    commandAdminAppraiserOrder.Parameters.AddWithValue("@Createdby", CreatedBy);
                    commandAdminAppraiserOrder.Parameters.AddWithValue("@CreatedDate", DateTime.Now);
                    commandAdminAppraiserOrder.Parameters.AddWithValue("@IsDeleted", IsDeleted);
                    string MortgageTypes = sheetData[i][11] == "1" ?  "1," : string.Empty;
                    MortgageTypes += sheetData[i][10] == "1" ? "2," : string.Empty;
                    MortgageTypes = MortgageTypes.TrimEnd(',');
                    commandAdminAppraiserOrder.Parameters.AddWithValue("@MortgageTypes", MortgageTypes);
                    try
                    {
                        commandAdminAppraiserOrder.ExecuteReader();
                    }
                    catch (Exception ex)
                    {
                        WriteToText(string.Format("Appraiser Name : {0} {1}, Exception : {2}", sheetData[i][0], sheetData[i][1], ex.ToString()));
                        continue;
                    }
                    commandAdminAppraiserOrder.Dispose();
                    connection.Close();

                    connection.Open();
                    // insert into AppraiserEOCoverage
                    string insertAppraiserEOCoverage = string.Format("INSERT INTO AppraiserEOCoverage(CoverageAmount, CoverageExpiration, AppraiserId, ClientId, Createdby, CreatedDate, IsDeleted) VALUES(@CoverageAmount, @CoverageExpiration, @AppraiserId, @ClientId, @Createdby, @CreatedDate, @IsDeleted)");
                    commandAppraiserEOCoverage = new SqlCommand(insertAppraiserEOCoverage, connection);
                    commandAppraiserEOCoverage.Parameters.AddWithValue("@CoverageAmount", Convert.ToDouble(sheetData[i][8]));
                    commandAppraiserEOCoverage.Parameters.AddWithValue("@CoverageExpiration", DateTime.FromOADate(Convert.ToDouble(sheetData[i][9])));
                    commandAppraiserEOCoverage.Parameters.AddWithValue("@AppraiserId", appraiserCount);
                    commandAppraiserEOCoverage.Parameters.AddWithValue("@ClientId", ClientId);
                    commandAppraiserEOCoverage.Parameters.AddWithValue("@Createdby", CreatedBy);
                    commandAppraiserEOCoverage.Parameters.AddWithValue("@CreatedDate", DateTime.Now);
                    commandAppraiserEOCoverage.Parameters.AddWithValue("@IsDeleted", IsDeleted);
                    try
                    {
                        commandAppraiserEOCoverage.ExecuteReader();
                    }
                    catch (Exception ex)
                    {
                        WriteToText(string.Format("Appraiser Name : {0} {1}, Exception : {2}", sheetData[i][0], sheetData[i][1], ex.ToString()));
                        continue;
                    }
                    commandAppraiserEOCoverage.Dispose();
                    connection.Close();

                    connection.Open();
                    // insert into AppraiserLoanNotes
                    string insertAppraiserLoanNotes = string.Format("INSERT INTO AppraiserLoanNotes(AppraiserId, NotesDetails, NoteType, ClientId, Createdby, CreatedDate, IsDeleted) VALUES(@AppraiserId, @NotesDetails, @NoteType, @ClientId, @Createdby, @CreatedDate, @IsDeleted)");
                    commandAppraiserLoanNotes = new SqlCommand(insertAppraiserLoanNotes, connection);
                    commandAppraiserLoanNotes.Parameters.AddWithValue("@AppraiserId", appraiserCount);
                    commandAppraiserLoanNotes.Parameters.AddWithValue("@NotesDetails", sheetData[i][12]??string.Empty);
                    commandAppraiserLoanNotes.Parameters.AddWithValue("@NoteType", NoteType);
                    commandAppraiserLoanNotes.Parameters.AddWithValue("@ClientId", ClientId);
                    commandAppraiserLoanNotes.Parameters.AddWithValue("@Createdby", CreatedBy);
                    commandAppraiserLoanNotes.Parameters.AddWithValue("@CreatedDate", DateTime.Now);
                    commandAppraiserLoanNotes.Parameters.AddWithValue("@IsDeleted", IsDeleted);
                    try
                    {
                        commandAppraiserLoanNotes.ExecuteReader();
                    }
                    catch (Exception ex)
                    {
                        WriteToText(string.Format("Appraiser Name : {0} {1}, Exception : {2}", sheetData[i][0], sheetData[i][1], ex.ToString()));
                        continue;
                    }
                    commandAppraiserLoanNotes.Dispose();

                    connection.Close();
                }
            }
        }

        public static void InsertAppraiserCountys(List<List<string>> sheetData)
        {
            // FirstName = 0;LastName = 1;Email = 2;LicenseNumber = 3;StateCode = 4;AppCounty = 5;certified_date = 6;ExpirationDate = 7;

            using (SqlConnection connection = new SqlConnection(connetionString))
            {
                int rows = sheetData.Count();
                for (int i = 1; i < rows; i++)
                {
                    connection.Open();
                    int appraiserId = GetAppraiserId(sheetData[i][0], sheetData[i][1], sheetData[i][2]);
                    if(appraiserId > 0)
                    {
                        SqlCommand insertCmd = new SqlCommand("INSERT INTO AppraiserStateCertifications(StateId, LicenseNumber, CertifiedDate, ExpirationDate, ClientId, CreatedBy, CreatedDate, IsDeleted, AppraiserId) VALUES(@StateId, @LicenseNumber, @CertifiedDate, @ExpirationDate, @ClientId, @CreatedBy, @CreatedDate, @IsDeleted, @AppraiserId)", connection);
                        // get the state id
                        string selectState = string.Format("SELECT Id from State WHERE StateCode = @StateCode");
                        SqlCommand commandState = new SqlCommand(selectState, connection);
                        commandState.Parameters.AddWithValue("@StateCode", sheetData[i][4]);
                        int stateId = (int)commandState.ExecuteScalar();
                        commandState.Dispose();
                        // got the state id
                        insertCmd.Parameters.AddWithValue("@StateId", stateId);
                        insertCmd.Parameters.AddWithValue("@LicenseNumber", sheetData[i][3]);
                        insertCmd.Parameters.AddWithValue("@CertifiedDate", DateTime.FromOADate(Convert.ToDouble(sheetData[i][6])));
                        insertCmd.Parameters.AddWithValue("@ExpirationDate", DateTime.FromOADate(Convert.ToDouble(sheetData[i][7])));
                        insertCmd.Parameters.AddWithValue("@ClientId", ClientId);
                        insertCmd.Parameters.AddWithValue("@CreatedBy", CreatedBy);
                        insertCmd.Parameters.AddWithValue("@CreatedDate", DateTime.Now);
                        insertCmd.Parameters.AddWithValue("@IsDeleted", IsDeleted);
                        insertCmd.Parameters.AddWithValue("@AppraiserId", appraiserId);
                        try
                        {
                            insertCmd.ExecuteReader();
                        }
                        catch(Exception ex)
                        {
                            WriteToText(string.Format("Insert County Failed. Appraiser name {0} {1} email {2}; Exception : {3}", sheetData[i][0], sheetData[i][1], sheetData[i][2], ex.ToString()));
                        }
                    }
                    connection.Close();
                }
            }
        }

        public static void InsertTitleCompany(List<List<string>> sheetData)
        {
            // title_company_name = 0;address_line1	= 1; City = 2; State = 3;zip_code = 4; PhoneNumber	= 5;FaxNumber = 6;	Email	= 7;BankName	= 8; Aba = 9;	BankCity = 10; BankState = 11; BankAccountName = 12; BankAccountNumber = 13; ForFurtherCredit = 14;	TitleInsurer = 15;
            SqlCommand commandTitle;
            SqlCommand commandState;
            SqlCommand commandBank;
            SqlCommand commandInsurance;
            using (SqlConnection connection = new SqlConnection(connetionString))
            {
                int rows = sheetData.Count();
                int titleCompanyId = 0;
                for (int i = 1; i < rows; i++)
                {
                    connection.Close();
                    connection.Open();
                    // insert into TitleCompany
                    string insertTitle = string.Format("INSERT INTO TitleCompany(TitleCompanyName, AddressLine1, City, StateId, ZipCode, PhoneNumber, Email, Fax, ClientId, Createdby, CreatedDate, IsDeleted) VALUES(@TitleCompanyName, @AddressLine1, @City, @StateId, @ZipCode, @PhoneNumber, @Email, @Fax, @ClientId, @Createdby, @CreatedDate, @IsDeleted);SELECT CAST(scope_identity() AS int)");
                    commandTitle = new SqlCommand(insertTitle, connection);
                    commandTitle.Parameters.AddWithValue("@TitleCompanyName", sheetData[i][0]);
                    commandTitle.Parameters.AddWithValue("@AddressLine1", sheetData[i][1]);
                    commandTitle.Parameters.AddWithValue("@City", sheetData[i][2]);
                    // get the state id
                    string selectState = string.Format("SELECT Id from State WHERE StateCode = @StateCode");
                    commandState = new SqlCommand(selectState, connection);
                    commandState.Parameters.AddWithValue("@StateCode", sheetData[i][3]);
                    int stateId = (int)commandState.ExecuteScalar();
                    commandState.Dispose();
                    // got the state id
                    commandTitle.Parameters.AddWithValue("@StateId", stateId);
                    commandTitle.Parameters.AddWithValue("@ZipCode", sheetData[i][4]);
                    commandTitle.Parameters.AddWithValue("@PhoneNumber", sheetData[i][5]);
                    commandTitle.Parameters.AddWithValue("@Email", sheetData[i][7]);
                    commandTitle.Parameters.AddWithValue("@Fax", sheetData[i][6]);
                    commandTitle.Parameters.AddWithValue("@ClientId", ClientId);
                    commandTitle.Parameters.AddWithValue("@Createdby", CreatedBy);
                    commandTitle.Parameters.AddWithValue("@CreatedDate", DateTime.Now);
                    commandTitle.Parameters.AddWithValue("@IsDeleted", IsDeleted);
                    try
                    {
                         titleCompanyId = (int)commandTitle.ExecuteScalar();
                    }
                    catch(Exception ex)
                    {
                        WriteToText(string.Format("company title : {0}, exception : {1}", sheetData[i][0], ex.ToString()));
                        continue;
                    }

                    string insertBank = string.Format("INSERT INTO TitleCompanyWiringInstruction(BankName, ABA, BankCity, BankStateId, BankAccountNumber, BankAccountName, ForFurtherCredit, ClientId, Createdby, CreatedDate, IsDeleted, TitleCompanyId) VALUES(@BankName, @ABA, @BankCity, @BankStateId, @BankAccountNumber, @BankAccountName, @ForFurtherCredit, @ClientId, @Createdby, @CreatedDate, @IsDeleted, @TitleCompanyId)");
                    commandBank = new SqlCommand(insertBank, connection);
                    commandBank.Parameters.AddWithValue("@BankName", sheetData[i][8]);
                    commandBank.Parameters.AddWithValue("@ABA", sheetData[i][9]);
                    commandBank.Parameters.AddWithValue("@BankCity", sheetData[i][10]);
                    // get the state id
                    selectState = string.Format("SELECT Id from State WHERE StateCode = @StateCode");
                    commandState = new SqlCommand(selectState, connection);
                    commandState.Parameters.AddWithValue("@StateCode", sheetData[i][3]);
                    stateId = (int)commandState.ExecuteScalar();
                    commandState.Dispose();
                    // got the state id
                    commandBank.Parameters.AddWithValue("@BankStateId", stateId);
                    commandBank.Parameters.AddWithValue("@BankAccountNumber", sheetData[i][13]);
                    commandBank.Parameters.AddWithValue("@BankAccountName", sheetData[i][12]);
                    commandBank.Parameters.AddWithValue("@ForFurtherCredit", sheetData[i][14]??string.Empty);
                    commandBank.Parameters.AddWithValue("@ClientId", ClientId);
                    commandBank.Parameters.AddWithValue("@Createdby", CreatedBy);
                    commandBank.Parameters.AddWithValue("@CreatedDate", DateTime.Now);
                    commandBank.Parameters.AddWithValue("@IsDeleted", IsDeleted);
                    commandBank.Parameters.AddWithValue("@TitleCompanyId", titleCompanyId);
                    try
                    {
                        commandBank.ExecuteReader();
                    }
                    catch(Exception ex)
                    {
                        WriteToText(string.Format("bank name : {0}, exception : {1}", sheetData[i][8], ex.ToString()));
                        continue;
                    }
                    if(!string.IsNullOrWhiteSpace(sheetData[i][15]))
                    {
                        int enumerationId = GetEnumerationId(sheetData[i][15]);
                        connection.Close();
                        connection.Open();
                        commandInsurance = new SqlCommand("INSERT INTO TitleInsurer(TitleCompanyId, SelectedInsurerId, ClientId, CreatedBy, CreatedDate, IsDeleted) VALUES(@TitleCompanyId, @SelectedInsurerId, @ClientId, @CreatedBy, @CreatedDate, @IsDeleted)", connection);
                        commandInsurance.Parameters.AddWithValue("@TitleCompanyId", titleCompanyId);
                        commandInsurance.Parameters.AddWithValue("@SelectedInsurerId", enumerationId);
                        commandInsurance.Parameters.AddWithValue("@ClientId", ClientId );
                        commandInsurance.Parameters.AddWithValue("@CreatedBy", CreatedBy);
                        commandInsurance.Parameters.AddWithValue("@CreatedDate", DateTime.Now );
                        commandInsurance.Parameters.AddWithValue("@IsDeleted", IsDeleted);
                        try
                        {
                            commandInsurance.ExecuteReader();
                        }
                        catch (Exception ex)
                        {
                            WriteToText(string.Format("title insurance name : {0}, exception : {1}", sheetData[i][15], ex.ToString()));
                            continue;
                        }
                    }
                    connection.Close();
                }
            }
        }

        public static int GetEnumerationId(string titleInsurance)
        {
            using (SqlConnection connection = new SqlConnection(connetionString))
            {
                connection.Open();
                int enumerationId = 0;
                SqlCommand selectCmd = new SqlCommand("SELECT id FROM Enumeration WHERE Value = @Value", connection);
                selectCmd.Parameters.AddWithValue("@Value", titleInsurance);
                try
                {
                    enumerationId = (int)selectCmd.ExecuteScalar();
                }
                catch (Exception ex)
                {
                    WriteToText(string.Format("Enumaration not found. titleInsurance : {0}", titleInsurance));
                    SqlCommand insertCmd = new SqlCommand("INSERT INTO Enumeration(Value, [Type], Sequence, IsDeleted, EnumType, IsMiniAudit, ClientId, CreatedBy, CreatedDate, [Key]) VALUES(@Value, @Type, @Sequence, @IsDeleted, @EnumType, @IsMiniAudit, @ClientId, @CreatedBy, @CreatedDate, @Key);SELECT CAST(scope_identity() AS int)", connection);
                    insertCmd.Parameters.AddWithValue("@Value", titleInsurance);
                    insertCmd.Parameters.AddWithValue("@Type", 21);
                    insertCmd.Parameters.AddWithValue("@Sequence", 0);
                    insertCmd.Parameters.AddWithValue("@IsDeleted", IsDeleted);
                    insertCmd.Parameters.AddWithValue("@EnumType", "TitleInsurer");
                    insertCmd.Parameters.AddWithValue("@IsMiniAudit", 0);
                    insertCmd.Parameters.AddWithValue("@ClientId", ClientId);
                    insertCmd.Parameters.AddWithValue("@CreatedBy", CreatedBy);
                    insertCmd.Parameters.AddWithValue("@CreatedDate", DateTime.Now);
                    insertCmd.Parameters.AddWithValue("@Key", titleInsurance.Replace(" ", string.Empty));
                    try
                    {
                        enumerationId = (int)insertCmd.ExecuteScalar();
                    }
                    catch(Exception ex1)
                    {
                        WriteToText(string.Format("Title Ensurance Name : {0}, Exception : {1}", titleInsurance, ex1.ToString()));
                    }
                    insertCmd.Dispose();
                } 
                selectCmd.Dispose();
                connection.Close();
                return enumerationId;
            }
        }

        public static int GetAppraiserId(string firstName, string lastName, string email)
        {
            using (SqlConnection connection = new SqlConnection(connetionString))
            {
                connection.Open();
                int appraiserId = 0;
                SqlCommand selectAppraiser = new SqlCommand("select Appraiser.id from Appraiser join AdminAppraiserOrder on Appraiser.id = AdminAppraiserOrder.AppraiserId where firstname = @FirstName and lastname = @LastName and email = @Email", connection);
                selectAppraiser.Parameters.AddWithValue("@FirstName", firstName);
                selectAppraiser.Parameters.AddWithValue("@LastName", lastName);
                selectAppraiser.Parameters.AddWithValue("@Email", email);

                try
                {
                    var result = selectAppraiser.ExecuteScalar();
                    if (result != null)
                        appraiserId = (int)result;
                }
                catch (Exception ex)
                {
                    WriteToText("Appraiser find error. exception : " + ex.ToString());
                }
                return appraiserId;
            }
        }
    }
}