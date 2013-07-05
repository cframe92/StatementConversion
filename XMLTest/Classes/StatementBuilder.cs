using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace XMLTest.Classes
{
    public class StatementBuilder
    {
        public static void Build(statementProduction statement, MoneyPerksStatement moneyPerksStatement)
        {
            using (FileStream outputStream = File.Create("C:\\" + TEMP_FILE_NAME))
            {
                CreateFirstPage(statement, outputStream);

                if (statement.CheckingAccounts.Count > 0)
                {
                    AddCheckingAccounts(statement);
                }

                if (statement.SavingsAccounts.Count > 0)
                {
                    AddSavingsAccounts(statement);
                }

                if (statement.ClubAccounts.Count > 0)
                {
                    AddClubAccounts(statement);
                }

                if (statement.CertificateAccounts.Count > 0)
                {
                    AddCertificateAccounts(statement);
                }

                if (statement.Loans.Count > 0)
                {
                    AddLoanAccounts(statement);
                }

                AddYtdSummaries(statement);
                AddMoneyPerksSummary(moneyPerksStatement);
                AddBottomAdvertising(statement);
                Doc.Close();
            }

            AddPageNumbersAndDisclosures(statement); // Re-opens document to overlay page numbers

            if (File.Exists("c:\\" + TEMP_FILE_NAME))
            {
                File.Delete("c:\\" + TEMP_FILE_NAME);
            }

            NumberOfStatementsBuilt++;
        }

        static void CreateFirstPage(statementProduction statement, FileStream outputStream)
        {
            //Adds first page template to statement
            using(FileStream templateInputStream = File.Open(Configuration.GetStatementTemplateFirstPageFilePath(), FileMode.Open))
            {
                PdfReader reader = new PdfReader(templateInputStream);
                Doc = new Document(reader.GetPageSize(1));
                Writer = PdfWriter.GetInstance(Doc, outputStream);
                StatementPageEvent pageEvent = new StatementPageEvent();
                Writer.PageEvent = pageEvent;
                Writer.SetFullCompression();
                Doc.Open();
                PdfContentByte contentByte = Writer.DirectContent;
                PdfImportedPage page = Writer.GetImportedPage(reader, 1);
                Doc.NewPage();
                contentByte.AddTemplate(page, 0, 0);
            }

            AddStatementHeading("Statement  of  Accounts", 409, 0);
            AddStatementHeading(statement.envelope[0].statement.beginningStatementDate.ToString("MMM  dd,  yyyy") + "  thru  " + statement.envelope[0].statement.endingStatementDate.ToString("MMM  dd,  yyyy"), 385, 6f);

            if (statement.envelope[0].statement.account.accountNumber > 4)
            {
                AddStatementHeading("Account  Number:        ******" + statement.envelope[0].statement.account.accountNumber.Substring("******".Length), 385, 6f);
            }

            AddInvisibleAccountNumber(statement);

            PdfPTable addressAndBalancesTable = new PdfPTable(2);
            float[] addressAndBalancesTableWidths = new float[] { 50f, 50f };
            addressAndBalancesTable.SetWidthPercentage(addressAndBalancesTableWidths, Doc.PageSize);
            addressAndBalancesTable.TotalWidth = 612f;
            addressAndBalancesTable.LockedWidth = true;
            AddAddress(statement, ref addressAndBalancesTable);
            AddHeaderBalances(statement, ref addressAndBalancesTable);
            Doc.Add(addressAndBalancesTable);
            AddTopAdvertising(statement);
        }
            
        static void AddAddress(statementProduction statement, ref PdfPTable table)
        {
            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(statement.envelope[0].address.ToUpper(), GetNormalFont(9));

            foreach (AdditionalName additionalName in statement.AdditionalNames)
            {
                chunk.Append("\n" + additionalName.Name + ", " + additionalName.TypeString);
            }

            if (statement.Address.AddressLine2 != null) chunk.Append("\n" + statement.Address.AddressLine2.ToUpper());
            if (statement.Address.AddressLine3 != null) chunk.Append("\n" + statement.Address.AddressLine3.ToUpper());
            if (statement.Address.AddressLine4 != null) chunk.Append("\n" + statement.Address.AddressLine4.ToUpper());
            if (statement.Address.AddressLine5 != null) chunk.Append("\n" + statement.Address.AddressLine5.ToUpper());
            if (statement.Address.AddressLine6 != null) chunk.Append("\n" + statement.Address.AddressLine6.ToUpper());
            //if (statement.Address.City != null) chunk.Append("\n" + statement.Address.City.ToUpper());
            //if (statement.Address.State != null) chunk.Append(" " + statement.Address.State.ToUpper());
            //if (statement.Address.ZipCode != null) chunk.Append(" " + statement.Address.ZipCode.ToUpper());
            chunk.setLineHeight(9f);
            Paragraph p = new Paragraph(chunk);
            p.IndentationLeft = 66;
            cell.AddElement(p);
            cell.PaddingTop = 60f;
            cell.BorderWidth = 0f;
            table.AddCell(cell);
        }

        static void GetBalance(statementProduction statement, ref PdfPTable table)
        {
            for(int i = 0; i < ; i ++)
            {

            }

        }

        static void AddHeaderBalances(statementProduction statement, ref PdfPTable table)
        {
            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk("Account  Balances  at  a  Glance:", GetBoldFont(12f));
            Paragraph p = new Paragraph(chunk);
            p.IndentationLeft = 81;
            cell.AddElement(p);
            cell.PaddingTop = 6f;
            cell.BorderWidth = 0f;
            cell.BorderColor = BaseColor.LIGHT_GRAY;
            PdfPTable balancesTable = new PdfPTable(2);
            float[] tableWidths = new float[] { 60f, 40f };
            balancesTable.SetWidthPercentage(tableWidths, Doc.PageSize);
            balancesTable.TotalWidth = 300f;
            balancesTable.LockedWidth = true;
            AddHeaderBalanceTitle("Total  Checking:", ref balancesTable);
            AddHeaderBalanceValue(statement.EndingBalanceChecking, ref balancesTable);
            AddHeaderBalanceTitle("Total  Savings:", ref balancesTable);
            AddHeaderBalanceValue(statement.EndingBalanceSavings, ref balancesTable);
            AddHeaderBalanceTitle("Total  Loans:", ref balancesTable);
            AddHeaderBalanceValue(statement.EndingBalanceLoans, ref balancesTable);
            AddHeaderBalanceTitle("Total  Certificates:", ref balancesTable);
            AddHeaderBalanceValue(statement.EndingBalanceCertificates, ref balancesTable);
            cell.AddElement(balancesTable);
            table.AddCell(cell);
        }

        static void AddHeaderBalanceTitle(string title, ref PdfPTable balancesTable)
        {
            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(title, GetNormalFont(12f));
            Paragraph p = new Paragraph(chunk);
            p.IndentationLeft = 75;
            cell.AddElement(p);
            cell.PaddingTop = -7f;
            cell.BorderWidth = 0f;
            balancesTable.AddCell(cell);
        }

        static void AddHeaderBalanceValue(decimal value, ref PdfPTable balancesTable)
        {
            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(FormatAmount(value), GetBoldFont(12f));
            Paragraph p = new Paragraph(chunk);
            p.Alignment = Element.ALIGN_RIGHT;
            p.IndentationRight = 28;
            cell.AddElement(p);
            cell.PaddingTop = -7f;
            cell.BorderWidth = 0f;
            balancesTable.AddCell(cell);
        }

        static void AddStatementHeading(string text, float indentationLeft, float paddingTop)
        {
            PdfPTable table = new PdfPTable(1);
            table.TotalWidth = 612f;
            table.LockedWidth = true;
            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(text, GetNormalFont(12f));
            Paragraph p = new Paragraph(chunk);
            p.IndentationLeft = indentationLeft;
            cell.AddElement(p);
            cell.PaddingTop = paddingTop;
            cell.BorderWidth = 0f;
            table.AddCell(cell);
            Doc.Add(table);
        }

        static void AddInvisibleAccountNumber(statementProduction statement)
        {
            PdfPTable table = new PdfPTable(1);
            table.TotalWidth = 612f;
            table.LockedWidth = true;
            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(statement.envelope[0].statement.account.accountNumber, GetNormalFont(5f));
            Paragraph p = new Paragraph(chunk);
            p.Font.SetColor(255, 255, 255);
            cell.AddElement(p);
            cell.PaddingTop = -9f;
            cell.BorderWidth = 0f;
            table.AddCell(cell);
            Doc.Add(table);
        }

        static void AddCheckingAccounts(statementProduction statement)
        {
            AddSectionHeading("CHECKING ACCOUNTS");


            for (int i = 0; i < statement.CheckingAccounts.Count(); i++)
            {
                statementProduction account = statement.CheckingAccounts[i];

                AddAccountSubHeading(statement.envelope[i].statement.account.subAccount[i].loan.description, i>0);
                AddAccountTransactions(account);

                // Adds APR
                if (statement.envelope[i].statement.account.subAccount[i].loan.beginning.annualRate > 0)
                {
                    PdfPTable table = new PdfPTable(1);
                    table.TotalWidth = 525f;
                    table.LockedWidth = true;
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("Annual Percentage Yield Earned " + statement.envelope[i].statement.account.subAccount[i].loan.beginning.annualRate.ToString("N3") + "% from " + account.AnnualPercentageRate.BeginningDate.ToString("MM/dd/yyyy") + " through " + account.AnnualPercentageRate.EndingDate.ToString("MM/dd/yyyy"), GetBoldItalicFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    //p.IndentationLeft = 70;
                    cell.AddElement(p);
                    cell.PaddingTop = -4f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                    Doc.Add(table);
                }

                if (account.CheckHolds.Count() > 0)
                {
                    AddCheckHolds(account);
                }

                if (statement.Checks.Count > 0)
                {
                    AddChecks(statement);
                }

                if (account.AtmWithdrawals.Count > 0)
                {
                    AddAtmWithdrawals(account);
                }

                if (account.AtmDeposits.Count > 0)
                {
                    AddAtmDeposits(account);
                }

                if (account.Withdrawals.Count > 0)
                {
                    AddWithdrawals(account);
                }

                if (account.Deposits.Count > 0)
                {
                    AddDeposits(account);
                }

                if ((account.TotalOverdraftFee.AmountYtd + account.TotalReturnedItemFee.AmountYtd) > 0)
                {
                    AddTotalFees(account);
                }
            }
        }

        static void AddSavingsAccounts(statementProductionEnvelope statement)
        {
            AddSectionHeading("SAVINGS ACCOUNTS");
            while(statement.statement.account.subAccount[]. == "")

            for(int i = 0; i < statement.statement.account.("").Count(); i++)
            {
                Account account = statement.SavingsAccounts[i];

                AddAccountSubHeading(account.Description, i > 0);
                AddAccountTransactions(account);

                
                // Adds APR
                if(account.AnnualPercentageRate.Rate > 0)
                {
                    PdfPTable table = new PdfPTable(1);
                    table.TotalWidth = 525f;
                    table.LockedWidth = true;
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("Annual Percentage Yield Earned " + account.AnnualPercentageRate.Rate.ToString("N3") + "% from " + account.AnnualPercentageRate.BeginningDate.ToString("MM/dd/yyyy") + " through " + account.AnnualPercentageRate.EndingDate.ToString("MM/dd/yyyy"), GetBoldItalicFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    //p.IndentationLeft = 70;
                    cell.AddElement(p);
                    cell.PaddingTop = -4f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                    Doc.Add(table);
                }

                if (account.CheckHolds.Count() > 0)
                {
                    AddCheckHolds(account);
                }

                if (account.AtmWithdrawals.Count > 0)
                {
                    AddAtmWithdrawals(account);
                }

                if (account.AtmDeposits.Count > 0)
                {
                    AddAtmDeposits(account);
                }

                if (account.Withdrawals.Count > 0)
                {
                    AddWithdrawals(account);
                }

                if (account.Deposits.Count > 0)
                {
                    AddDeposits(account);
                }

                if ((account.TotalOverdraftFee.AmountYtd + account.TotalReturnedItemFee.AmountYtd) > 0)
                {
                    AddTotalFees(account);
                }
            }
        }

        static void AddClubAccounts(statementProduction statement)
        {
            AddSectionHeading("CLUB ACCOUNTS");


            for (int i = 0; i < statement.ClubAccounts.Count(); i++)
            {
                Account account = statement.ClubAccounts[i];

                AddAccountSubHeading(account.Description, i > 0);
                AddAccountTransactions(account);


                // Adds APR
                if (account.AnnualPercentageRate.Rate > 0)
                {
                    PdfPTable table = new PdfPTable(1);
                    table.TotalWidth = 525f;
                    table.LockedWidth = true;
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("Annual Percentage Yield Earned " + statement.envelope[i].statement.account.subAccount[i].loan.beginning.annualRate.ToString("N3") + "% from " + account.AnnualPercentageRate.BeginningDate.ToString("MM/dd/yyyy") + " through " + account.AnnualPercentageRate.EndingDate.ToString("MM/dd/yyyy"), GetBoldItalicFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    //p.IndentationLeft = 70;
                    cell.AddElement(p);
                    cell.PaddingTop = -4f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                    Doc.Add(table);
                }

                if (account.CheckHolds.Count() > 0)
                {
                    AddCheckHolds(account);
                }

                if (account.AtmWithdrawals.Count > 0)
                {
                    AddAtmWithdrawals(account);
                }

                if (account.AtmDeposits.Count > 0)
                {
                    AddAtmDeposits(account);
                }

                if (account.Withdrawals.Count > 0)
                {
                    AddWithdrawals(account);
                }

                if (account.Deposits.Count > 0)
                {
                    AddDeposits(account);
                }

                if ((account.TotalOverdraftFee.AmountYtd + account.TotalReturnedItemFee.AmountYtd) > 0)
                {
                    AddTotalFees(account);
                }
            }
        }

        static void AddCertificateAccounts(statementProduction statement)
        {
            AddSectionHeading("CERTIFICATE ACCOUNTS");
            var count = 0;

            for(int i = 0; i < count; i++)
            {
                if(statement.envelope[i].statement.account.subAccount[i].share.category.Value == "Share")
                {
                    statementProductionEnvelopeStatementAccountSubAccount shareActs = statement.envelope[i].statement.account.subAccount[i];
                    string descriptionAndMaturityDate = statement.envelope[i].statement.account.subAccount[i].share.description + "   Maturity Date - " + statement.envelope[i].statement.account.subAccount[i].share.maturityDate.ToString("MMM dd, yyyy");


                    AddAccountSubHeading(descriptionAndMaturityDate, i > 0);
                    AddAccountTransactions(account);


                    if (account.CheckHolds.Count() > 0)
                    {
                        AddCheckHolds(account);
                    }

                    if (account.AtmWithdrawals.Count > 0)
                    {
                        AddAtmWithdrawals(account);
                    }

                    if (account.AtmDeposits.Count > 0)
                    {
                        AddAtmDeposits(account);
                    }

                    if (account.Withdrawals.Count > 0)
                    {
                        AddWithdrawals(account);
                    }

                    if (account.Deposits.Count > 0)
                    {
                        AddDeposits(account);
                    }

                    if ((account.TotalOverdraftFee.AmountYtd + account.TotalReturnedItemFee.AmountYtd) > 0)
                    {
                        AddTotalFees(account);
                    }

                }
            }
        }

        static void AddLoanAccounts(statementProduction statement)
        {
            AddSectionHeading("LOAN ACCOUNTS");

            for(int i = 0; i < statement.epilogue.loanCount; i++)
            {
                
                AddAccountSubHeading(statement.envelope[i].statement.account.subAccount[i].loan.description, i > 0);
                AddLoanPaymentInformation(statement.envelope[i].statement.account.subAccount[i].loan);
                AddLoanTransactions(statement.envelope[i].statement.account.subAccount[i].loan);
                if (!loan.Closed)
                {
                    AddLoanTransactionsFooter("Closing Date of Billing Cycle " + loan.ClosingDateOfBillingCycle.ToString("MM/dd/yyyy") + "\n" +
                        "** INTEREST CHARGE CALCULATION: The balance used to compute interest charges is the unpaid balance each day after payments and credits to that balance have been subtracted and any additions to the balance have been made.");
                    AddFeeSummary(loan);
                    AddInterestChargedSummary(loan);
                }
                AddYearToDateTotals(loan);

                if(statement.envelope[i].statement.account.subAccount[i].loan.category.option == "A")
                {
                    AddAdvances(loan);
                }
               

                if (loan.Payments.Count() > 0)
                {
                    AddLoanPaymentsSortTable(loan);
                }
            }
        }

        static void AddSectionHeading(string title)
        {
            if (Writer.GetVerticalPosition(false) <= 175)
            {
                Doc.NewPage();
            }

            AddHeadingStroke();
            PdfPTable table = new PdfPTable(1);
            table.TotalWidth = 525f;
            table.LockedWidth = true;
            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(title, GetBoldFont(16f));
            chunk.SetCharacterSpacing(0f);
            Paragraph p = new Paragraph(chunk);
            cell.AddElement(p);
            cell.PaddingTop = -7f;
            cell.BorderWidth = 0f;
            table.AddCell(cell);
            Doc.Add(table);
        }

        static void AddAccountTransactions(statementProduction account)
        {
            PdfPTable  table = new PdfPTable(5);
            table.HeaderRows = 1;
            float[] tableWidths = new float[] { 51, 280, 62, 65, 67 };
            table.TotalWidth = 525f;
            table.SetWidths(tableWidths);
            table.LockedWidth = true;
            AddAccountTransactionTitle("Date", Element.ALIGN_LEFT, ref table);
            AddAccountTransactionTitle("Transaction Description", Element.ALIGN_LEFT, ref table);
            AddAccountTransactionTitle("Additions", Element.ALIGN_RIGHT, ref table);
            AddAccountTransactionTitle("Subtractions", Element.ALIGN_RIGHT, ref table);
            AddAccountTransactionTitle("Balance", Element.ALIGN_RIGHT, ref table);
            AddBalanceForward(account.envelope[].statement.account.subAccount[].loan.beginning.balance, ref table);

            foreach (statementProductionEnvelopeStatement transaction in account.envelope[0].statement.account.subAccount[0].loan.transaction[0])
            {
                if (!transaction.account.)
                {
                    AddAccountTransaction(transaction, ref table);
                }
                else
                {
                    AddCommentOnlyTransaction(transaction, ref table);
                }
            }

            if (!account.Closed)
            {
                AddEndingBalance(account, ref table);
            }
            else
            {
                AddShareClosed(account, ref table);
            }

            Doc.Add(table);
        }

        static void AddAccountTransaction(statementProduction transaction, ref PdfPTable table)
        {
            // Adds Date
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(transaction.envelope[].statement.ToString("MMM dd"), GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }
            // Adds Transaction Description
            {
                string description = string.Empty;

                if (transaction.envelope[0].statement.account. > 0)
                {
                    description = transaction.envelope[0].statement.account.subAccount[].loan.description;
                }

                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(description, GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);

                // Adds additional description lines
                for(int i = 0; i < transaction.DescriptionLines.Count; i++)
                {
                    if(i > 0)
                    {
                        chunk = new Chunk(transaction.DescriptionLines[i], GetNormalFont(9f));
                        chunk.setLineHeight(11f);
                        chunk.SetCharacterSpacing(0f);
                        p = new Paragraph(chunk);
                        p.IndentationLeft = 20;
                        cell.AddElement(p);
                        cell.NoWrap = false;
                    }
                }

                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            if (transaction.Amount >= 0)
            {
                AddAccountTransactionAmount(transaction.Amount, ref table); // Additions
                AddAccountTransactionAmount(0, ref table); // Subtractions
            }
            else
            {
                AddAccountTransactionAmount(0, ref table); // Additions
                AddAccountTransactionAmount(transaction.Amount, ref table); // Subtractions
            }

            AddAccountBalance(transaction.Balance, ref table);
        }

        static void AddCommentOnlyTransaction(statementProduction transaction, ref PdfPTable table)
        {
            // Adds Date
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }
            // Adds Transaction Description
            {
                string description = string.Empty;

                if (transaction.DescriptionLines.Count > 0)
                {
                    description = transaction.DescriptionLines[0];
                }

                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(description, GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                p.IndentationLeft = 20;
                cell.AddElement(p);
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                cell.NoWrap = true;
                table.AddCell(cell);
            }

            AddAccountTransactionAmount(0, ref table); // Additions
            AddAccountTransactionAmount(0, ref table); // Subtractions
            AddAccountTransactionAmount(0, ref table);
        }

        static void AddBalanceForward(Account account, ref PdfPTable table)
        {
            // Adds Date
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(account.BeginDate.ToString("MMM dd"), GetBoldItalicFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);

            }
            // Adds Transaction Description
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("Balance Forward", GetBoldItalicFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            AddAccountTransactionAmount(0, ref table);
            AddAccountTransactionAmount(0, ref table);
            AddAccountBalance(account.BalanceForward, ref table);
        }

        static void AddEndingBalance(statementProductionEnvelopeStatement account, ref PdfPTable table)
        {
            // Date
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(account.endingStatementDate.ToString("MMM dd"), GetBoldItalicFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);

            }

            // Transaction Description
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("Ending Balance", GetBoldItalicFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            AddAccountTransactionAmount(0, ref table);
            AddAccountTransactionAmount(0, ref table);
            AddAccountBalance(account.endingStatementDate, ref table);
        }

        static void AddShareClosed(statementProduction account, ref PdfPTable table)
        {
            // Date
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(string.Empty, GetBoldItalicFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);

            }

            // Transaction Description
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(account.ShortDescription + " " + "Closed\n*** This is the final statement you will receive for this account ***\n*** Please retain this final statement for tax reporting purposes ***", GetBoldItalicFont(9f));
                chunk.SetCharacterSpacing(0f);
                chunk.setLineHeight(11f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.NoWrap = true;
                cell.PaddingTop = -1f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            AddAccountTransactionAmount(0, ref table);
            AddAccountTransactionAmount(0, ref table);
            AddAccountTransactionAmount(0, ref table);
        }

        static void AddLoanClosed(statementProduction loan, ref PdfPTable table)
        {
            // Adds Date
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -6f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }
            // Adds Transaction Description
            {

                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(loan.ShortDescription + " " + "Closed\n*** This is the final statement you will receive for this account ***\n*** Please retain this final statement for tax reporting purposes ***", GetBoldItalicFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.NoWrap = true;
                cell.PaddingTop = -6f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            for (int i = 0; i < 5; i++)
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                p.Alignment = Element.ALIGN_RIGHT;
                cell.AddElement(p);
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                cell.BorderColor = BaseColor.LIGHT_GRAY;
                table.AddCell(cell);
            }
        }

        static void AddAccountTransactionAmount(decimal amount, ref PdfPTable table)
        {
            string amountFormatted = string.Empty;

            if (amount != 0)
            {
                amountFormatted = FormatAmount(amount);
            }

            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(amountFormatted, GetBoldFont(9f));
            chunk.SetCharacterSpacing(0f);
            Paragraph p = new Paragraph(chunk);
            p.Alignment = Element.ALIGN_RIGHT;
            cell.AddElement(p);
            cell.PaddingTop = -4f;
            cell.BorderWidth = 0f;
            cell.BorderColor = BaseColor.LIGHT_GRAY;
            table.AddCell(cell);
        }

        static void AddLoanAccountTransactionAmount(decimal amount, ref PdfPTable table)
        {
            string amountFormatted = FormatAmount(amount);

            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(amountFormatted, GetNormalFont(9f));
            chunk.SetCharacterSpacing(0f);
            Paragraph p = new Paragraph(chunk);
            p.Alignment = Element.ALIGN_RIGHT;
            cell.AddElement(p);
            cell.PaddingTop = -4f;
            cell.BorderWidth = 0f;
            cell.BorderColor = BaseColor.LIGHT_GRAY;
            table.AddCell(cell);
        }

        static void AddMoneyPerksTransactionAmount(int amount, ref PdfPTable table)
        {
            string amountFormatted = string.Empty;

            if (amount != 0)
            {
                amountFormatted = amount.ToString();
            }

            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(amountFormatted, GetBoldFont(9f));
            chunk.SetCharacterSpacing(0f);
            Paragraph p = new Paragraph(chunk);
            p.Alignment = Element.ALIGN_RIGHT;
            cell.AddElement(p);
            cell.PaddingTop = -4f;
            cell.BorderWidth = 0f;
            cell.BorderColor = BaseColor.LIGHT_GRAY;
            table.AddCell(cell);
        }

        static void AddAccountBalance(decimal balance, ref PdfPTable table)
        {
            string amountFormatted = FormatAmount(balance);

            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(amountFormatted, GetBoldFont(9f));
            chunk.SetCharacterSpacing(0f);
            Paragraph p = new Paragraph(chunk);
            p.Alignment = Element.ALIGN_RIGHT;
            cell.AddElement(p);
            cell.PaddingTop = -4f;
            cell.BorderWidth = 0f;
            cell.BorderColor = BaseColor.LIGHT_GRAY;
            table.AddCell(cell);
        }

        static void AddMoneyPerksBalance(int balance, ref PdfPTable table)
        {
            string amountFormatted =balance.ToString();

            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(amountFormatted, GetBoldFont(9f));
            chunk.SetCharacterSpacing(0f);
            Paragraph p = new Paragraph(chunk);
            p.Alignment = Element.ALIGN_RIGHT;
            cell.AddElement(p);
            cell.PaddingTop = -4f;
            cell.BorderWidth = 0f;
            cell.BorderColor = BaseColor.LIGHT_GRAY;
            table.AddCell(cell);
        }

        static void AddAccountTransactionTitle(string title, int alignment, ref PdfPTable table)
        {
            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(title, GetBoldFont(9f));
            chunk.SetCharacterSpacing(0f);
            chunk.SetUnderline(0.75f, -2);
            Paragraph p = new Paragraph(chunk);
            p.Alignment = alignment;
            cell.AddElement(p);
            cell.PaddingTop = 6f;
            cell.BorderWidth = 0f;
            table.AddCell(cell);
        }

        static void AddAccountSubHeading(string subtitle, bool stroke)
        {
            if (Writer.GetVerticalPosition(false) <= 130)
            {
                Doc.NewPage();
            }

            float cellPaddingTop = -1f;

            if (stroke)
            {
                cellPaddingTop = -6f;
                AddSubHeadingStroke();
            }

            PdfPTable table = new PdfPTable(1);
            table.TotalWidth = 525f;
            table.LockedWidth = true;
            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(subtitle, GetBoldFont(12f));
            chunk.SetCharacterSpacing(0f);
            Paragraph p = new Paragraph(chunk);
            cell.AddElement(p);
            cell.PaddingTop = cellPaddingTop;
            cell.BorderWidth = 0f;
            table.AddCell(cell);
            Doc.Add(table);
        }

        static void AddCheckHolds(statementProductionEnvelopeStatementAccount account)
        {
            foreach(CheckHold hold in account.CheckHolds)
            {
                PdfPTable table = new PdfPTable(1);
                table.TotalWidth = 525f;
                table.LockedWidth = true;
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("Check hold placed on " + hold.EffectiveDate.ToString("MM/dd/yyyy") + " in the amount of $" + FormatAmount(hold.Amount) + " to be released on " + hold.ExpiredDate.ToString("MM/dd/yyyy"), GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                p.IndentationLeft = 70;
                cell.AddElement(p);
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
                Doc.Add(table);
            }
        }

        static void AddChecks(statementProduction statement)
        {
            bool asteriskFound = false;
            int rowBreakPointIndex = (int)Math.Ceiling((double)statement.Checks.Count / 2);
            PdfPTable table = new PdfPTable(7);
            table.HeaderRows = 1;
            float[] tableWidths = new float[] { 91, 43, 76, 105, 91, 43, 76 };
            table.TotalWidth = 525f;
            table.SetWidths(tableWidths);
            table.LockedWidth = true;

            AddSortTableHeading("CHECK SUMMARY");

            AddSortTableTitle("Check #", Element.ALIGN_LEFT, ref table);
            AddSortTableTitle("Amount", Element.ALIGN_RIGHT, ref table);
            AddSortTableTitle("Date", Element.ALIGN_RIGHT, ref table);
            AddSortTableTitle(string.Empty, Element.ALIGN_RIGHT, ref table);
            AddSortTableTitle("Check #", Element.ALIGN_LEFT, ref table);
            AddSortTableTitle("Amount", Element.ALIGN_RIGHT, ref table);
            AddSortTableTitle("Date", Element.ALIGN_RIGHT, ref table);

            // Create columns
            List<SortTableRow> rows = new List<SortTableRow>();

            for (int i = 0; i < statement.Checks.Count; i++)
            {
                if (statement.Checks[i].CheckNumber.Contains('*'))
                {
                    asteriskFound = true;
                }

                if ((i + 1) <= rowBreakPointIndex)
                {
                    rows.Add(new SortTableRow());
                    rows[i].Column.Add(statement.Checks[i].CheckNumber);
                    rows[i].Column.Add(FormatAmount(statement.Checks[i].Amount));
                    rows[i].Column.Add(statement.Checks[i].Date.ToString("MMM dd"));
                    rows[i].Column.Add(string.Empty);
                    rows[i].Column.Add(string.Empty);
                    rows[i].Column.Add(string.Empty);
                    rows[i].Column.Add(string.Empty);
                }
                else
                {
                    rows[i - rowBreakPointIndex].Column[4] = statement.Checks[i].CheckNumber;
                    rows[i - rowBreakPointIndex].Column[5] = FormatAmount(statement.Checks[i].Amount);
                    rows[i - rowBreakPointIndex].Column[6] = statement.Checks[i].Date.ToString("MMM dd");
                }
            }

            for (int i = 0; i < rows.Count; i++)
            {
                AddSortTableValue(rows[i].Column[0], Element.ALIGN_LEFT, ref table);  // Adds Check #
                AddSortTableValue(rows[i].Column[1], Element.ALIGN_RIGHT, ref table); // Adds Amount
                AddSortTableValue(rows[i].Column[2], Element.ALIGN_RIGHT, ref table); // Adds Date
                AddSortTableValue(rows[i].Column[3], Element.ALIGN_RIGHT, ref table); // Empty column title
                AddSortTableValue(rows[i].Column[4], Element.ALIGN_LEFT, ref table);  // Adds Check #
                AddSortTableValue(rows[i].Column[5], Element.ALIGN_RIGHT, ref table); // Adds Amount
                AddSortTableValue(rows[i].Column[6], Element.ALIGN_RIGHT, ref table); // Adds Date
            }

            Doc.Add(table);

            if (asteriskFound)
            {
                AddChecksFootnote();
            }

            if (statement.Checks.Count() > 1)
            {
                AddSortTableSubtotal(statement.Checks.Count().ToString() + " Checks Cleared for " + FormatAmount(statement.ChecksTotal));
            }
        }

        static void AddAtmWithdrawals(Account account)
        {
            int rowBreakPointIndex = (int)Math.Ceiling((double)account.AtmWithdrawals.Count / 2);
            PdfPTable table = new PdfPTable(6);
            table.HeaderRows = 1;
            float[] tableWidths = new float[] { 36, 63, 163.5f, 36, 63, 163.5f };
            table.TotalWidth = 525f;
            table.SetWidths(tableWidths);
            table.LockedWidth = true;
            table.SpacingBefore = 10f;

            AddSortTableHeading("ATM WITHDRAWALS AND OTHER CHARGES");

            AddSortTableTitle("Date", Element.ALIGN_LEFT, ref table);
            AddSortTableTitle("Amount", Element.ALIGN_RIGHT, ref table);
            AddSortTableTitle("Description", Element.ALIGN_LEFT, 10f, ref table);
            AddSortTableTitle("Date", Element.ALIGN_LEFT, ref table);
            AddSortTableTitle("Amount", Element.ALIGN_RIGHT, ref table);
            AddSortTableTitle("Description", Element.ALIGN_LEFT, 10f, ref table);

            // Create columns
            List<SortTableRow> rows = new List<SortTableRow>();

            for (int i = 0; i < account.AtmWithdrawals.Count; i++)
            {
                if ((i + 1) <= rowBreakPointIndex)
                {
                    rows.Add(new SortTableRow());
                    rows[i].Column.Add(account.AtmWithdrawals[i].Date.ToString("MMM dd"));
                    rows[i].Column.Add(FormatAmount(account.AtmWithdrawals[i].Amount));
                    rows[i].Column.Add(account.AtmWithdrawals[i].Description);
                    rows[i].Column.Add(string.Empty);
                    rows[i].Column.Add(string.Empty);
                    rows[i].Column.Add(string.Empty);
                }
                else
                {
                    rows[i - rowBreakPointIndex].Column[3] = account.AtmWithdrawals[i].Date.ToString("MMM dd");
                    rows[i - rowBreakPointIndex].Column[4] = FormatAmount(account.AtmWithdrawals[i].Amount);
                    rows[i - rowBreakPointIndex].Column[5] = account.AtmWithdrawals[i].Description;
                }
            }

            for (int i = 0; i < rows.Count; i++)
            {
                AddSortTableValue(rows[i].Column[0], Element.ALIGN_LEFT, ref table);  // Adds Date
                AddSortTableValue(rows[i].Column[1], Element.ALIGN_RIGHT, ref table); // Adds Amount
                AddSortTableValue(rows[i].Column[2], Element.ALIGN_LEFT, 10f, ref table); // Adds Description
                AddSortTableValue(rows[i].Column[3], Element.ALIGN_LEFT, ref table);  // Adds Date
                AddSortTableValue(rows[i].Column[4], Element.ALIGN_RIGHT, ref table); // Adds Amount
                AddSortTableValue(rows[i].Column[5], Element.ALIGN_LEFT, 10f, ref table); // Adds Description
            }

            Doc.Add(table);

            if (account.AtmWithdrawals.Count() > 1)
            {
                AddSortTableSubtotal(account.AtmWithdrawals.Count().ToString() + " ATM Withdrawals and Other Charges for " + FormatAmount(account.AtmWithdrawalsTotal));
            }
        }

        static void AddAtmDeposits(Account account)
        {
            int rowBreakPointIndex = (int)Math.Ceiling((double)account.AtmDeposits.Count / 2);
            PdfPTable table = new PdfPTable(6);
            table.HeaderRows = 1;
            float[] tableWidths = new float[] { 36, 63, 163.5f, 36, 63, 163.5f };
            table.TotalWidth = 525f;
            table.SetWidths(tableWidths);
            table.LockedWidth = true;
            table.SpacingBefore = 10f;

            AddSortTableHeading("ATM DEPOSITS AND OTHER CHARGES");

            AddSortTableTitle("Date", Element.ALIGN_LEFT, ref table);
            AddSortTableTitle("Amount", Element.ALIGN_RIGHT, ref table);
            AddSortTableTitle("Description", Element.ALIGN_LEFT, 10f, ref table);
            AddSortTableTitle("Date", Element.ALIGN_LEFT, ref table);
            AddSortTableTitle("Amount", Element.ALIGN_RIGHT, ref table);
            AddSortTableTitle("Description", Element.ALIGN_LEFT, 10f, ref table);

            // Create columns
            List<SortTableRow> rows = new List<SortTableRow>();

            for (int i = 0; i < account.AtmDeposits.Count; i++)
            {
                if ((i + 1) <= rowBreakPointIndex)
                {
                    rows.Add(new SortTableRow());
                    rows[i].Column.Add(account.AtmDeposits[i].Date.ToString("MMM dd"));
                    rows[i].Column.Add(FormatAmount(account.AtmDeposits[i].Amount));
                    rows[i].Column.Add(account.AtmDeposits[i].Description);
                    rows[i].Column.Add(string.Empty);
                    rows[i].Column.Add(string.Empty);
                    rows[i].Column.Add(string.Empty);
                }
                else
                {
                    rows[i - rowBreakPointIndex].Column[3] = account.AtmDeposits[i].Date.ToString("MMM dd");
                    rows[i - rowBreakPointIndex].Column[4] = FormatAmount(account.AtmDeposits[i].Amount);
                    rows[i - rowBreakPointIndex].Column[5] = account.AtmDeposits[i].Description;
                }
            }

            for (int i = 0; i < rows.Count; i++)
            {
                AddSortTableValue(rows[i].Column[0], Element.ALIGN_LEFT, ref table);  // Adds Date
                AddSortTableValue(rows[i].Column[1], Element.ALIGN_RIGHT, ref table); // Adds Amount
                AddSortTableValue(rows[i].Column[2], Element.ALIGN_LEFT, 10f, ref table); // Adds Description
                AddSortTableValue(rows[i].Column[3], Element.ALIGN_LEFT, ref table);  // Adds Date
                AddSortTableValue(rows[i].Column[4], Element.ALIGN_RIGHT, ref table); // Adds Amount
                AddSortTableValue(rows[i].Column[5], Element.ALIGN_LEFT, 10f, ref table); // Adds Description
            }

            Doc.Add(table);

            if (account.AtmDeposits.Count() > 1)
            {
                AddSortTableSubtotal(account.AtmDeposits.Count().ToString() + " ATM Deposits and Other Charges for " + FormatAmount(account.AtmDepositsTotal));
            }
        }

        static void AddWithdrawals(Account account)
        {
            int rowBreakPointIndex = (int)Math.Ceiling((double)account.Withdrawals.Count / 2);
            PdfPTable table = new PdfPTable(6);
            table.HeaderRows = 1;
            float[] tableWidths = new float[] { 36, 63, 163.5f, 36, 63, 163.5f };
            table.TotalWidth = 525f;
            table.SetWidths(tableWidths);
            table.LockedWidth = true;
            table.SpacingBefore = 10f;

            AddSortTableHeading("WITHDRAWALS AND OTHER CHARGES");

            AddSortTableTitle("Date", Element.ALIGN_LEFT, ref table);
            AddSortTableTitle("Amount", Element.ALIGN_RIGHT, ref table);
            AddSortTableTitle("Description", Element.ALIGN_LEFT, 10f, ref table);
            AddSortTableTitle("Date", Element.ALIGN_LEFT, ref table);
            AddSortTableTitle("Amount", Element.ALIGN_RIGHT, ref table);
            AddSortTableTitle("Description", Element.ALIGN_LEFT, 10f, ref table);

            // Create columns
            List<SortTableRow> rows = new List<SortTableRow>();

            for (int i = 0; i < account.Withdrawals.Count; i++)
            {
                if ((i + 1) <= rowBreakPointIndex)
                {
                    rows.Add(new SortTableRow());
                    rows[i].Column.Add(account.Withdrawals[i].Date.ToString("MMM dd"));
                    rows[i].Column.Add(FormatAmount(account.Withdrawals[i].Amount));
                    rows[i].Column.Add(account.Withdrawals[i].Description);
                    rows[i].Column.Add(string.Empty);
                    rows[i].Column.Add(string.Empty);
                    rows[i].Column.Add(string.Empty);
                }
                else
                {
                    rows[i - rowBreakPointIndex].Column[3] = account.Withdrawals[i].Date.ToString("MMM dd");
                    rows[i - rowBreakPointIndex].Column[4] = FormatAmount(account.Withdrawals[i].Amount);
                    rows[i - rowBreakPointIndex].Column[5] = account.Withdrawals[i].Description;
                }
            }

            for (int i = 0; i < rows.Count; i++)
            {
                AddSortTableValue(rows[i].Column[0], Element.ALIGN_LEFT, ref table);  // Adds Date
                AddSortTableValue(rows[i].Column[1], Element.ALIGN_RIGHT, ref table); // Adds Amount
                AddSortTableValue(rows[i].Column[2], Element.ALIGN_LEFT, 10f, ref table); // Adds Description
                AddSortTableValue(rows[i].Column[3], Element.ALIGN_LEFT, ref table);  // Adds Date
                AddSortTableValue(rows[i].Column[4], Element.ALIGN_RIGHT, ref table); // Adds Amount
                AddSortTableValue(rows[i].Column[5], Element.ALIGN_LEFT, 10f, ref table); // Adds Description
            }

            Doc.Add(table);

            if (account.Withdrawals.Count() > 1)
            {
                AddSortTableSubtotal(account.Withdrawals.Count().ToString() + " Withdrawals and Other Charges for " + FormatAmount(account.WithdrawalsTotal));
            }
        }

        static void AddAdvances(Loan loan)
        {
            int rowBreakPointIndex = (int)Math.Ceiling((double)loan.Advances.Count / 2);
            PdfPTable table = new PdfPTable(6);
            table.HeaderRows = 1;
            float[] tableWidths = new float[] { 36, 63, 163.5f, 36, 63, 163.5f };
            table.TotalWidth = 525f;
            table.SetWidths(tableWidths);
            table.LockedWidth = true;
            table.SpacingBefore = 10f;

            AddSortTableHeading("LOAN ADVANCES AND OTHER CHARGES");

            AddSortTableTitle("Date", Element.ALIGN_LEFT, ref table);
            AddSortTableTitle("Amount", Element.ALIGN_RIGHT, ref table);
            AddSortTableTitle("Description", Element.ALIGN_LEFT, 10f, ref table);
            AddSortTableTitle("Date", Element.ALIGN_LEFT, ref table);
            AddSortTableTitle("Amount", Element.ALIGN_RIGHT, ref table);
            AddSortTableTitle("Description", Element.ALIGN_LEFT, 10f, ref table);

            // Create columns
            List<SortTableRow> rows = new List<SortTableRow>();

            for (int i = 0; i < loan.Advances.Count; i++)
            {
                if ((i + 1) <= rowBreakPointIndex)
                {
                    rows.Add(new SortTableRow());
                    rows[i].Column.Add(loan.Advances[i].Date.ToString("MMM dd"));
                    rows[i].Column.Add(FormatAmount(loan.Advances[i].Amount));
                    rows[i].Column.Add(loan.Advances[i].Description);
                    rows[i].Column.Add(string.Empty);
                    rows[i].Column.Add(string.Empty);
                    rows[i].Column.Add(string.Empty);
                }
                else
                {
                    rows[i - rowBreakPointIndex].Column[3] = loan.Advances[i].Date.ToString("MMM dd");
                    rows[i - rowBreakPointIndex].Column[4] = FormatAmount(loan.Advances[i].Amount);
                    rows[i - rowBreakPointIndex].Column[5] = loan.Advances[i].Description;
                }
            }

            for (int i = 0; i < rows.Count; i++)
            {
                AddSortTableValue(rows[i].Column[0], Element.ALIGN_LEFT, ref table);  // Adds Date
                AddSortTableValue(rows[i].Column[1], Element.ALIGN_RIGHT, ref table); // Adds Amount
                AddSortTableValue(rows[i].Column[2], Element.ALIGN_LEFT, 10f, ref table); // Adds Description
                AddSortTableValue(rows[i].Column[3], Element.ALIGN_LEFT, ref table);  // Adds Date
                AddSortTableValue(rows[i].Column[4], Element.ALIGN_RIGHT, ref table); // Adds Amount
                AddSortTableValue(rows[i].Column[5], Element.ALIGN_LEFT, 10f, ref table); // Adds Description
            }

            Doc.Add(table);

            if (loan.Advances.Count() > 1)
            {
                AddSortTableSubtotal(loan.Advances.Count().ToString() + " Advances and Other Charges for " + FormatAmount(loan.AdvancesTotal));
            }
        }

        static void AddDeposits(Account account)
        {
            int rowBreakPointIndex = (int)Math.Ceiling((double)account.Deposits.Count / 2);
            PdfPTable table = new PdfPTable(6);
            table.HeaderRows = 1;
            float[] tableWidths = new float[] { 36, 63, 163.5f, 36, 63, 163.5f };
            table.TotalWidth = 525f;
            table.SetWidths(tableWidths);
            table.LockedWidth = true;
            table.SpacingBefore = 10f;

            AddSortTableHeading("DEPOSITS AND OTHER CREDITS");

            AddSortTableTitle("Date", Element.ALIGN_LEFT, ref table);
            AddSortTableTitle("Amount", Element.ALIGN_RIGHT, ref table);
            AddSortTableTitle("Description", Element.ALIGN_LEFT, 10f, ref table);
            AddSortTableTitle("Date", Element.ALIGN_LEFT, ref table);
            AddSortTableTitle("Amount", Element.ALIGN_RIGHT, ref table);
            AddSortTableTitle("Description", Element.ALIGN_LEFT, 10f, ref table);

            // Create columns
            List<SortTableRow> rows = new List<SortTableRow>();

            for (int i = 0; i < account.Deposits.Count; i++)
            {
                if ((i + 1) <= rowBreakPointIndex)
                {
                    rows.Add(new SortTableRow());
                    rows[i].Column.Add(account.Deposits[i].Date.ToString("MMM dd"));
                    rows[i].Column.Add(FormatAmount(account.Deposits[i].Amount));
                    rows[i].Column.Add(account.Deposits[i].Description);
                    rows[i].Column.Add(string.Empty);
                    rows[i].Column.Add(string.Empty);
                    rows[i].Column.Add(string.Empty);
                }
                else
                {
                    rows[i - rowBreakPointIndex].Column[3] = account.Deposits[i].Date.ToString("MMM dd");
                    rows[i - rowBreakPointIndex].Column[4] = FormatAmount(account.Deposits[i].Amount);
                    rows[i - rowBreakPointIndex].Column[5] = account.Deposits[i].Description;
                }
            }

            for (int i = 0; i < rows.Count; i++)
            {
                AddSortTableValue(rows[i].Column[0], Element.ALIGN_LEFT, ref table);  // Adds Date
                AddSortTableValue(rows[i].Column[1], Element.ALIGN_RIGHT, ref table); // Adds Amount
                AddSortTableValue(rows[i].Column[2], Element.ALIGN_LEFT, 10f, ref table); // Adds Description
                AddSortTableValue(rows[i].Column[3], Element.ALIGN_LEFT, ref table);  // Adds Date
                AddSortTableValue(rows[i].Column[4], Element.ALIGN_RIGHT, ref table); // Adds Amount
                AddSortTableValue(rows[i].Column[5], Element.ALIGN_LEFT, 10f, ref table); // Adds Description
            }

            Doc.Add(table);

            if (account.Deposits.Count() > 1)
            {
                AddSortTableSubtotal(account.Deposits.Count().ToString() + " Deposits and Other Credits for " + FormatAmount(account.DepositsTotal));
            }
        }

        static void AddLoanPaymentsSortTable(Loan loan)
        {
            List<Deposit> loanPayments = loan.Payments;

            int rowBreakPointIndex = (int)Math.Ceiling((double)loanPayments.Count / 2);
            PdfPTable table = new PdfPTable(6);
            table.HeaderRows = 1;
            float[] tableWidths = new float[] { 36, 63, 163.5f, 36, 63, 163.5f };
            table.TotalWidth = 525f;
            table.SetWidths(tableWidths);
            table.LockedWidth = true;
            table.SpacingBefore = 10f;

            AddSortTableHeading("LOAN PAYMENTS AND OTHER CREDITS");

            AddSortTableTitle("Date", Element.ALIGN_LEFT, ref table);
            AddSortTableTitle("Amount", Element.ALIGN_RIGHT, ref table);
            AddSortTableTitle("Description", Element.ALIGN_LEFT, 10f, ref table);
            AddSortTableTitle("Date", Element.ALIGN_LEFT, ref table);
            AddSortTableTitle("Amount", Element.ALIGN_RIGHT, ref table);
            AddSortTableTitle("Description", Element.ALIGN_LEFT, 10f, ref table);

            // Create columns
            List<SortTableRow> rows = new List<SortTableRow>();

            for (int i = 0; i < loanPayments.Count; i++)
            {
                if ((i + 1) <= rowBreakPointIndex)
                {
                    rows.Add(new SortTableRow());
                    rows[i].Column.Add(loanPayments[i].Date.ToString("MMM dd"));
                    rows[i].Column.Add(FormatAmount(Math.Abs(loanPayments[i].Amount)));
                    rows[i].Column.Add(loanPayments[i].Description);
                    rows[i].Column.Add(string.Empty);
                    rows[i].Column.Add(string.Empty);
                    rows[i].Column.Add(string.Empty);
                }
                else
                {
                    rows[i - rowBreakPointIndex].Column[3] = loanPayments[i].Date.ToString("MMM dd");
                    rows[i - rowBreakPointIndex].Column[4] = FormatAmount(Math.Abs(loanPayments[i].Amount));
                    rows[i - rowBreakPointIndex].Column[5] = loanPayments[i].Description;
                }
            }

            for (int i = 0; i < rows.Count; i++)
            {
                AddSortTableValue(rows[i].Column[0], Element.ALIGN_LEFT, ref table);  // Adds Date
                AddSortTableValue(rows[i].Column[1], Element.ALIGN_RIGHT, ref table); // Adds Amount
                AddSortTableValue(rows[i].Column[2], Element.ALIGN_LEFT, 10f, ref table); // Adds Description
                AddSortTableValue(rows[i].Column[3], Element.ALIGN_LEFT, ref table);  // Adds Date
                AddSortTableValue(rows[i].Column[4], Element.ALIGN_RIGHT, ref table); // Adds Amount
                AddSortTableValue(rows[i].Column[5], Element.ALIGN_LEFT, 10f, ref table); // Adds Description
            }

            Doc.Add(table);

            if (loanPayments.Count() > 1)
            {
                AddSortTableSubtotal(loanPayments.Count().ToString() + " Payments and Other Credits for " + FormatAmount(Math.Abs(loan.PaymentsTotal)));
            }
        }

        static void AddSortTableHeading(string title)
        {
            if (Writer.GetVerticalPosition(false) <= 130)
            {
                Doc.NewPage();
            }

            PdfPTable table = new PdfPTable(1);
            table.TotalWidth = 525f;
            table.LockedWidth = true;
            table.SpacingBefore = 10f;
            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(title, GetBoldFont(9f));
            chunk.SetCharacterSpacing(0f);
            Paragraph p = new Paragraph(chunk);
            p.Alignment = Element.ALIGN_CENTER;
            cell.AddElement(p);
            cell.BorderWidth = 0f;
            table.AddCell(cell);
            Doc.Add(table);
        }

        static void AddSortTableTitle(string title,  int alignment, ref PdfPTable table)
        {
            AddSortTableTitle(title, alignment, 0, ref table);
        }

        static void AddSortTableTitle(string title, int alignment, float indentation, ref PdfPTable table)
        {
            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(title, GetBoldFont(9f));
            chunk.SetCharacterSpacing(0f);
            chunk.SetUnderline(0.75f, -2);
            Paragraph p = new Paragraph(chunk);
            p.Alignment = alignment;
            p.IndentationLeft = indentation;
            cell.AddElement(p);
            cell.PaddingTop = -4f;
            cell.BorderWidth = 0f;
            table.AddCell(cell);
        }

        static void AddSortTableValue(string value, int alignment, ref PdfPTable table)
        {
            AddSortTableValue(value, alignment, 0, ref table);
        }

        static void AddSortTableSubtotal(string value)
        {
            PdfPTable table = new PdfPTable(1);
            table.TotalWidth = 525f;
            table.LockedWidth = true;
            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(value, GetBoldItalicFont(9f));
            chunk.SetCharacterSpacing(0f);
            Paragraph p = new Paragraph(chunk);
            p.IndentationLeft = 70;
            cell.AddElement(p);
            cell.PaddingTop = -4f;
            cell.BorderWidth = 0f;
            table.AddCell(cell);
            Doc.Add(table);
        }

        static void AddChecksFootnote()
        {
            PdfPTable table = new PdfPTable(1);
            table.TotalWidth = 525f;
            table.LockedWidth = true;
            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk("* Asterisk next to number indicates skip in number sequence", GetBoldItalicFont(9f));
            chunk.SetCharacterSpacing(0f);
            Paragraph p = new Paragraph(chunk);
            //p.IndentationLeft = 70;
            cell.AddElement(p);
            cell.PaddingTop = -4f;
            cell.BorderWidth = 0f;
            table.AddCell(cell);
            Doc.Add(table);
        }

        static void AddSortTableValue(string value, int alignment, float indentation, ref PdfPTable table)
        {
            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(value, GetNormalFont(9f));
            chunk.SetCharacterSpacing(0f);
            Paragraph p = new Paragraph(chunk);
            p.Alignment = alignment;
            p.IndentationLeft = indentation;
            cell.AddElement(p);
            cell.PaddingTop = -4f;
            cell.BorderWidth = 0f;
            table.AddCell(cell);
        }

        static void AddTotalFees(Account account)
        {
            PdfPTable table = new PdfPTable(5);
            table.HeaderRows = 1;
            float[] tableWidths = new float[] { 12, 153, 79, 93, 188 };
            table.TotalWidth = 525f;
            table.SetWidths(tableWidths);
            table.LockedWidth = true;
            table.SpacingBefore = 10f;

            // For layout only
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.Padding = 6f;
                cell.PaddingTop = -2f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            AddFeeTitle("", new Border(1, 1, 0, 1) , ref table);
            AddFeeTitle("Total for\nthis period", new Border(1, 1, 0, 1), ref table);
            AddFeeTitle("Total\nyear-to-date", new Border(1, 1, 0, 1), ref table);

            // For layout only
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.Padding = 6f;
                cell.PaddingTop = -2f;
                cell.BorderWidth = 0f;
                cell.BorderWidthLeft = 1f;
                table.AddCell(cell);
            }

            // For layout only
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.Padding = 6f;
                cell.PaddingTop = -2f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            // Total Overdraft Fees
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("Total Overdraft Fees", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.Padding = 6f;
                cell.PaddingTop = -2f;
                cell.BorderWidth = 0f;
                cell.BorderWidthLeft = 1f;
                cell.BorderWidthTop = 0;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = 0;
                cell.BorderColor = BaseColor.BLACK;
                table.AddCell(cell);
            }
            AddFeeValue(account.TotalOverdraftFee.AmountThisPeriod, new Border(1, 0, 0, 0), -2f, ref table);
            AddFeeValue(account.TotalOverdraftFee.AmountYtd, new Border(1, 0, 0, 0), -2f, ref table);

            // For layout only
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.Padding = 6f;
                cell.PaddingTop = -2f;
                cell.BorderWidth = 0f;
                cell.BorderWidthLeft = 1f;
                table.AddCell(cell);
            }

            // For layout only
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.Padding = 6f;
                cell.PaddingTop = -8f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            // Total Returned Item Fees
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("Total Returned Item Fees", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.Padding = 6f;
                cell.PaddingTop = -8f;
                cell.BorderWidth = 0f;
                cell.BorderWidthLeft = 1f;
                cell.BorderWidthTop = 0;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = 1;
                cell.BorderColor = BaseColor.BLACK;
                table.AddCell(cell);
            }

            AddFeeValue(account.TotalReturnedItemFee.AmountThisPeriod, new Border(1, 0, 0, 1), -8f, ref table);
            AddFeeValue(account.TotalReturnedItemFee.AmountYtd, new Border(1, 0, 0, 1), -8f, ref table);

            // For layout only
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.Padding = 6f;
                cell.PaddingTop = -8f;
                cell.BorderWidth = 0f;
                cell.BorderWidthLeft = 1f;
                table.AddCell(cell);
            }

            Doc.Add(table);
        }

        static void AddFeeTitle(string title, Border border, ref PdfPTable table)
        {
            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(title, GetBoldFont(9));
            chunk.SetCharacterSpacing(0f);
            chunk.setLineHeight(11f);
            Paragraph p = new Paragraph(chunk);
            p.Alignment = Element.ALIGN_CENTER;
            cell.AddElement(p);
            cell.Padding = 6f;
            cell.PaddingTop = -2f;
            cell.BorderWidth = 0f;
            cell.BorderWidthLeft = border.WidthLeft;
            cell.BorderWidthTop = border.WidthTop;
            cell.BorderWidthRight = border.WidthRight;
            cell.BorderWidthBottom = border.WidthBottom;
            cell.BorderColor = BaseColor.BLACK;
            table.AddCell(cell);
        }

        static void AddFeeValue(decimal value, Border border, float paddingTop, ref PdfPTable table)
        {
            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(FormatAmount(value), GetBoldFont(9f));
            chunk.SetCharacterSpacing(0f);
            Paragraph p = new Paragraph(chunk);
            p.Alignment = Element.ALIGN_RIGHT;
            cell.AddElement(p);
            cell.Padding = 6f;
            cell.PaddingTop = paddingTop;
            cell.PaddingRight = 35f;
            cell.BorderWidth = 0f;
            cell.BorderWidthLeft = border.WidthLeft;
            cell.BorderWidthTop = border.WidthTop;
            cell.BorderWidthRight = border.WidthRight;
            cell.BorderWidthBottom = border.WidthBottom;
            cell.BorderColor = BaseColor.BLACK;
            table.AddCell(cell);
        }

        static void AddLoanPaymentInformation(Loan loan)
        {
            // Annual percentage rate
            {
                PdfPTable table = new PdfPTable(1);
                table.TotalWidth = 525f;
                table.LockedWidth = true;
                PdfPCell cell = new PdfPCell();
                Chunk chunk = null;
                if (loan.CreditLimit == null)
                {
                    chunk = new Chunk("Annual Percentage Rate:  " + loan.AnnualPercentageRate.ToString("N3") + "%", GetBoldFont(9f));
                }
                else
                {
                    chunk = new Chunk("Annual Percentage Rate:  " + loan.AnnualPercentageRate.ToString("N3") + "%    Credit Limit:    " + FormatAmount(loan.CreditLimit.Limit) + "    Available Credit:    " + FormatAmount(loan.CreditLimit.Available), GetBoldFont(9f));
                }
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                p.IndentationLeft = 28f;
                cell.AddElement(p);
                cell.PaddingTop = 7f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
                Doc.Add(table);
            }

            // PAYMENT INFORMATION
            {
                PdfPTable table = new PdfPTable(1);
                table.TotalWidth = 525f;
                table.LockedWidth = true;
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("PAYMENT INFORMATION", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = 7f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
                Doc.Add(table);
            }

            // Summary table
            {
                PdfPTable table = new PdfPTable(3);
                table.HeaderRows = 0;
                float[] tableWidths = new float[] { 93, 55, 381 };
                table.TotalWidth = 525f;
                table.SetWidths(tableWidths);
                table.LockedWidth = true;
                table.SpacingBefore = 0f;

                // Previous Balance Title
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("Previous Balance:", GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.PaddingTop = -6f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                }
                // Previous Balance
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk(loan.PreviousBalance.ToString("N"), GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    p.Alignment = Element.ALIGN_RIGHT;
                    cell.AddElement(p);
                    cell.PaddingTop = -6f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                }
                // For layout only
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk(string.Empty, GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.PaddingTop = -6f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                }
                // New Balance Title
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("New Balance:", GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.PaddingTop = -6f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                }
                // New Balance
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk(loan.NewBalance.ToString("N"), GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    p.Alignment = Element.ALIGN_RIGHT;
                    cell.AddElement(p);
                    cell.PaddingTop = -6f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                }
                // For layout only
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk(string.Empty, GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.PaddingTop = -6f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                }
                // Minimum Payment Title
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("Minimum Payment:", GetBoldFont(9f));
                    if (loan.MinimumPayment == 0)
                    {
                        chunk = new Chunk("Minimum Payment: No Payment Due", GetBoldFont(9f));
                    }
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.PaddingTop = -6f;
                    cell.BorderWidth = 0f;
                    cell.NoWrap = true;
                    table.AddCell(cell);
                }
                // Minimum Payment
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk(string.Empty, GetBoldFont(9f));
                    if (loan.MinimumPayment != 0)
                    {
                        chunk = new Chunk(loan.MinimumPayment.ToString("N"), GetBoldFont(9f));
                    }
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    p.Alignment = Element.ALIGN_RIGHT;
                    cell.AddElement(p);
                    cell.PaddingTop = -6f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                }
                // For layout only
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk(string.Empty, GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.PaddingTop = -6f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                }
                // Payment Due Date Title
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("Payment Due Date:", GetBoldFont(9f));
                    if (loan.PaymentDueDate.ToString("MM/dd/yyyy") == "01/01/0001")
                    {
                        chunk = new Chunk("Payment Due Date: No Payment Due", GetBoldFont(9f));
                    }
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.PaddingTop = -6f;
                    cell.BorderWidth = 0f;
                    cell.NoWrap = true;
                    table.AddCell(cell);
                }
                // Payment Due Date
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk(string.Empty, GetBoldFont(9f));
                    if (loan.PaymentDueDate.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        chunk = new Chunk(loan.PaymentDueDate.ToString("MM/dd/yyyy"), GetBoldFont(9f));
                    }
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    p.Alignment = Element.ALIGN_RIGHT;
                    cell.AddElement(p);
                    cell.PaddingTop = -6f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                }
                // For layout only
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk(string.Empty, GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.PaddingTop = -6f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                }

                Doc.Add(table);
            }

            // Next Payment Due Date after statement
            {
                PdfPTable table = new PdfPTable(3);
                table.HeaderRows = 0;
                float[] tableWidths = new float[] { 180, 55, 290 };
                table.TotalWidth = 525f;
                table.SetWidths(tableWidths);
                table.LockedWidth = true;
                table.SpacingBefore = 0;

                // Next Payment Due Date after statement Title
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("Next Payment Due Date after statement:", GetBoldFont(9f));
                    if (loan.NextPaymentDueDateAfterStatement.ToString("MM/dd/yyyy") == "01/01/0001")
                    {
                        chunk = new Chunk("Next Payment Due Date after statement: No Payment Due", GetBoldFont(9f));
                    }
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.PaddingTop = -6f;
                    cell.BorderWidth = 0f;
                    cell.NoWrap = true;
                    table.AddCell(cell);
                }
                // Next Payment Due Date after statement
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("", GetBoldFont(9f));
                    if (loan.NextPaymentDueDateAfterStatement.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        chunk = new Chunk(loan.NextPaymentDueDateAfterStatement.ToString("MM/dd/yyyy"), GetBoldFont(9f));
                    }
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.PaddingTop = -6f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                }
                // For layout only
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk(string.Empty, GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.PaddingTop = -6f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                }

                Doc.Add(table);
            }
        }

        static void AddLoanTransactions(Loan loan)
        {
            PdfPTable table = new PdfPTable(7);
            table.HeaderRows = 1;
            float[] tableWidths = new float[] { 51, 228, 48, 48, 48, 48, 54 };
            table.TotalWidth = 525f;    
            table.SetWidths(tableWidths);
            table.LockedWidth = true;
            table.SpacingBefore = 0f;
            AddLoanTransactionTitle("Eff\nDate", Element.ALIGN_LEFT, 2, ref table);
            AddLoanTransactionTitle("Transaction Description", Element.ALIGN_LEFT, 1, ref table);
            AddLoanTransactionTitle("Amount", Element.ALIGN_RIGHT, 1, ref table);
            AddLoanTransactionTitle("Interest\nCharged", Element.ALIGN_RIGHT, 2, ref table);
            AddLoanTransactionTitle("Late\nFees", Element.ALIGN_RIGHT, 2, ref table);
            AddLoanTransactionTitle("Principal", Element.ALIGN_RIGHT, 1, ref table);
            AddLoanTransactionTitle("Balance\nSubject to\nInterest\nRate **", Element.ALIGN_RIGHT, 4, ref table);

            foreach (LoanTransaction transaction in loan.Transactions)
            {
                AddLoanTransaction(transaction, ref table);
            }

            if (loan.Transactions.Count() == 0)
            {
                AddNoTransactionsThisPeriodMessage(ref table);
            }

            if (loan.TotalFee.AmountThisPeriod > 0)
            {
                AddSeeFeeSummaryMessage(loan ,ref table);
            }

            if (loan.Closed)
            {
                AddLoanClosed(loan, ref table);
            }

            //AddEndingBalance(account, ref table);
            Doc.Add(table);
        }

        static void AddLoanTransaction(LoanTransaction transaction, ref PdfPTable table)
        {
            // Adds Date
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(transaction.PostingDate.ToString("MMM dd"), GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -6f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }
            // Adds Transaction Description
            {
                string description = string.Empty;

                if (transaction.DescriptionLines.Count > 0)
                {
                    description = transaction.DescriptionLines[0];
                }

                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(description, GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);

                // Adds additional description lines
                for (int i = 0; i < transaction.DescriptionLines.Count; i++)
                {
                    if (i > 0)
                    {
                        chunk = new Chunk(transaction.DescriptionLines[i], GetNormalFont(9f));
                        chunk.setLineHeight(11f);
                        chunk.SetCharacterSpacing(0f);
                        p = new Paragraph(chunk);
                        p.IndentationLeft = 20;
                        cell.AddElement(p);
                    }
                }

                cell.PaddingTop = -6f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            //if (transaction.Amount >= 0)
            //{
            //    AddAccountTransactionAmount(transaction.Amount, ref table); // Additions
            //    AddAccountTransactionAmount(0, ref table); // Subtractions
            //}
            //else
            //{
            //    AddAccountTransactionAmount(0, ref table); // Additions
            //    AddAccountTransactionAmount(transaction.Amount, ref table); // Subtractions
            //}

            AddLoanAccountTransactionAmount(transaction.Amount, ref table); // Amount
            AddLoanAccountTransactionAmount(transaction.InterestCharged, ref table); // Interest Charged
            AddLoanAccountTransactionAmount(transaction.LateFees, ref table); // Late Fees
            AddLoanAccountTransactionAmount(transaction.Principal, ref table); // Principal
            AddLoanAccountTransactionAmount(transaction.Balance, ref table); // Balance subject to interest rate **


            //AddAccountBalance(transaction.Balance, ref table);
        }

        static void AddNoTransactionsThisPeriodMessage(ref PdfPTable table)
        {
            // Adds Date
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -6f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }
            // Adds Transaction Description
            {

                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("No Transactions This Period", GetBoldItalicFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -6f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            for(int i = 0; i < 5; i++)
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                p.Alignment = Element.ALIGN_RIGHT;
                cell.AddElement(p);
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                cell.BorderColor = BaseColor.LIGHT_GRAY;
                table.AddCell(cell);
            }
        }

        static void AddSeeFeeSummaryMessage(Loan loan, ref PdfPTable table)
        {
            // Adds Date
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(loan.ClosingDateOfBillingCycle.ToString("MMM dd"), GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -6f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }
            // Adds Transaction Description
            {

                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("See Fee Summary Below", GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -6f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            for (int i = 0; i < 5; i++)
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                p.Alignment = Element.ALIGN_RIGHT;
                cell.AddElement(p);
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                cell.BorderColor = BaseColor.LIGHT_GRAY;
                table.AddCell(cell);
            }
        }

        static void AddLoanTransactionTitle(string title, int alignment, int numOfLines, ref PdfPTable table)
        {
            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(title, GetBoldFont(9f));
            chunk.SetCharacterSpacing(0f);
            chunk.setLineHeight(10f);
            //chunk.SetUnderline(0.75f, -2);
            Paragraph p = new Paragraph(chunk);
            p.Alignment = alignment;
            cell.AddElement(p);
            cell.PaddingTop = 6f;
            cell.BorderWidth = 0f;
            cell.BorderColor = BaseColor.GRAY;
            cell.VerticalAlignment = PdfPCell.ALIGN_BOTTOM;
            cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
            cell.AddElement(Underline(chunk, alignment, numOfLines));
            table.AddCell(cell);
        }

        static void AddLoanTransactionsFooter(string value)
        {
            PdfPTable table = new PdfPTable(1);
            table.TotalWidth = 525f;
            table.LockedWidth = true;
            table.SpacingBefore = 12f;
            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(value, GetBoldFont(9f));
            chunk.SetCharacterSpacing(0f);
            chunk.setLineHeight(11f);
            Paragraph p = new Paragraph(chunk);
            p.IndentationLeft = 15;
            cell.AddElement(p);
            cell.PaddingTop = -4f;
            cell.BorderWidth = 0f;
            table.AddCell(cell);
            Doc.Add(table);
        }

        static void AddFeeSummary(Loan loan)
        {
            if (Writer.GetVerticalPosition(false) <= 130)
            {
                Doc.NewPage();
            }

            // Title
            {
                PdfPTable table = new PdfPTable(1);
                table.TotalWidth = 525f;
                table.LockedWidth = true;
                table.SpacingBefore = 12f;
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("FEE SUMMARY", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                chunk.setLineHeight(11f);
                Paragraph p = new Paragraph(chunk);
                //p.IndentationLeft = 15;
                cell.AddElement(p);
                cell.PaddingTop = -4f;
                cell.PaddingLeft = 10f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
                Doc.Add(table);
            }

            AddLoanFees(loan);

            // TOTAL FEES FOR THIS PERIOD
            {
                PdfPTable table = new PdfPTable(4);
                table.HeaderRows = 0;
                float[] tableWidths = new float[] { 51, 234, 48, 192 };
                table.TotalWidth = 525f;
                table.SetWidths(tableWidths);
                table.LockedWidth = true;
                table.SpacingBefore = 0f;

                // For layout only
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    p.Alignment = Element.ALIGN_RIGHT;
                    cell.AddElement(p);
                    cell.PaddingTop = -4f;
                    cell.BorderWidth = 0f;
                    cell.BorderColor = BaseColor.LIGHT_GRAY;
                    table.AddCell(cell);
                }

                // TOTAL FEES FOR THIS PERIOD
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("TOTAL FEES FOR THIS PERIOD", GetNormalFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    p.Alignment = Element.ALIGN_LEFT;
                    cell.AddElement(p);
                    cell.PaddingTop = -4f;
                    cell.BorderWidth = 0f;
                    cell.BorderColor = BaseColor.LIGHT_GRAY;
                    table.AddCell(cell);
                }

                // value
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk(loan.TotalFee.AmountThisPeriod.ToString("N"), GetNormalFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    p.Alignment = Element.ALIGN_RIGHT;
                    cell.AddElement(p);
                    cell.PaddingTop = -4f;
                    cell.BorderWidth = 0f;
                    cell.BorderColor = BaseColor.LIGHT_GRAY;
                    table.AddCell(cell);
                }

                // For layout only
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    p.Alignment = Element.ALIGN_RIGHT;
                    cell.AddElement(p);
                    cell.PaddingTop = -4f;
                    cell.BorderWidth = 0f;
                    cell.BorderColor = BaseColor.LIGHT_GRAY;
                    table.AddCell(cell);
                }

                Doc.Add(table);
            }
        }

        static void AddInterestChargedSummary(Loan loan)
        {
            if (Writer.GetVerticalPosition(false) <= 130)
            {
                Doc.NewPage();
            }

            // Title
            {
                PdfPTable table = new PdfPTable(1);
                table.TotalWidth = 525f;
                table.LockedWidth = true;
                table.SpacingBefore = 12f;
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("INTEREST CHARGED SUMMARY", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                chunk.setLineHeight(11f);
                Paragraph p = new Paragraph(chunk);
                //p.IndentationLeft = 15;
                cell.AddElement(p);
                cell.PaddingTop = -4f;
                cell.PaddingLeft = 10f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
                Doc.Add(table);
            }

            AddLoanInterestTransactions(loan);

            // TOTAL INTEREST FOR THIS PERIOD
            {
                PdfPTable table = new PdfPTable(4);
                table.HeaderRows = 0;
                float[] tableWidths = new float[] { 51, 234, 48, 192 };
                table.TotalWidth = 525f;
                table.SetWidths(tableWidths);
                table.LockedWidth = true;
                table.SpacingBefore = 0f;

                // For layout only
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    p.Alignment = Element.ALIGN_RIGHT;
                    cell.AddElement(p);
                    cell.PaddingTop = -4f;
                    cell.BorderWidth = 0f;
                    cell.BorderColor = BaseColor.LIGHT_GRAY;
                    table.AddCell(cell);
                }

                // TOTAL INTEREST FOR THIS PERIOD
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("TOTAL INTEREST FOR THIS PERIOD", GetNormalFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    p.Alignment = Element.ALIGN_LEFT;
                    cell.AddElement(p);
                    cell.PaddingTop = -4f;
                    cell.BorderWidth = 0f;
                    cell.BorderColor = BaseColor.LIGHT_GRAY;
                    table.AddCell(cell);
                }

                // value
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk(loan.TotalInterestChargedThisPeriod.ToString("N"), GetNormalFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    p.Alignment = Element.ALIGN_RIGHT;
                    cell.AddElement(p);
                    cell.PaddingTop = -4f;
                    cell.BorderWidth = 0f;
                    cell.BorderColor = BaseColor.LIGHT_GRAY;
                    table.AddCell(cell);
                }

                // For layout only
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    p.Alignment = Element.ALIGN_RIGHT;
                    cell.AddElement(p);
                    cell.PaddingTop = -4f;
                    cell.BorderWidth = 0f;
                    cell.BorderColor = BaseColor.LIGHT_GRAY;
                    table.AddCell(cell);
                }

                Doc.Add(table);
            }
        }

        static void AddLoanFees(Loan loan)
        {
            PdfPTable table = new PdfPTable(4);
            table.HeaderRows = 1;
            float[] tableWidths = new float[] { 51, 234, 48, 192 };
            table.TotalWidth = 525f;
            table.SetWidths(tableWidths);
            table.LockedWidth = true;
            table.SpacingBefore = 0f;
            AddLoanTransactionTitle("Date", Element.ALIGN_LEFT, 1, ref table);
            AddLoanTransactionTitle("Description", Element.ALIGN_LEFT, 1, ref table);
            AddLoanTransactionTitle("Amount", Element.ALIGN_RIGHT, 1, ref table);
            AddLoanTransactionTitle(string.Empty, Element.ALIGN_RIGHT, 1, ref table);

            foreach (Fee fee in loan.Fees)
            {
                AddLoanFee(fee, ref table);
            }

            Doc.Add(table);
        }

        static void AddLoanFee(Fee fee, ref PdfPTable table)
        {
            // Adds Date
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(fee.Date.ToString("MMM dd"), GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -6f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }
            // Adds Transaction Description
            {
                //string description = fee.Description.Length > 30 ? fee.Description.Substring(0, 30) : fee.Description;
                string description = fee.Description;

                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(description, GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -6f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            AddLoanAccountTransactionAmount(fee.AmountThisPeriod, ref table); // Amount

            // For layout only
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -6f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            //AddAccountBalance(transaction.Balance, ref table);
        }

        static void AddLoanInterestTransactions(Loan loan)
        {
            PdfPTable table = new PdfPTable(4);
            table.HeaderRows = 1;
            float[] tableWidths = new float[] { 51, 234, 48, 192 };
            table.TotalWidth = 525f;
            table.SetWidths(tableWidths);
            table.LockedWidth = true;
            table.SpacingBefore = 0f;
            AddLoanTransactionTitle("Date", Element.ALIGN_LEFT, 1, ref table);
            AddLoanTransactionTitle("Description", Element.ALIGN_LEFT, 1, ref table);
            AddLoanTransactionTitle("Amount", Element.ALIGN_RIGHT, 1, ref table);
            AddLoanTransactionTitle(string.Empty, Element.ALIGN_RIGHT, 1, ref table);

            foreach (LoanTransaction transaction in loan.Transactions)
            {
                if (transaction.InterestCharged > 0)
                {
                    AddLoanInterestTransaction(transaction, ref table);
                }
            }

            Doc.Add(table);
        }

        static void AddLoanInterestTransaction(LoanTransaction transaction, ref PdfPTable table)
        {
            // Adds Date
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(transaction.PostingDate.ToString("MMM dd"), GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -6f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }
            // Adds Transaction Description
            {
                string description = transaction.DescriptionLine1.Length > 30 ? transaction.DescriptionLine1.Substring(0, 30) : transaction.DescriptionLine1;

                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(description, GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -6f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            AddLoanAccountTransactionAmount(transaction.InterestCharged, ref table); // Amount

            // For layout only
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -6f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            //AddAccountBalance(transaction.Balance, ref table);
        }

        static void AddYearToDateTotals(Loan loan)
        {
            PdfPTable table = new PdfPTable(5);
            table.HeaderRows = 1;
            float[] tableWidths = new float[] { 12, 153, 79, 93, 188 };
            table.TotalWidth = 525f;
            table.SetWidths(tableWidths);
            table.LockedWidth = true;
            table.SpacingBefore = 20f;
            table.KeepTogether = true;

            // For layout only
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(string.Empty, GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.Padding = 0;
                cell.PaddingTop = 0.5f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            // Title
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("YEAR TO DATE TOTALS", GetBoldFont(9));
                chunk.SetCharacterSpacing(0f);
                chunk.setLineHeight(11f);
                Paragraph p = new Paragraph(chunk);
                p.Alignment = Element.ALIGN_LEFT;
                cell.AddElement(p);
                cell.Padding = 0;
                cell.PaddingTop = 0.5f;
                cell.PaddingLeft = 6f;
                cell.BorderWidth = 0f;
                cell.BorderWidthTop = 1f;
                cell.BorderWidthLeft = 1f;
                cell.BorderColor = BaseColor.BLACK;
                table.AddCell(cell);
            }

            // For layout only
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(string.Empty, GetBoldFont(9));
                chunk.SetCharacterSpacing(0f);
                chunk.setLineHeight(11f);
                Paragraph p = new Paragraph(chunk);
                p.Alignment = Element.ALIGN_LEFT;
                cell.AddElement(p);
                cell.Padding = 0;
                cell.PaddingTop = 0.5f;
                cell.BorderWidth = 0f;
                cell.BorderWidthTop = 1f;
                cell.BorderColor = BaseColor.BLACK;
                table.AddCell(cell);
            }

            // For layout only
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(string.Empty, GetBoldFont(9));
                chunk.SetCharacterSpacing(0f);
                chunk.setLineHeight(11f);
                Paragraph p = new Paragraph(chunk);
                p.Alignment = Element.ALIGN_LEFT;
                cell.AddElement(p);
                cell.Padding = 0;
                cell.PaddingTop = 0.5f;
                cell.BorderWidth = 0f;
                cell.BorderWidthTop = 1f;
                cell.BorderWidthRight = 1f;
                cell.BorderColor = BaseColor.BLACK;
                table.AddCell(cell);
            }

            // For layout only
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.Padding = 0;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            // For layout only
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.Padding = 0;
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            // Total Fees Charged this Year Title
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("Total Fees Charged this Year", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.Padding = 0;
                cell.PaddingTop = -4f;
                cell.PaddingLeft = 6f;
                cell.BorderWidth = 0f;
                cell.BorderWidthLeft = 1f;
                cell.BorderColor = BaseColor.BLACK;
                table.AddCell(cell);
            }

            // Total Fees Charged this Year Value
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(FormatAmount(loan.TotalFee.AmountYtd), GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                p.Alignment = Element.ALIGN_RIGHT;
                cell.AddElement(p);
                cell.Padding = 0;
                cell.PaddingTop = -4f;
                cell.PaddingRight = 35f;
                cell.BorderWidth = 0f;
                cell.BorderColor = BaseColor.BLACK;
                table.AddCell(cell);
            }

            // For layout only
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.Padding = 0;
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                cell.BorderWidthRight = 1f;
                cell.BorderColor = BaseColor.BLACK;
                table.AddCell(cell);
            }

            // For layout only
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.Padding = 0;
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            // For layout only
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.Padding = 0;
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            // Total Interest Charged this Year Title
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("Total Interest Charged this Year", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.Padding = 0;
                cell.PaddingTop = -4f;
                cell.PaddingBottom = (loan.ExistedLastYear) ? 0f : 5f;
                cell.PaddingLeft = 6f;
                cell.BorderWidth = 0f;
                cell.BorderWidthBottom = (loan.ExistedLastYear) ? 0f : 1f;
                cell.BorderWidthLeft = 1f;
                cell.BorderColor = BaseColor.BLACK;
                table.AddCell(cell);
            }

            // Total Interest Charged this Year Value
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(FormatAmount(loan.TotalInterestChargedYtd), GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                p.Alignment = Element.ALIGN_RIGHT;
                cell.AddElement(p);
                cell.Padding = 0;
                cell.PaddingTop = -4f;
                cell.PaddingBottom = (loan.ExistedLastYear) ? 0f : 5f;
                cell.PaddingRight = 35f;
                cell.BorderWidth = 0f;
                cell.BorderWidthBottom = (loan.ExistedLastYear) ? 0f : 1f;
                cell.BorderColor = BaseColor.BLACK;
                table.AddCell(cell);
            }

            // For layout only
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.Padding = 0;
                cell.PaddingTop = -4f;
                cell.PaddingBottom = (loan.ExistedLastYear) ? 0f : 5f;
                cell.BorderWidth = 0f;
                cell.BorderWidthBottom = (loan.ExistedLastYear) ? 0f : 1f;
                cell.BorderWidthRight = 1f;
                cell.BorderColor = BaseColor.BLACK;
                table.AddCell(cell);
            }

            // For layout only
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.Padding = 0;
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            if (loan.ExistedLastYear)
            {
                // For layout only
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("", GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.Padding = 0;
                    cell.PaddingTop = -4f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                }

                // Total Fees Charged Last Year
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("Total Fees Charged Last Year", GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.Padding = 0;
                    cell.PaddingTop = -4f;
                    cell.PaddingLeft = 6f;
                    cell.BorderWidth = 0f;
                    cell.BorderWidthBottom = 0f;
                    cell.BorderWidthLeft = 1f;
                    cell.BorderColor = BaseColor.BLACK;
                    table.AddCell(cell);
                }

                // Total Interest Charged this Year Value
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk(FormatAmount(loan.TotalFee.AmountLastYear), GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    p.Alignment = Element.ALIGN_RIGHT;
                    cell.AddElement(p);
                    cell.Padding = 0;
                    cell.PaddingTop = -4f;
                    cell.PaddingRight = 35f;
                    cell.BorderWidth = 0f;
                    cell.BorderWidthBottom = 0f;
                    cell.BorderColor = BaseColor.BLACK;
                    table.AddCell(cell);
                }

                // For layout only
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("", GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.Padding = 0;
                    cell.PaddingTop = -4f;
                    cell.BorderWidth = 0f;
                    cell.BorderWidthBottom = 0f;
                    cell.BorderWidthRight = 1f;
                    cell.BorderColor = BaseColor.BLACK;
                    table.AddCell(cell);
                }

                // For layout only
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("", GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.Padding = 0;
                    cell.PaddingTop = -4f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                }

                // For layout only
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("", GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.Padding = 0;
                    cell.PaddingTop = -4f;
                    cell.PaddingBottom = 5f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                }

                // Total Interest Charged Last Year
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("Total Interest Charged Last Year", GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.Padding = 0;
                    cell.PaddingTop = -4f;
                    cell.PaddingLeft = 6f;
                    cell.PaddingBottom = 5f;
                    cell.BorderWidth = 0f;
                    cell.BorderWidthBottom = 1f;
                    cell.BorderWidthLeft = 1f;
                    cell.BorderColor = BaseColor.BLACK;
                    table.AddCell(cell);
                }

                // Total Interest Charged Last Year Value
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk(FormatAmount(loan.TotalInterestChargedLastYear), GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    p.Alignment = Element.ALIGN_RIGHT;
                    cell.AddElement(p);
                    cell.Padding = 0;
                    cell.PaddingTop = -4f;
                    cell.PaddingBottom = 5f;
                    cell.PaddingRight = 35f;
                    cell.BorderWidth = 0f;
                    cell.BorderWidthBottom = 1f;
                    cell.BorderColor = BaseColor.BLACK;
                    table.AddCell(cell);
                }

                // For layout only
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("", GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.Padding = 0;
                    cell.PaddingTop = -4f;
                    cell.PaddingBottom = 5f;
                    cell.BorderWidth = 0f;
                    cell.BorderWidthBottom = 1f;
                    cell.BorderWidthRight = 1f;
                    cell.BorderColor = BaseColor.BLACK;
                    table.AddCell(cell);
                }

                // For layout only
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("", GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.Padding = 0;
                    cell.PaddingTop = -4f;
                    cell.PaddingBottom = 5f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                }
            }

            

            Doc.Add(table);
        }

        static void AddYtdSummaries(Statement statement)
        {
            PdfPTable leftTable;
            PdfPCell leftTableCell;

            AddSectionHeading("YTD SUMMARIES");

            // A table to create 2 columns
            {
                PdfPTable table = new PdfPTable(2);
                table.HeaderRows = 0;
                float[] tableWidths = new float[] { 255, 270 };
                table.TotalWidth = 525f;
                table.LockedWidth = true;
                table.SpacingBefore = 15f;

                // TOTAL DIVIDENDS PAID
                {
                    leftTable = new PdfPTable(2);
                    leftTable.HeaderRows = 0;
                    float[] leftTableWidths = new float[] { 214.5f, 48f };
                    leftTable.TotalWidth = 262.5f;
                    leftTable.SetWidths(leftTableWidths);
                    leftTable.LockedWidth = true;
                    leftTable.SpacingBefore = 0f;

                    // Title
                    {
                        PdfPCell cell = new PdfPCell();
                        Chunk chunk = new Chunk("TOTAL DIVIDENDS PAID", GetBoldFont(9f));
                        chunk.SetCharacterSpacing(0f);
                        Paragraph p = new Paragraph(chunk);
                        p.IndentationLeft = 20f;
                        cell.AddElement(p);
                        cell.PaddingTop = -6f;
                        cell.BorderWidth = 0f;
                        leftTable.AddCell(cell);
                    }

                    // For layout only
                    {
                        PdfPCell cell = new PdfPCell();
                        Chunk chunk = new Chunk("", GetBoldFont(9f));
                        chunk.SetCharacterSpacing(0f);
                        Paragraph p = new Paragraph(chunk);
                        cell.AddElement(p);
                        cell.PaddingTop = -6f;
                        cell.BorderWidth = 0f;
                        leftTable.AddCell(cell);
                    }

                    foreach (Account account in statement.Accounts.OrderBy(o => o.Description).ToList())
                    {
                        // Account
                        {
                            PdfPCell cell = new PdfPCell();
                            Chunk chunk = new Chunk(account.Description, GetNormalFont(9f));
                            chunk.SetCharacterSpacing(0f);
                            Paragraph p = new Paragraph(chunk);
                            p.IndentationLeft = 20f;
                            cell.AddElement(p);
                            cell.PaddingTop = -6f;
                            cell.BorderWidth = 0f;
                            leftTable.AddCell(cell);
                        }

                        // Value
                        {
                            PdfPCell cell = new PdfPCell();
                            Chunk chunk = new Chunk(FormatAmount(account.Dividends), GetNormalFont(9f));
                            chunk.SetCharacterSpacing(0f);
                            Paragraph p = new Paragraph(chunk);
                            p.Alignment = Element.ALIGN_RIGHT;
                            cell.AddElement(p);
                            cell.PaddingTop = -6f;
                            cell.BorderWidth = 0f;
                            leftTable.AddCell(cell);
                        }
                    }

                    foreach (IrsContribution irsContributionYtd in statement.IrsContributionsYtd)
                    {
                        // Total
                        {
                            PdfPCell cell = new PdfPCell();
                            Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));
                            chunk = new Chunk(irsContributionYtd.Description, GetNormalFont(9f));
                            chunk.SetCharacterSpacing(0f);
                            Paragraph p = new Paragraph(chunk);
                            p.IndentationLeft = 20f;
                            cell.AddElement(p);
                            cell.PaddingTop = -6f;
                            cell.BorderWidth = 0f;
                            leftTable.AddCell(cell);
                        }

                        // Value
                        {
                            PdfPCell cell = new PdfPCell();
                            Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));
                            chunk = new Chunk(FormatAmount(irsContributionYtd.Amount), GetNormalFont(9f));
                            chunk.SetCharacterSpacing(0f);
                            Paragraph p = new Paragraph(chunk);
                            p.Alignment = Element.ALIGN_RIGHT;
                            cell.AddElement(p);
                            cell.PaddingTop = -6f;
                            cell.BorderWidth = 0f;
                            leftTable.AddCell(cell);
                        }
                    }

                    // Total
                    {
                        PdfPCell cell = new PdfPCell();
                        Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));

                        if (statement.DividendsTotalSet)
                        {
                            chunk = new Chunk("Total Year To Date Dividends Paid", GetNormalFont(9f));
                        }

                        chunk.SetCharacterSpacing(0f);
                        Paragraph p = new Paragraph(chunk);
                        p.IndentationLeft = 20f;
                        cell.AddElement(p);
                        cell.PaddingTop = -6f;
                        cell.BorderWidth = 0f;
                        leftTable.AddCell(cell);
                    }

                    // Value
                    {
                        PdfPCell cell = new PdfPCell();
                        Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));

                        if (statement.DividendsTotalSet)
                        {
                            chunk = new Chunk(FormatAmount(statement.DividendsTotal), GetNormalFont(9f));
                        }
                        chunk.SetCharacterSpacing(0f);
                        Paragraph p = new Paragraph(chunk);
                        p.Alignment = Element.ALIGN_RIGHT;
                        cell.AddElement(p);
                        cell.PaddingTop = -6f;
                        cell.BorderWidth = 0f;
                        leftTable.AddCell(cell);
                    }

                    // Total
                    {
                        PdfPCell cell = new PdfPCell();
                        Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));

                        if (statement.NontaxableDividendYtdSet)
                        {
                            chunk = new Chunk("Nontaxable Dividends", GetNormalFont(9f));
                        }

                        chunk.SetCharacterSpacing(0f);
                        Paragraph p = new Paragraph(chunk);
                        p.IndentationLeft = 20f;
                        cell.AddElement(p);
                        cell.PaddingTop = -6f;
                        cell.BorderWidth = 0f;
                        leftTable.AddCell(cell);
                    }

                    // Value
                    {
                        PdfPCell cell = new PdfPCell();
                        Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));
                        if (statement.NontaxableDividendYtdSet)
                        {
                            chunk = new Chunk(FormatAmount(statement.NontaxableDividendYtd), GetNormalFont(9f));
                        }
                        chunk.SetCharacterSpacing(0f);
                        Paragraph p = new Paragraph(chunk);
                        p.Alignment = Element.ALIGN_RIGHT;
                        cell.AddElement(p);
                        cell.PaddingTop = -6f;
                        cell.BorderWidth = 0f;
                        leftTable.AddCell(cell);
                    }

                    foreach (YtdTotal ytdTotal in statement.YtdTotals)
                    {
                        // Total
                        {
                            PdfPCell cell = new PdfPCell();
                            Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));
                            chunk = new Chunk(ytdTotal.Description, GetNormalFont(9f));
                            chunk.SetCharacterSpacing(0f);
                            Paragraph p = new Paragraph(chunk);
                            p.IndentationLeft = 20f;
                            cell.AddElement(p);
                            cell.PaddingTop = -6f;
                            cell.BorderWidth = 0f;
                            leftTable.AddCell(cell);
                        }

                        // Value
                        {
                            PdfPCell cell = new PdfPCell();
                            Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));
                            chunk = new Chunk(FormatAmount(ytdTotal.Amount), GetNormalFont(9f));
                            chunk.SetCharacterSpacing(0f);
                            Paragraph p = new Paragraph(chunk);
                            p.Alignment = Element.ALIGN_RIGHT;
                            cell.AddElement(p);
                            cell.PaddingTop = -6f;
                            cell.BorderWidth = 0f;
                            leftTable.AddCell(cell);
                        }
                    }

                    // Total
                    {
                        PdfPCell cell = new PdfPCell();
                        Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));

                        if (statement.InterestPaidTotalYtdSet)
                        {
                            chunk = new Chunk("Total Year To Date Interest Paid", GetNormalFont(9f));
                        }

                        chunk.SetCharacterSpacing(0f);
                        Paragraph p = new Paragraph(chunk);
                        p.IndentationLeft = 20f;
                        cell.AddElement(p);
                        cell.PaddingTop = -6f;
                        cell.BorderWidth = 0f;
                        leftTable.AddCell(cell);
                    }

                    // Value
                    {
                        PdfPCell cell = new PdfPCell();
                        Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));
                        if (statement.InterestPaidTotalYtdSet)
                        {
                            chunk = new Chunk(FormatAmount(statement.InterestPaidTotalYtd), GetNormalFont(9f));
                        }
                        chunk.SetCharacterSpacing(0f);
                        Paragraph p = new Paragraph(chunk);
                        p.Alignment = Element.ALIGN_RIGHT;
                        cell.AddElement(p);
                        cell.PaddingTop = -6f;
                        cell.BorderWidth = 0f;
                        leftTable.AddCell(cell);
                    }

                    leftTableCell = new PdfPCell();
                    if(statement.Accounts.Count > 0) leftTableCell.AddElement(leftTable);
                    leftTableCell.BorderWidth = 0;
                    leftTableCell.Padding = 0;

                    table.AddCell(leftTableCell);
                }

                // TOTAL LOAN INTEREST PAID
                {
                    leftTable = new PdfPTable(2);
                    leftTable.HeaderRows = 0;
                    float[] leftTableWidths = new float[] { 214.5f, 48f };
                    leftTable.TotalWidth = 262.5f;
                    leftTable.SetWidths(leftTableWidths);
                    leftTable.LockedWidth = true;
                    leftTable.SpacingBefore = 0f;

                    // Title
                    {
                        PdfPCell cell = new PdfPCell();
                        Chunk chunk = new Chunk("TOTAL LOAN INTEREST PAID", GetBoldFont(9f));
                        chunk.SetCharacterSpacing(0f);
                        Paragraph p = new Paragraph(chunk);
                        p.IndentationLeft = 20f;
                        cell.AddElement(p);
                        cell.PaddingTop = -6f;
                        cell.BorderWidth = 0f;
                        leftTable.AddCell(cell);
                    }

                    // For layout only
                    {
                        PdfPCell cell = new PdfPCell();
                        Chunk chunk = new Chunk("", GetBoldFont(9f));
                        chunk.SetCharacterSpacing(0f);
                        Paragraph p = new Paragraph(chunk);
                        cell.AddElement(p);
                        cell.PaddingTop = -6f;
                        cell.BorderWidth = 0f;
                        leftTable.AddCell(cell);
                    }


                    foreach (Loan loan in statement.Loans.OrderBy(o => o.Description).ToList())
                    {
                        // Loan
                        {
                            PdfPCell cell = new PdfPCell();
                            Chunk chunk = new Chunk(loan.Description, GetNormalFont(9f));
                            chunk.SetCharacterSpacing(0f);
                            Paragraph p = new Paragraph(chunk);
                            p.IndentationLeft = 20f;
                            cell.AddElement(p);
                            cell.PaddingTop = -6f;
                            cell.BorderWidth = 0f;
                            leftTable.AddCell(cell);
                        }

                        // Value
                        {
                            PdfPCell cell = new PdfPCell();
                            Chunk chunk = new Chunk(FormatAmount(loan.TotalInterestChargedYtd), GetNormalFont(9f));
                            chunk.SetCharacterSpacing(0f);
                            Paragraph p = new Paragraph(chunk);
                            p.Alignment = Element.ALIGN_RIGHT;
                            cell.AddElement(p);
                            cell.PaddingTop = -6f;
                            cell.BorderWidth = 0f;
                            leftTable.AddCell(cell);
                        }
                    }

                    leftTableCell = new PdfPCell();
                    if(statement.Loans.Count > 0) leftTableCell.AddElement(leftTable);
                    leftTableCell.BorderWidth = 0;
                    leftTableCell.Padding = 0;

                    table.AddCell(leftTableCell);
                }

 

                Doc.Add(table);
            }
        }

        static void AddMoneyPerksSummary(MoneyPerksStatement moneyPerksStatement)
        {
            if (moneyPerksStatement!=null)
            {
                AddSectionHeading("MONEYPERKS POINTS SUMMARY");

                PdfPTable table = new PdfPTable(5);
                table.HeaderRows = 1;
                float[] tableWidths = new float[] { 51, 280, 62, 65, 67 };
                table.TotalWidth = 525f;
                table.SetWidths(tableWidths);
                table.LockedWidth = true;
                AddMoneyPerksTransactionTitle("Date", Element.ALIGN_LEFT, 1, ref table);
                AddMoneyPerksTransactionTitle("Transaction Description", Element.ALIGN_LEFT, 1, ref table);
                AddMoneyPerksTransactionTitle("Points\nAwarded", Element.ALIGN_RIGHT, 2,  ref table);
                AddMoneyPerksTransactionTitle("Points\nRedeemed", Element.ALIGN_RIGHT, 2, ref table);
                AddMoneyPerksTransactionTitle("Balance", Element.ALIGN_RIGHT, 1,  ref table);
                AddMoneyPerksBeginningBalance(moneyPerksStatement, ref table);

                foreach (MoneyPerksTransaction transaction in moneyPerksStatement.Transactions)
                {
                    AddMoneyPerksTransaction(transaction, ref table);
                }

                AddMoneyPerksEndingBalance(moneyPerksStatement, ref table);
                Doc.Add(table);
            }
        }

        static void AddMoneyPerksTransactionTitle(string title, int alignment, int numOfLines, ref PdfPTable table)
        {
            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(title, GetBoldFont(9f));
            chunk.SetCharacterSpacing(0f);
            chunk.setLineHeight(10f);
            //chunk.SetUnderline(0.75f, -2);
            Paragraph p = new Paragraph(chunk);
            p.Alignment = alignment;
            cell.AddElement(p);
            cell.PaddingTop = 6f;
            cell.BorderWidth = 0f;
            cell.BorderColor = BaseColor.GRAY;
            cell.VerticalAlignment = PdfPCell.ALIGN_BOTTOM;
            cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
            cell.AddElement(Underline(chunk, alignment, numOfLines));
            table.AddCell(cell);
        }

        static void AddMoneyPerksBeginningBalance(MoneyPerksStatement moneyPerksStatement, ref PdfPTable table)
        {
            // Adds Date
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(string.Empty, GetBoldItalicFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);

            }
            // Adds Transaction Description
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("Beginning Balance", GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            AddMoneyPerksTransactionAmount(0, ref table);
            AddMoneyPerksTransactionAmount(0, ref table);
            AddMoneyPerksBalance(moneyPerksStatement.BeginningBalance, ref table);
        }

        static void AddMoneyPerksEndingBalance(MoneyPerksStatement moneyPerksStatement, ref PdfPTable table)
        {
            // Adds Date
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(string.Empty, GetBoldItalicFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);

            }
            // Adds Transaction Description
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("Ending Balance", GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            AddMoneyPerksTransactionAmount(0, ref table);
            AddMoneyPerksTransactionAmount(0, ref table);
            AddMoneyPerksBalance(moneyPerksStatement.EndingBalance, ref table);
        }

        static void AddMoneyPerksTransaction(MoneyPerksTransaction transaction, ref PdfPTable table)
        {
            // Adds Date
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(transaction.Date.ToString("MMM dd"), GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }
            // Adds Transaction Description
            {
                string description = transaction.Description;

                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(description, GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            if (transaction.Amount >= 0)
            {
                AddMoneyPerksTransactionAmount(transaction.Amount, ref table); // Additions
                AddMoneyPerksTransactionAmount(0, ref table); // Subtractions
            }
            else
            {
                AddMoneyPerksTransactionAmount(0, ref table); // Additions
                AddMoneyPerksTransactionAmount(transaction.Amount, ref table); // Subtractions
            }

            AddMoneyPerksBalance(transaction.Balance, ref table);
        }

        static void AddTopAdvertising(Statement statement)
        {
            // Advertisement Bottom
            Font font = new Font(Font.FontFamily.HELVETICA, 12f, Font.BOLD, new BaseColor(0, 0, 0));
            PdfPTable table = new PdfPTable(1);
            table.TotalWidth = 525f;
            table.LockedWidth = true;
            table.SpacingBefore = 34f;
            PdfPCell cell = new PdfPCell();
            Chunk chunk = null;

            for (int i = 0; i < statement.AdvertisementTop.TotalLines; i++)
            {
                if (chunk == null)
                {
                    chunk = new Chunk(statement.AdvertisementTop.MessageLines[i], font);
                }
                else
                {
                    chunk.Append("\n" + statement.AdvertisementTop.MessageLines[i]);
                }
            }

            if (chunk == null)
            {
                chunk = new Chunk(string.Empty, font);
                table.SpacingBefore = 14f;
            }

            chunk.SetCharacterSpacing(0f);
            chunk.setLineHeight(12f);
            Paragraph p = new Paragraph(chunk);
            p.Alignment = Element.ALIGN_CENTER;
            //p.IndentationLeft = 385;
            cell.AddElement(p);
            //cell.PaddingTop = -1f;
            cell.BorderWidth = 0f;
            cell.BorderColor = BaseColor.LIGHT_GRAY;
            table.AddCell(cell);

            Doc.Add(table);
        }

        static void AddBottomAdvertising(Statement statement)
        {
            if (Writer.GetVerticalPosition(false) <= 130)
            {
                Doc.NewPage();
            }

            // Advertisement Bottom Stroke
            PdfPTable table = new PdfPTable(1);
            table.TotalWidth = 525f;
            table.LockedWidth = true;
            table.SpacingBefore = 10f;
            PdfPCell cell = new PdfPCell();
            cell.BorderWidthBottom = 5f;
            cell.BorderColor = BaseColor.BLACK;
            table.AddCell(cell);
            Doc.Add(table);

            // Advertisement Bottom
            Font font = new Font(Font.FontFamily.HELVETICA, 12f, Font.BOLD, new BaseColor(0, 0, 0));
            table = new PdfPTable(1);
            table.TotalWidth = 525f;
            table.LockedWidth = true;
            table.SpacingBefore = 10f;
            cell = new PdfPCell();
            Chunk chunk = null;

            for (int i = 0; i < statement.AdvertisementBottom.TotalLines; i++)
            {
                if (chunk == null)
                {
                    chunk = new Chunk(statement.AdvertisementBottom.MessageLines[i], font);
                }
                else
                {
                    chunk.Append("\n" + statement.AdvertisementBottom.MessageLines[i]);
                }
            }

            if (chunk != null)
            {
                chunk.SetCharacterSpacing(0f);
                chunk.setLineHeight(12f);
                Paragraph p = new Paragraph(chunk);
                p.Alignment = Element.ALIGN_CENTER;
                //p.IndentationLeft = 385;
                cell.AddElement(p);
                //cell.PaddingTop = -1f;
                cell.BorderWidth = 0f;
                cell.BorderColor = BaseColor.LIGHT_GRAY;
                table.AddCell(cell);
                Doc.Add(table);
            }
        }

        static void AddPageNumbersAndDisclosures(Statement statement)
        {
            // Adds page numbers
            PdfReader statementReader = new PdfReader("C:\\" + TEMP_FILE_NAME);
            PdfReader statementBackReader = new PdfReader(Configuration.GetStatementDisclosuresTemplateFilePath());

            using (FileStream fs = new FileStream(Configuration.GetStatementsOutputPath() + statement.AccountNumber + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None))
            {
                using (PdfStamper stamper = new PdfStamper(statementReader, fs))
                {
                    stamper.SetFullCompression();
                    int pageCount = statementReader.NumberOfPages + 1; // Adds 1 for the disclosures page that will be added later
                    for (int i = 1; i <= pageCount - 1; i++)
                    {
                        if (i == 1)
                        {
                            // Page count on first page
                            Chunk chunk = new Chunk("Page:   1 of   " + pageCount, GetBoldFont(12f));
                            ColumnText.ShowTextAligned(stamper.GetOverContent(i), Element.ALIGN_RIGHT, new Phrase(chunk), 578, 595, 0);
                            if (i != pageCount)
                            {
                                chunk = new Chunk("--- Continued on following page ---", GetBoldFont(9f));
                                ColumnText.ShowTextAligned(stamper.GetOverContent(i), Element.ALIGN_CENTER, new Phrase(chunk), 300, 20, 0);
                            }
                        }
                        else if (i != pageCount)
                        {
                            float startY = 750f;
                            float lineHeight = 10;

                            ColumnText.ShowTextAligned(stamper.GetOverContent(i), Element.ALIGN_RIGHT, new Phrase(new Chunk(statement.BeginDate.ToString("MMM dd, yyyy") + "  thru  " + statement.EndDate.ToString("MMM dd, yyyy"), GetBoldFont(9f))), 578, startY, 0);
                            ColumnText.ShowTextAligned(stamper.GetOverContent(i), Element.ALIGN_RIGHT, new Phrase(new Chunk("Account  Number:   ******" + statement.AccountNumber.Substring("******".Length), GetBoldFont(9f))), 578, startY - lineHeight, 0);
                            ColumnText.ShowTextAligned(stamper.GetOverContent(i), Element.ALIGN_RIGHT, new Phrase(new Chunk("Page:  " + i.ToString() + " of " + pageCount.ToString(), GetBoldFont(9f))), 578, startY - (lineHeight * 2), 0);
                            Chunk chunk = new Chunk("--- Continued on reverse side ---", GetBoldFont(9f));
                            ColumnText.ShowTextAligned(stamper.GetOverContent(i), Element.ALIGN_CENTER, new Phrase(chunk), 300, 20, 0);
                        }
                        else
                        {
                            float startY = 750f;
                            float lineHeight = 10;
                            ColumnText.ShowTextAligned(stamper.GetOverContent(i), Element.ALIGN_RIGHT, new Phrase(new Chunk(statement.BeginDate.ToString("MMM dd, yyyy") + "  thru  " + statement.EndDate.ToString("MMM dd, yyyy"), GetBoldFont(9f))), 578, startY, 0);
                            ColumnText.ShowTextAligned(stamper.GetOverContent(i), Element.ALIGN_RIGHT, new Phrase(new Chunk("Account  Number:   ******" + statement.AccountNumber.Substring("******".Length), GetBoldFont(9f))), 578, startY - lineHeight, 0);
                            ColumnText.ShowTextAligned(stamper.GetOverContent(i), Element.ALIGN_RIGHT, new Phrase(new Chunk("Page:  " + i.ToString() + " of " + pageCount.ToString(), GetBoldFont(9f))), 578, startY - (lineHeight * 2), 0);
                        }
                    }

                    stamper.InsertPage(pageCount, PageSize.LETTER);
                    PdfContentByte cb = stamper.GetOverContent(pageCount);
                    PdfImportedPage p = stamper.GetImportedPage(statementBackReader, 1);
                    cb.AddTemplate(p, 0, 0);
                }
            }
        }

        static void AddHeadingStroke()
        {
            Doc.Add(Stroke(525f, 20f, 0, 5f, BaseColor.BLACK, Element.ALIGN_CENTER));
        }

        static void AddSubHeadingStroke()
        {
            Doc.Add(Stroke(525f, 10f, 0, 0.5f, BaseColor.BLACK, Element.ALIGN_CENTER));
        }


        static PdfPTable Stroke(float width, float spacingAbove, float spacingLeft, float thickness, BaseColor color, int alignment)
        {
            PdfPTable table = new PdfPTable(1);
            table.TotalWidth = width;
            table.LockedWidth = true;
            table.SpacingBefore = spacingAbove;
            PdfPCell cell = new PdfPCell();
            cell.BorderWidth = 0;
            cell.BorderWidthBottom = thickness;
            cell.BorderColor = color;
            cell.PaddingLeft = spacingLeft;
            table.AddCell(cell);
            table.HorizontalAlignment = alignment;
            return table;
        }

        /// <summary>
        /// Produces a stroke with a width that will fit underneath a chunk of text, even if the text is multiple lines long
        /// </summary>
        /// <param name="?"></param>
        /// <returns></returns>
        static PdfPTable Underline(Chunk textChunk, int alignment, int numOfLines)
        {
            if (numOfLines > 1)
            {
                string[] words = textChunk.ToString().Split('\n');
                string longestWord = string.Empty;
                Chunk wordChunk = null;
                Chunk longestWordChunk = new Chunk(longestWord);

                foreach (string word in words)
                {
                    wordChunk = new Chunk(word);
                    wordChunk.Font = textChunk.Font;
                    wordChunk.SetCharacterSpacing(textChunk.GetCharacterSpacing());

                    longestWordChunk = new Chunk(longestWord);
                    longestWordChunk.Font = textChunk.Font;
                    longestWordChunk.SetCharacterSpacing(textChunk.GetCharacterSpacing());

                    if (wordChunk.GetWidthPoint() > longestWordChunk.GetWidthPoint())
                    {
                        longestWord = word;
                    }
                }

                longestWordChunk = new Chunk(longestWord);
                longestWordChunk.Font = textChunk.Font;
                longestWordChunk.SetCharacterSpacing(textChunk.GetCharacterSpacing());

                return Stroke(longestWordChunk.GetWidthPoint(), -3f, 0, 0.5f, BaseColor.BLACK, alignment);
            }
            else
            {
                return Stroke(textChunk.GetWidthPoint(), -3f, 0, 0.5f, BaseColor.BLACK, alignment);
            }
        }

        static Font GetNormalFont(float size)
        {
            return new Font(Font.FontFamily.HELVETICA, size, Font.NORMAL, new BaseColor(0, 0, 0));
        }

        static Font GetBoldFont(float size)
        {
            return new Font(Font.FontFamily.HELVETICA, size, Font.BOLD, new BaseColor(0, 0, 0));
        }

        static Font GetItalicFont(float size)
        {
            return new Font(Font.FontFamily.HELVETICA, size, Font.ITALIC, new BaseColor(0, 0, 0));
        }

        static Font GetBoldItalicFont(float size)
        {
            return new Font(Font.FontFamily.HELVETICA, size, Font.BOLDITALIC, new BaseColor(0, 0, 0));
        }

        static string FormatAmount(decimal amount)
        {
            string formattedAmount = amount.ToString("N");

            // Puts negative sign at end
            //if (formattedAmount.StartsWith("-"))
            //{
            //    formattedAmount = formattedAmount.Replace("-", string.Empty);
            //    formattedAmount += "-";
            //}

            return formattedAmount;
        }

        public static int GetNumberOfStatementsBuilt()
        {
            return NumberOfStatementsBuilt;
        }

        static Document Doc
        {
            get;
            set;
        }

        static PdfWriter Writer
        {
            get;
            set;
        }

        private static int NumberOfStatementsBuilt
        {
            get;
            set;
        }

        public static string TEMP_FILE_NAME = "statement_pdf.temp";
    }

    class StatementPageEvent : PdfPageEventHelper
    {
        public override void OnStartPage(PdfWriter writer, Document Document)
        {
            string nextPageTemplate = Configuration.GetStatementTemplateFilePath();

            if (Document.PageNumber > 1)
            {
                Document.SetMargins(STATEMENT_MARGIN_SIDES, STATEMENT_MARGIN_SIDES, STATEMENT_MARGIN_TOP, STATEMENT_MARGIN_BOTTOM);
                Document.NewPage();

                using (FileStream templateInputStream = File.Open(nextPageTemplate, FileMode.Open))
                {
                    // Loads existing PDF
                    PdfReader reader = new PdfReader(templateInputStream);
                    PdfContentByte contentByte = writer.DirectContent;
                    PdfImportedPage page = writer.GetImportedPage(reader, 1);

                    // Copies first page of existing PDF into output PDF
                    //Document.NewPage();
                    contentByte.AddTemplate(page, 0, 0);
                }
            }
            else
            {
                Document.SetMargins(STATEMENT_MARGIN_SIDES, STATEMENT_MARGIN_SIDES, FIRST_PAGE_STATEMENT_MARGIN_TOP, STATEMENT_MARGIN_BOTTOM);
            }
        }


        public static float FIRST_PAGE_STATEMENT_MARGIN_TOP = 12f;
        public static float STATEMENT_MARGIN_TOP = 70f;
        public static float STATEMENT_MARGIN_BOTTOM = 30f;
        public static float STATEMENT_MARGIN_SIDES = 12f;
    }

    }

