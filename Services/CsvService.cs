// Services/CsvService.cs

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using Dossier.Models;
using CsvHelper;
using CsvHelper.Configuration;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Dossier.Services
{
    public class CsvService : ICsvService
    {
        private static readonly List<string> BaseColumnOrder = new List<string>
        {
            "ReceivedDate", "DueDate", "StudentNo", "Programme", "Forename", "Surname",
            "FeeStatus", "QualificationName",
            "DegreeSubject", "InstitutionName", "THERanking", "CountryOfStudy",
            "EquivalencyNote", "OverallGradeGPA", "DegreeStatus", "UKGrade"
        };

        private static readonly List<string> TrailingColumns = new List<string>
        {
            "Decision", "AT", "Note", "Progr. Adm", "Comment"
        };

        private static readonly HashSet<string> RightAlignedColumns = new HashSet<string>
        {
            "THERanking", "DegreeStatus", "UKGrade"
        };

        private static readonly HashSet<string> DateColumns = new HashSet<string>
        {
            "ReceivedDate", "DueDate", "DateOfBirth"
        };

        private List<string> GetColumnOrder(OutputSettings? settings)
        {
            var columns = new List<string>(BaseColumnOrder);

            // Add trailing columns first
            columns.AddRange(TrailingColumns);

            // Optional fields go at the very end
            if (settings != null)
            {
                if (settings.IncludeGender) columns.Add("Gender");
                if (settings.IncludeNationality) columns.Add("Nationality");
                if (settings.IncludeDateOfBirth) columns.Add("DateOfBirth");
                if (settings.IncludeEmail) columns.Add("Email");
                if (settings.IncludePaid) columns.Add("Paid");
            }

            return columns;
        }

        public List<InTrayRecord> LoadInTrayRecords(string filePath)
        {
            try
            {
                var extension = Path.GetExtension(filePath).ToLower();
                
                if (extension == ".xlsx")
                {
                    return LoadInTrayRecordsFromExcel(filePath);
                }
                
                var config = new CsvConfiguration(CultureInfo.InvariantCulture)
                {
                    HeaderValidated = null,
                    MissingFieldFound = null
                };
                
                using var reader = new StringReader(File.ReadAllText(filePath));
                using var csv = new CsvReader(reader, config);
                return csv.GetRecords<InTrayRecord>().ToList();
            }
            catch (IOException ioEx)
            {
                throw new InvalidOperationException($"Cannot read Document 1 (Exported new applicants inTray file).\n\nPlease close the file if it's open in Excel or another program.\n\nFile: {Path.GetFileName(filePath)}", ioEx);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Error loading InTray records: {ex.Message}", ex);
            }
        }

        private List<InTrayRecord> LoadInTrayRecordsFromExcel(string filePath)
        {
            var records = new List<InTrayRecord>();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                if (worksheet.Dimension == null) return records;

                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;

                var headerMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                for (int col = 1; col <= colCount; col++)
                {
                    var header = worksheet.Cells[1, col].Text.Trim();
                    if (!string.IsNullOrEmpty(header))
                    {
                        headerMap[header] = col;
                    }
                }

                string GetValue(int row, string headerName)
                {
                    if (headerMap.TryGetValue(headerName, out int colIndex))
                    {
                        var cell = worksheet.Cells[row, colIndex];
                        return cell.Text?.Trim() ?? "";
                    }
                    return "";
                }

                for (int row = 2; row <= rowCount; row++)
                {
                    var record = new InTrayRecord
                    {
                        ReceivedOn = GetValue(row, "Received on"),
                        StudentNo = GetValue(row, "Student No."),
                        Name = GetValue(row, "Name")
                    };
                    records.Add(record);
                }
            }

            return records;
        }
        
        public List<ApplicationRecord> LoadApplicationRecords(string filePath)
        {
            try
            {
                var extension = Path.GetExtension(filePath).ToLower();
                
                if (extension == ".xlsx")
                {
                    return LoadApplicationRecordsFromExcel(filePath);
                }
                
                var config = new CsvConfiguration(CultureInfo.InvariantCulture)
                {
                    HeaderValidated = null,
                    MissingFieldFound = null
                };
                
                using var reader = new StringReader(File.ReadAllText(filePath));
                using var csv = new CsvReader(reader, config);
                return csv.GetRecords<ApplicationRecord>().ToList();
            }
            catch (IOException ioEx)
            {
                throw new InvalidOperationException($"Cannot read Document 2 (Dept. Application Reports file).\n\nPlease close the file if it's open in Excel or another program.\n\nFile: {Path.GetFileName(filePath)}", ioEx);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Error loading Application records: {ex.Message}", ex);
            }
        }

        private List<ApplicationRecord> LoadApplicationRecordsFromExcel(string filePath)
        {
            var records = new List<ApplicationRecord>();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                if (worksheet.Dimension == null) return records;

                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;

                var headerMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                for (int col = 1; col <= colCount; col++)
                {
                    var header = worksheet.Cells[1, col].Text.Trim();
                    if (!string.IsNullOrEmpty(header))
                    {
                        headerMap[header] = col;
                    }
                }

                string GetValue(int row, string headerName)
                {
                    if (headerMap.TryGetValue(headerName, out int colIndex))
                    {
                        return worksheet.Cells[row, colIndex].Text?.Trim() ?? "";
                    }
                    return "";
                }

                for (int row = 2; row <= rowCount; row++)
                {
                    var record = new ApplicationRecord
                    {
                        ApplicantID = GetValue(row, "Applicant ID"),
                        Programme = GetValue(row, "Programme"),
                        Forename = GetValue(row, "Forename"),
                        Surname = GetValue(row, "Surname"),
                        FeeStatus = GetValue(row, "Fee Status"),
                        QualificationName = GetValue(row, "Qualification name"),
                        DegreeSubject = GetValue(row, "Degree subject"),
                        InstitutionName = GetValue(row, "Institution name"),
                        CountryOfStudy = GetValue(row, "Country of study"),
                        OverallGradeGPA = GetValue(row, "Overall  grade/GPA"),
                        EquivalencyNote = GetValue(row, "Equivalency note"),
                        GradeAchievedPending = GetValue(row, "Grade Achieved/Pending"),
                        // Optional fields
                        Gender = GetValue(row, "Gender"),
                        Nationality = GetValue(row, "Country of Nationality"),
                        DateOfBirth = GetValue(row, "Date of Birth"),
                        Email = GetValue(row, "Email address"),
                        Paid = GetValue(row, "Paid")
                    };
                    records.Add(record);
                }
            }

            return records;
        }
        
        public List<string> GenerateOutputFiles(List<OutputRecord> data, string outputFolderPath, OutputSettings? settings = null)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var columnOrder = GetColumnOrder(settings);
            var programmeGroups = data.GroupBy(record => record.Programme).ToList();
            var outputPaths = new List<string>();

            foreach (var group in programmeGroups)
            {
                var programme = group.Key;
                var records = group.OrderBy(r => ParseDate(r.ReceivedDate)).ToList();

                var outputPath = Path.Combine(
                    outputFolderPath,
                    programme + "_Latest_" + DateTime.Now.ToString("dd_MMM_yyyy_HHmm") + ".xlsx");

                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add(programme);
                    worksheet.View.FreezePanes(2, 1);
                    worksheet.View.ZoomScale = 180;

                    for (int col = 0; col < columnOrder.Count; col++)
                    {
                        var headerCell = worksheet.Cells[1, col + 1];
                        headerCell.Value = columnOrder[col];
                        headerCell.Style.Font.Bold = true;
                        headerCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        headerCell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(240, 240, 240));
                        headerCell.Style.Font.Color.SetColor(System.Drawing.Color.Black);
                        headerCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        headerCell.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    }
                    
                    int row = 2;
                    foreach (var record in records)
                    {
                        for (int col = 0; col < columnOrder.Count; col++)
                        {
                            string columnName = columnOrder[col];
                            string value = columnName switch
                            {
                                "ReceivedDate" => record.ReceivedDate ?? "",
                                "DueDate" => record.DueDate ?? "",
                                "StudentNo" => record.StudentNo ?? "",
                                "Programme" => record.Programme ?? "",
                                "Forename" => record.Forename ?? "",
                                "Surname" => record.Surname ?? "",
                                "FeeStatus" => record.FeeStatus ?? "",
                                "QualificationName" => record.QualificationName ?? "",
                                "DegreeSubject" => record.DegreeSubject ?? "",
                                "InstitutionName" => record.InstitutionName ?? "",
                                "THERanking" => record.THERanking ?? "NR",
                                "CountryOfStudy" => record.CountryOfStudy ?? "",
                                "EquivalencyNote" => record.EquivalencyNote ?? "",
                                "OverallGradeGPA" => record.OverallGradeGPA ?? "",
                                "DegreeStatus" => record.DegreeStatus ?? "",
                                "UKGrade" => record.UKGrade ?? "",
                                // Optional fields
                                "Gender" => record.Gender ?? "",
                                "Nationality" => record.Nationality ?? "",
                                "DateOfBirth" => record.DateOfBirth ?? "",
                                "Email" => record.Email ?? "",
                                "Paid" => record.Paid ?? "",
                                _ => ""
                            };
                            
                            var cell = worksheet.Cells[row, col + 1];
                            
                            if (RightAlignedColumns.Contains(columnName))
                            {
                                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            }
                            else
                            {
                                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                            }
                            
                            if (columnName == "StudentNo")
                            {
                                if (long.TryParse(value, out long studentNoValue))
                                {
                                    cell.Value = studentNoValue;
                                    cell.Style.Numberformat.Format = "0";
                                }
                                else
                                {
                                    cell.Value = value;
                                }
                            }
                            else if (columnName == "THERanking")
                            {
                                cell.Style.Numberformat.Format = "@";
                                cell.Value = value;
                            }
                            else if (DateColumns.Contains(columnName) && !string.IsNullOrWhiteSpace(value))
                            {
                                if (DateTime.TryParse(value, out DateTime dateValue))
                                {
                                    cell.Value = dateValue;
                                    cell.Style.Numberformat.Format = "dd/mm/yyyy";
                                }
                                else
                                {
                                    cell.Value = value;
                                }
                            }
                            else if (columnName == "OverallGradeGPA" && !string.IsNullOrWhiteSpace(value))
                            {
                                cell.Style.Numberformat.Format = "@";
                                cell.Value = value;
                            }
                            else if (columnName == "UKGrade" && !string.IsNullOrWhiteSpace(value))
                            {
                                if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out double gradeValue))
                                {
                                    cell.Value = gradeValue;
                                    cell.Style.Numberformat.Format = "0.0";
                                }
                                else
                                {
                                    cell.Value = value;
                                }
                            }
                            else
                            {
                                cell.Value = value;
                            }
                        }
                        row++;
                    }
                    
                    int theRankingCol = columnOrder.IndexOf("THERanking") + 1;
                    if (row > 2)
                    {
                        var theRankingRange = worksheet.Cells[2, theRankingCol, row - 1, theRankingCol];
                        var ignoredError = worksheet.IgnoredErrors.Add(theRankingRange);
                        ignoredError.NumberStoredAsText = true;
                    }

                    for (int col = 1; col <= columnOrder.Count; col++)
                    {
                        string columnName = columnOrder[col - 1];
                        if (columnName == "EquivalencyNote")
                        {
                            worksheet.Column(col).Width = 18;
                        }
                        else if (columnName == "InstitutionName")
                        {
                            worksheet.Column(col).Width = 48;
                        }
                        else
                        {
                            worksheet.Column(col).AutoFit();
                        }
                    }
                    
                    package.SaveAs(new FileInfo(outputPath));
                }
                
                outputPaths.Add(outputPath);
            }
            
            return outputPaths;
        }
        
        private DateTime ParseDate(string dateString)
        {
            if (string.IsNullOrWhiteSpace(dateString))
                return DateTime.MinValue;
            
            if (DateTime.TryParseExact(dateString, "dd/MM/yy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime result))
                return result;
            
            if (DateTime.TryParse(dateString, out result))
                return result;
            
            return DateTime.MinValue;
        }
    }
}
