using ClosedXML.Excel;
using log4net;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleExcel
{
    internal class XLParser
    {
        private readonly ILog _logger;
        // Constructor that accepts a logger
        public XLParser(ILog logger)
        {
            _logger = logger ?? LogManager.GetLogger(typeof(XLParser));
            _logger.Debug("XLParser instance created");
        }

        // Optional default constructor if needed for backward compatibility
        public XLParser() : this(LogManager.GetLogger(typeof(XLParser)))
        {
        }
        public void OpenFilewithOption(string filePath, string option)
        {
            _logger.Debug($"open the file {filePath}");

            // Open the file using ClosedXML
            try
            {

                var workbook = new XLWorkbook(filePath);
                if (workbook.TryGetWorksheet(option, out IXLWorksheet salesSheet))
                {
                    // Read the first cell value from the specified worksheet
                    // Read A1:U17
                    var range = salesSheet.Range("A1:U17");
                    var data = range.Cells().Select(cell => cell.GetString()).ToList();
                    // print the data to the console
                    _logger.Debug($"Data from worksheet '{option}':");
                    foreach (var row in data)
                    {
                        _logger.Debug(row);
                    }

                    // set A1:U17 some values
                    _logger.Debug($"Setting values in worksheet '{option}' from A1:U17");
                    range = salesSheet.Range("A1:U17");
                    for (int i = 0; i < range.RowCount(); i++)
                    {
                        for (int j = 0; j < range.ColumnCount(); j++)
                        {
                            range.Cell(i + 1, j + 1).Value = i * range.ColumnCount() + j + 1;
                        }
                    }
                    // Save the changes to the workbook , use a new file name to avoid overwriting
                    string newFileName = System.IO.Path.GetFileNameWithoutExtension(filePath) + "_modified.xlsx";
                    // if the file already exists, delete it
                    if (System.IO.File.Exists(newFileName))
                    {
                        System.IO.File.Delete(newFileName);
                    }
                    _logger.Debug($"Saving modified workbook as '{newFileName}'");
                    workbook.SaveAs(newFileName);
                }
                else
                {
                    _logger.Warn($"Worksheet '{option}' not found in the workbook.");
                }

            }
            catch (Exception ex)
            {
                _logger.Error($"Failed to open file: {filePath}", ex);
                throw;
            }


        }
    }
}
