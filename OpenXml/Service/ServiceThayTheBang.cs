using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using OpenXml.Model;

namespace OpenXml.Service
{
    public static class ServiceThayTheBang
    {
        public static void XuLy()
        {
            #region FAKE VALUE
            DataSet dataSet = getData();
            string _templatePath = "D:\\OPEN_XML_SOLUTION\\OpenXML\\OpenXml\\Files\\Template\\BIEUMAUSO06.xlsx";
            string _outputPath = "D:\\OPEN_XML_SOLUTION\\OpenXML\\OpenXml\\Files\\Upload\\BIEUMAUSO06_" + RandomNumber(5) + ".xlsx";
            #endregion

            if (!File.Exists(_outputPath))
                File.Copy(_templatePath, _outputPath, true); // Tạo file output từ templatePath

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(_outputPath, true))
            {
                WorkbookPart? workbookPart = document.WorkbookPart;
                Workbook? workbook = workbookPart!.Workbook;

                SetDataSource(dataSet, workbookPart);

                // Lưu thông tin thay đổi
                //workbook.Save();
                document.Save();
            }
        }
        private static void SetDataSource(DataSet dataSet, WorkbookPart workbookPart)
        {

            foreach (WorksheetPart worksheetPart in workbookPart.WorksheetParts)
            {
                Worksheet worksheet = worksheetPart.Worksheet;
                SheetData? sheetData = worksheet.GetFirstChild<SheetData>();
                if (sheetData == null)
                    return;

                InsertRow(worksheet, "F10");
                foreach (DataTable dataTable in dataSet.Tables)
                {
                    string tableName = dataTable.TableName;
                    foreach (DataColumn column in dataTable.Columns)
                    {
                        string searchText = $"{tableName}.{column.ColumnName}";
                        // Lặp qua mỗi hàng để tìm list<cell> có giá trị
                        List<CellAddress> cellAddress = GetCellAddressesByText(sheetData, workbookPart, searchText);

                        if (cellAddress.Count == 0)
                            continue;

                        foreach (CellAddress celladr in cellAddress)
                        {
                            int rowIndex = 0;
                            foreach (DataRow row in dataTable.Rows)
                            {
                                rowIndex++;
                                string cellReference = celladr.foundColumn + (celladr.foundRowIndex + rowIndex);
                                //InsertRow(worksheet, cellReference);
                                //Cell? cellBelow = GetCell(worksheet, cellReference);
                                //cellBelow!.CellValue = new CellValue(row[column].ToString() ?? "");
                                //cellBelow.DataType = cellBelow.DataType;
                            }
                        }
                    }
                }
            }
        }
        /// <summary>
        /// Hàm GetCellAddressesByText: Tìm kiếm tất cả các địa chỉ của cell chứa giá trị tương đương với searchText trong sheet.
        /// </summary>
        /// <param name="sheetData">Dữ liệu của sheet cần tìm kiếm.</param>
        /// <param name="workbookPart">Phần WorkbookPart của workbook.</param>
        /// <param name="searchText">Giá trị cần tìm kiếm trong các ô.</param>
        /// <returns>Danh sách các địa chỉ của cell chứa giá trị tương đương với searchText.</returns>
        private static List<CellAddress> GetCellAddressesByText(SheetData sheetData, WorkbookPart? workbookPart, string? searchText)
        {
            if (sheetData == null || workbookPart == null || searchText == null)
                return null;

            List<CellAddress> result = new List<CellAddress>();

            foreach (Row? row in sheetData.Elements<Row>())
            {
                foreach (Cell cell in row.Elements<Cell>())
                {
                    string? cellValue = GetCellValue(workbookPart, cell); // Lấy giá trị của ô từ workbook
                    if (cellValue == searchText)
                    {
                        result.Add(new CellAddress(cell, row.RowIndex, GetColumnName(cell.CellReference)));  // Nếu tìm thấy giá trị cần tìm, trả về địa chỉ của ô
                    }
                }
            }

            return result; // Trả về kết quả (có thể là null nếu không tìm thấy)
        }
        /// <summary>
        /// Hàm GetCellValue: Lấy giá trị của một ô.
        /// </summary>
        /// <param name="workbookPart">Phần WorkbookPart của workbook.</param>
        /// <param name="cell">Ô cần lấy giá trị.</param>
        /// <returns>Giá trị của ô.</returns>
        public static string? GetCellValue(WorkbookPart workbookPart, Cell cell)
        {
            SharedStringTablePart? stringTablePart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                if (stringTablePart != null)
                {
                    return stringTablePart.SharedStringTable.ElementAt(int.Parse(cell.InnerText)).InnerText;
                }
            }
            return cell.InnerText;
        }
        /// <summary>
        /// Hàm GetColumnName: Chuyển đổi tham chiếu cột thành tên cột (ví dụ: "A1" thành "A").
        /// </summary>
        /// <param name="cellReference">Tham chiếu cột cần chuyển đổi.</param>
        /// <returns>Tên của cột.</returns>
        private static string? GetColumnName(string? cellReference)
        {
            string columnName = "";
            foreach (char c in cellReference)
            {
                if (Char.IsLetter(c))
                    columnName += c;
                else
                    break;
            }
            return columnName;
        }
        /// <summary>
        /// Hàm GetCell: Lấy ô từ một bảng tính dựa trên tham chiếu ô (ví dụ: "A1").
        /// </summary>
        /// <param name="worksheet">Bảng tính chứa ô cần lấy.</param>
        /// <param name="cellReference">Tham chiếu của ô cần lấy.</param>
        /// <returns>Ô được lấy từ bảng tính.</returns>
        private static Cell? GetCell(Worksheet? worksheet, string? cellReference)
        {
            if (worksheet == null || cellReference == null)
                return null;
            Cell? cell = null;
            SheetData? sheetData = worksheet.GetFirstChild<SheetData>();

            // Parse the cell reference
            string columnName = Regex.Replace(cellReference, @"[\d-]", string.Empty);
            int rowIndex = int.Parse(Regex.Replace(cellReference, @"[^\d]", string.Empty));

            // Find the row
            Row? row = sheetData!.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
            if (row != null)
            {
                // Find the cell in the row
                cell = row.Elements<Cell>().FirstOrDefault(c => string.Compare(c.CellReference!.Value, cellReference, true) == 0);

                // If cell is null, create a new one
                if (cell == null)
                {
                    // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                    Cell? refCell = null;
                    foreach (Cell currentCell in row.Elements<Cell>())
                    {
                        if (string.Compare(currentCell.CellReference?.Value, cellReference, true) > 0)
                        {
                            refCell = currentCell;
                            break;
                        }
                    }
                    cell = new Cell() { CellReference = new StringValue(cellReference) };
                    row.InsertBefore(cell, refCell);
                }
            }
            else
            {
                // If the row doesn't exist, create a new row and add the cell to it
                row = new Row() { RowIndex = (uint)rowIndex };
                cell = new Cell() { CellReference = new StringValue(cellReference) };
                row.Append(cell);
                sheetData.Append(row);
            }

            return cell;
        }
        private static void InsertRow(Worksheet worksheet, string cellReference)
        {
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();

            //Row? currentRow = GetRow(sheetData, cellReference);
            Row newRow = new Row();

            sheetData.InsertAt(newRow, 11);
        }

        private static Row? GetRow(SheetData sheetData, string cellReference)
        {
            int rowIndex = int.Parse(Regex.Match(cellReference, @"\d+").Value);

            foreach (Row row in sheetData.Elements<Row>())
            {
                if (row.RowIndex == rowIndex)
                {
                    return row;
                }
            }

            return null;
        }


        /// <summary>
        /// Hàm RandomNumber: Random chuỗi số bất kỳ
        /// </summary>
        /// <returns></returns>
        private static string RandomNumber(int numberRD)
        {
            string randomStr = "";
            try
            {
                int[] myIntArray = new int[numberRD];
                int x;
                //that is to create the random # and add it to uour string
                Random autoRand = new Random();
                for (x = 0; x < numberRD; x++)
                {
                    myIntArray[x] = Convert.ToInt32(autoRand.Next(0, 9));
                    randomStr += (myIntArray[x].ToString());
                }
            }
            catch
            {
                randomStr = "999";
            }
            return randomStr;
        }
        /// <summary>
        /// Load data test
        /// </summary>
        /// <returns></returns>
        private static DataSet getData()
        {
            DataSet dataSet = new DataSet();
            try
            {
                #region table 1
                DataTable table1 = new DataTable();
                table1.TableName = "DATA";

                DataColumn dc = new DataColumn("STT", typeof(String));
                table1.Columns.Add(dc);
                dc = new DataColumn("Name", typeof(String));
                table1.Columns.Add(dc);

                DataRow newRow = table1.NewRow();
                newRow["STT"] = "0";
                newRow["Name"] = "41234";
                table1.Rows.Add(newRow);

                newRow = table1.NewRow();
                newRow["STT"] = "1";
                newRow["Name"] = "1234123";
                table1.Rows.Add(newRow);

                newRow = table1.NewRow();
                newRow["STT"] = "2";
                newRow["Name"] = "12341234";
                table1.Rows.Add(newRow);
                #endregion

                #region table 2
                DataTable table2 = new DataTable();
                table2.TableName = "DATA1";

                DataColumn dc2 = new DataColumn("ID", typeof(String));
                table2.Columns.Add(dc2);
                dc2 = new DataColumn("Value", typeof(String));
                table2.Columns.Add(dc2);

                DataRow newRow2 = table2.NewRow();
                newRow2["ID"] = "1000";
                newRow2["Value"] = "412342";
                table2.Rows.Add(newRow2);

                newRow2 = table2.NewRow();
                newRow2["ID"] = "1001";
                newRow2["Value"] = "12342412";
                table2.Rows.Add(newRow2);

                newRow2 = table2.NewRow();
                newRow2["ID"] = "1002";
                newRow2["Value"] = "23412341";
                table2.Rows.Add(newRow2);
                #endregion

                dataSet.Tables.Add(table1.Copy());
                dataSet.Tables.Add(table2.Copy());
                return dataSet;
            }
            catch (Exception ex)
            {
                Console.WriteLine("ServiceThayTheBang - getData()" + ex.ToString());
                return dataSet;
            }
        }
    }
}
