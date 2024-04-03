using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXml.Service
{
    public static class ServiceDocDuLieuExcel
    {
        public static List<T> XuLy<T>(string fileURL)
        {
            if (!File.Exists(fileURL))
            {
                throw new Exception("Không tìm thấy file");
            }
            List<T> dataList = new List<T>();

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileURL, false))
            {
                WorkbookPart workbookPart = document.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                Worksheet worksheet = worksheetPart.Worksheet;
                SharedStringTable sharedStringTable = workbookPart.SharedStringTablePart.SharedStringTable;

                // Lấy danh sách các cột trong file Excel (từ hàng 1)
                Row headerRow = worksheet.GetFirstChild<SheetData>().Elements<Row>().FirstOrDefault();
                List<string> headers = new List<string>();
                foreach (Cell cell in headerRow.Elements<Cell>())
                {
                    string header = GetCellValue(cell, sharedStringTable);
                    headers.Add(header);
                }

                // Đọc dữ liệu từ các hàng còn lại trong file Excel
                foreach (Row row in worksheet.GetFirstChild<SheetData>().Elements<Row>().Skip(1))
                {
                    T obj = Activator.CreateInstance<T>();

                    for (int i = 0; i < headers.Count; i++)
                    {
                        Cell cell = row.Elements<Cell>().ElementAt(i);
                        string value = GetCellValue(cell, sharedStringTable);

                        // Gán giá trị cho thuộc tính tương ứng của đối tượng
                        var property = typeof(T).GetProperty(headers[i]);
                        if (property != null)
                        {
                            if (property.PropertyType == typeof(string))
                            {
                                property.SetValue(obj, value);
                            }
                            else if (property.PropertyType == typeof(int))
                            {
                                int intValue;
                                if (int.TryParse(value, out intValue))
                                {
                                    property.SetValue(obj, intValue);
                                }
                            }
                            // Xử lý các kiểu dữ liệu thuộc tính khác của đối tượng

                        }
                    }

                    dataList.Add(obj);
                }
            }
            JsonConvert.SerializeObject(dataList);
            return dataList;
        }
        private static string GetCellValue(Cell cell, SharedStringTable sharedStringTable)
        {
            string value = string.Empty;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                int index = int.Parse(cell.InnerText);
                value = sharedStringTable.ElementAt(index).InnerText;
            }
            else if (cell.CellValue != null)
            {
                value = cell.CellValue.InnerText;
            }

            return value;
        }
    }
}
