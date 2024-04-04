using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXml.Model;

namespace OpenXml.Service
{
    public static class ServiceDocDuLieuExcel
    {
        public static DocDuLieuExcel<T> XuLy<T>(string fileURL)
        {
            if (!File.Exists(fileURL))
            {
                throw new Exception("Không tìm thấy file");
            }
            List<T> dataList = new List<T>();
            List<DataError> errors = new List<DataError>();
            var res = new DocDuLieuExcel<T>();

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileURL, true))
            {
                WorkbookPart workbookPart = document.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                Worksheet worksheet = worksheetPart.Worksheet;
                SharedStringTable sharedStringTable = workbookPart.SharedStringTablePart.SharedStringTable;

                // Lấy danh sách các cột trong file Excel (từ hàng 1)
                Row headerRow = worksheet.GetFirstChild<SheetData>().Elements<Row>().FirstOrDefault();
                List<string> headers = new List<string>();
                foreach (Cell cell in headerRow.Elements<Cell>().Skip(1))
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
                        Cell cell = row.Elements<Cell>().ElementAtOrDefault(i + 1); //bỏ qua cột đầu
                        string value = GetCellValue(cell, sharedStringTable);
                        var colNum = i + 2;
                        var rowNum = int.Parse(row.RowIndex);

                        // Gán giá trị cho thuộc tính tương ứng của đối tượng
                        var property = typeof(T).GetProperty(headers[i]);
                        if (property != null)
                        {
                            if (property.PropertyType == typeof(string))
                            {
                                property.SetValue(obj, value);
                            }
                            else if (property.PropertyType == typeof(DateTime))
                            {
                                DateTime intValue;
                                if (DateTime.TryParse(value, out intValue))
                                {
                                    property.SetValue(obj, intValue);
                                    if (string.IsNullOrEmpty(value))
                                    {
                                        errors.Add(new DataError()
                                        {
                                            col = colNum,
                                            row = rowNum,
                                            isNullOrEmpty = true,
                                            isWrongType = false
                                        });
                                    }
                                }
                                else if (!DateTime.TryParse(value, out intValue) && string.IsNullOrEmpty(value))
                                {
                                    errors.Add(new DataError()
                                    {
                                        col = colNum,
                                        row = rowNum,
                                        isNullOrEmpty = true,
                                        isWrongType = false
                                    });
                                }
                                else if (!DateTime.TryParse(value, out intValue) && !string.IsNullOrEmpty(value))
                                {
                                    errors.Add(new DataError()
                                    {
                                        col = colNum,
                                        row = rowNum,
                                        isNullOrEmpty = false,
                                        isWrongType = true
                                    });
                                }
                            }
                            else if (property.PropertyType == typeof(double))
                            {
                                double intValue;
                                if (double.TryParse(value, out intValue))
                                {
                                    property.SetValue(obj, intValue);
                                }
                                else if (!double.TryParse(value, out intValue) && string.IsNullOrEmpty(value))
                                {
                                    errors.Add(new DataError()
                                    {
                                        col = colNum,
                                        row = rowNum,
                                        isNullOrEmpty = true,
                                        isWrongType = false
                                    });
                                }
                                else if (!double.TryParse(value, out intValue) && !string.IsNullOrEmpty(value))
                                {
                                    errors.Add(new DataError()
                                    {
                                        col = colNum,
                                        row = rowNum,
                                        isNullOrEmpty = false,
                                        isWrongType = true
                                    });
                                }
                            }
                            else if (property.PropertyType == typeof(bool))
                            {
                                Decimal intValue;
                                if (Decimal.TryParse(value, out intValue))
                                {
                                    if (intValue == 1)
                                        property.SetValue(obj, true);
                                    else if (intValue == 0)
                                        property.SetValue(obj, false);
                                    else
                                    {
                                        errors.Add(new DataError()
                                        {
                                            col = rowNum,
                                            row = rowNum,
                                            isNullOrEmpty = false,
                                            isWrongType = true
                                        });
                                    }
                                }
                                else if (!Decimal.TryParse(value, out intValue) && string.IsNullOrEmpty(value))
                                {
                                    errors.Add(new DataError()
                                    {
                                        col = rowNum,
                                        row = rowNum,
                                        isNullOrEmpty = true,
                                        isWrongType = false
                                    });
                                }
                                else if (!Decimal.TryParse(value, out intValue) && !string.IsNullOrEmpty(value))
                                {
                                    errors.Add(new DataError()
                                    {
                                        col = rowNum,
                                        row = rowNum,
                                        isNullOrEmpty = false,
                                        isWrongType = true
                                    });
                                }
                            }
                            else
                            {
                                throw new Exception("Kiểu dữ liệu của model chỉ có thể là Double, String , Boolean hoặc Datetime");
                            }
                        }
                    }

                    dataList.Add(obj);
                }
            }
            res.Data = dataList;
            res.dataError = errors;
            return res;
        }
        private static string GetCellValue(Cell cell, SharedStringTable sharedStringTable)
        {
            string value = string.Empty;
            if (cell == null)
            {
                value = null;
            }
            else if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
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
