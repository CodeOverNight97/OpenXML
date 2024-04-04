using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXml.Model
{
    public class TestDocDuLieuExcelmodel
    {
        public string a { get; set; }
        public double b { get; set; }
        public double c { get; set; }
    }
    public class DocDuLieuExcel<T>
    {
        public DocDuLieuExcel(){
            Data = default(List<T>);
            dataError = new List<DataError>();
        }
        public List<T> Data { get; set; }
        public List<DataError> dataError { get; set; }
    }
    public class DataError
    {
        public int col { get; set; }
        public int row { get; set; }
        public bool isNullOrEmpty { get; set; }
        public bool isWrongType { get; set; }
    }
}
