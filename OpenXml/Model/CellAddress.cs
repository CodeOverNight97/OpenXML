using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXml.Model
{
    public class CellAddress    
    {
        public Cell? foundCell { get; set; }
        public uint? foundRowIndex { get; set; }
        public string? foundColumn { get; set; } 
        public CellAddress(Cell? _foundCell = null, uint? _foundRowIndex = null, string? _foundColumn = null) {
            this.foundCell = _foundCell;
            this.foundRowIndex = _foundRowIndex;
            this.foundColumn = _foundColumn; 
        }
    } 
}
