using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CfsImportManager.TablesInfo
{
    public class MainTableInfo : TableInfoBase
    {
        public List<IXLCell> ReferenceTablesCells { get; set; }
        public List<ReferenceTableInfo> ReferenceTables { get; set; }
    }
}
