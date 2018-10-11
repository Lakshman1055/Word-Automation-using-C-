using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDM
{
    public class TableMetaData
    {
        public string TableName { get; set; } = String.Empty;

        public string ColumnName { get; set; } = String.Empty;

        public string Id { get; set; } = String.Empty;

        public string PK { get; set; } = String.Empty;

        public string DataType { get; set; } = String.Empty;

        public string Length { get; set; } = String.Empty;

        public string Precision { get; set; } = String.Empty;

        public string Scale { get; set; } = String.Empty;

        public string Identity { get; set; } = String.Empty;

        public string Comments { get; set; } = String.Empty;

    }
}
