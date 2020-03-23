using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExecExecl.Model
{
    /// <summary>
    /// Excel 导出类
    /// </summary>
    public class ExportExcelDto
    {
        public int RowNumber { get; set; }

        /// <summary>
        /// 姓名
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 积分日期
        /// </summary>
        public DateTime IntegralDate { get; set; }

        /// <summary>
        /// 积分
        /// </summary>
        public int Integral { get; set; }
    }
}
