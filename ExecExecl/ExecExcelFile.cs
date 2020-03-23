using ExecExecl.Model;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExecExecl
{
    public static class ExecExcelFile
    {
        /// <summary>
        /// 读取并处理Excel,产生新的模板Excel
        /// </summary>
        /// <param name="filePath">文件完整路劲(包含文件名)</param>
        /// <param name="fileName">文件名 为空时从文件路径解析</param>
        /// <returns></returns>
        public static string ReadAndGenerateExcel(string filePath, string fileName = "")
        {
            string newFilePath = "";

            var ltExportExcelDto = ReadExcel(filePath, fileName);
            ///产生新的Excel
            newFilePath = GenerateNewExcel(filePath, ltExportExcelDto);
            return newFilePath;
        }

        /// <summary>
        /// 读取Excel数据
        /// </summary>
        /// <param name="filePath">文件完整路劲(包含文件名)</param>
        /// <param name="fileName">文件名 为空时从文件路径解析</param>
        /// <returns></returns>
        public static List<ExportExcelDto> ReadExcel(string filePath, string fileName = "")
        {
            List<ExportExcelDto> ltOutput = new List<ExportExcelDto>();
            FileInfo sourceFileInfo = new FileInfo(filePath);
            if (sourceFileInfo.Exists)
            {
                if (string.IsNullOrEmpty(fileName))
                {
                    fileName = sourceFileInfo.Name;
                }


                using (ExcelPackage package = new ExcelPackage(sourceFileInfo))
                {
                    ExcelWorksheets worksheets = package.Workbook.Worksheets;

                    foreach (ExcelWorksheet worksheet in worksheets)
                    {
                        int rowIndex = 3;

                        int rowCount = worksheet.Cells.Rows;
                        //获得有数据的区域
                        var lastAddress = worksheet.Dimension.Address;
                        //获得有数据的区域最上且最左的单元格
                        //var startCell = worksheet.Dimension.Start.Address;
                        //获得有数据的区域最下且最右的单元格
                        //var endCell = worksheet.Dimension.End.Address;

                        int minRowNum = worksheet.Dimension.Start.Row; //工作区开始行号
                        int maxRowNum = worksheet.Dimension.End.Row; //工作区结束行号
                        int lasColumn = worksheet.Dimension.Columns;

                        try
                        {
                            //开始遍历Excel (从固定行开始读取)
                            for (; rowIndex <= maxRowNum; rowIndex++)
                            {
                                ExportExcelDto exportExcelDto = new ExportExcelDto()
                                {
                                    RowNumber = rowIndex
                                };

                                exportExcelDto.Name = ReadCellStringValue(worksheet, "A", rowIndex);
                                exportExcelDto.Integral = ReadCellIntValue(worksheet, lasColumn, rowIndex);

                                ltOutput.Add(exportExcelDto);
                            }
                        }
                        catch (Exception ex)
                        {
                        }
                    }
                }
            }

            return ltOutput;
        }

        /// <summary>
        /// 产生新的Excel 
        /// </summary>
        private static string GenerateNewExcel(string filePath, List<ExportExcelDto> ltExportExcelDto)
        {
            //先统计再导出
            var ltData = from t in ltExportExcelDto
                         group t by t.Name into g
                         select new ExportExcelDto
                         {
                             Name = g.Key,
                             Integral = g.Sum(m => m.Integral)
                         };
            ltData = ltData.OrderBy(aa => aa.Name).ToList();
            string fileName = "NewFile.xlsx";
            using (ExcelPackage package = new ExcelPackage())
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");
                //表头
                worksheet.Cells[$"A1"].Value = "姓名";
                worksheet.Cells[$"B1"].Value = "总计";
                //遍历错误集合处理数据
                int index = 2;
                foreach (var exportDto in ltData)
                {
                    //表头
                    worksheet.Cells[$"A{index}"].Value = exportDto.Name;
                    worksheet.Cells[$"B{index}"].Value = exportDto.Integral;

                    index++;
                }
                if (File.Exists("NewFile.xlsx"))
                {
                    File.Delete("NewFile.xlsx");
                }

                using (var fs = File.OpenWrite("NewFile.xlsx"))
                {
                    package.SaveAs(fs);
                }
            }
            return fileName;
        }

        #region 读取Excel用
        /// <summary>
        /// 读取excel单元格的字符串值（列的标识为空 则返回""）
        /// </summary>
        /// <param name="worksheet">Excel的当前Sheet</param>
        /// <param name="col">列的标识</param>
        /// <param name="rowIndex">行所以</param>
        /// <returns>如果列的标识为空 则返回""</returns>
        private static string ReadCellStringValue(ExcelWorksheet worksheet, string col, int rowIndex)
        {
            if (string.IsNullOrEmpty(col))
            {
                return "";
            }
            else
            {
                return worksheet.Cells[col + rowIndex].GetValue<string>();
            }
        }

        /// <summary>
        /// 读取excel单元格的数值（列的标识为空 则返回0）
        /// </summary>
        /// <param name="worksheet">Excel的当前Sheet</param>
        /// <param name="col">列的标识</param>
        /// <param name="rowIndex">行所以</param>
        /// <returns>如果列的标识为空 则返回0</returns>
        private static int ReadCellIntValue(ExcelWorksheet worksheet, string col, int rowIndex)
        {
            if (string.IsNullOrEmpty(col))
            {
                return 0;
            }
            else
            {
                return worksheet.Cells[col + rowIndex].GetValue<int>();
            }
        }

        /// <summary>
        /// 读取excel单元格的数值（列的标识为空 则返回0）
        /// </summary>
        /// <param name="worksheet">Excel的当前Sheet</param>
        /// <param name="col">列索引</param>
        /// <param name="rowIndex">行索引</param>
        /// <returns>如果列的标识为空 则返回0</returns>
        private static int ReadCellIntValue(ExcelWorksheet worksheet, int col, int rowIndex)
        {
            if (col == 0)
            {
                return 0;
            }
            else
            {
                return worksheet.Cells[rowIndex, col].GetValue<int>();
            }
        }
        /// <summary>
        /// 读取excel单元格的批注文本（列的标识为空 则返回""）
        /// </summary>
        /// <param name="worksheet">Excel的当前Sheet</param>
        /// <param name="col">列的标识</param>
        /// <param name="rowIndex">行所以</param>
        /// <returns>如果列的标识为空 则返回0</returns>
        private static string ReadCellCommentValue(ExcelWorksheet worksheet, string col, int rowIndex)
        {
            if (string.IsNullOrEmpty(col))
            {
                return "";
            }
            else

            {
                //拿到 批注
                return worksheet.Cells[col + rowIndex].Comment?.Text.Replace(worksheet.Cells[col + rowIndex].Comment?.Author, "").Replace("\n", "").Replace("\r", "");
            }
        }
        #endregion
    }
}
