/*
//-----------------------------------------------------------------------
// <版权 开源 文件名称="Readable.cs">
//  版本 (c)  V1.0.0.0  
//  创建者:   少林
//  创建时间:   2021-08-26 10:31:54
//  功能描述:   读取excel文件内容
//  历史版本:
//          2021-08-26 少林 读取excel文件内容
// </copyright>
//-----------------------------------------------------------------------
//-----------------------------------------------------------------------
// <copyright open code source file="Readable.cs">
//  Copyright (c)  V1.0.0.0  
//  creator:   arison
//  create time:   2021-08-26 10:31:54
//  function description:   read it from the excel file
//  history version:
//          2021-08-26 arison read it from the excel file
// </copyright>
//-----------------------------------------------------------------------
*/
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel.Utilies
{
    /// <summary>
    /// 读取excel文件方法类
    ///  read utily class(read excel file)
    /// </summary>
    public class Readable
    {
        /// <summary>
        /// 从excel文件里面获取DataSet数据集(多个sheet就存在多个内容表)
        /// get a dataset from the excel file(exist sheets then tables)
        /// </summary>
        /// <param name="fileName">
        /// excel文件路径
        /// excel file path
        /// </param>
        /// <param name="isFirstRowColumn">第一行是否是DataTable的列名</param>
        /// <returns></returns>
        public static DataSet GetDataSet(string fileName, bool isFirstRowColumn = true)
        {
            DataSet set = null;
            IWorkbook workbook = GetWorkbook(fileName);
            int len = workbook.NumberOfSheets;
            set = new DataSet();
            for (int i = 0; i < len; i++)
            {
                DataTable table = null;
                SheetToDataTable(workbook.GetSheetAt(i), isFirstRowColumn, ref table);
                if (table != null)
                {
                    set.Tables.Add(table);
                }
            }
            return set;
        }

        /// <summary>
        /// 将excel中的数据导入到DataTable中
        /// </summary>
        /// <param name="sheetName">excel工作薄sheet的名称</param>
        /// <param name="isFirstRowColumn">第一行是否是DataTable的列名</param>
        /// <returns>返回的DataTable</returns>
        public static DataTable GetDataTableBySheetIndex(string fileName, int index, bool isFirstRowColumn = true)
        {
            DataTable data = new DataTable();
            IWorkbook workbook = GetWorkbook(fileName);
            var sheet = workbook.GetSheetAt(index);
            if (sheet != null)
            {
                SheetToDataTable(sheet, isFirstRowColumn, ref data);
            }
            else
            {
                throw new Exception("sheet名称不存在");
            }
            return data;
        }


        /// <summary>
        /// 将excel中的数据导入到DataTable中
        /// </summary>
        /// <param name="sheetName">excel工作薄sheet的名称</param>
        /// <param name="isFirstRowColumn">第一行是否是DataTable的列名</param>
        /// <returns>返回的DataTable</returns>
        public static DataTable GetDataTableBySheetName(string fileName, string sheetName, bool isFirstRowColumn = true)
        {
            DataTable data = new DataTable();
            IWorkbook workbook = GetWorkbook(fileName);
            var sheet = workbook.GetSheet(sheetName);
            if (sheet != null)
            {
                SheetToDataTable(sheet, isFirstRowColumn, ref data);
            }
            else
            {
                throw new Exception("sheet名称不存在");
            }
            return data;
        }


        /// <summary>
        /// 获取excel文件的workbook对象
        /// </summary>
        /// <param name="fileName">excel文件路径</param>
        /// <returns>返回的IWorkbook对象</returns>
        public static IWorkbook GetWorkbook(string fileName)
        {
            IWorkbook workbook = null;  //新建IWorkbook对象
            var fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read);
            string suffix = fileName.Substring(fileName.LastIndexOf(".") + 1, fileName.Length - fileName.LastIndexOf(".") - 1);
            suffix = suffix.Trim().ToLower();
            if (suffix == "xls")
            {
                workbook = new HSSFWorkbook(fileStream);
            }
            else if (suffix == "xlsx")
            {
                workbook = new XSSFWorkbook(fileStream);
            }
            else
            {
                throw new Exception("传入文件的后缀必须为xlsx或者xls");
            }
            return workbook;
        }

        /// <summary>
        /// 将Excel中的工作薄转换为DataTable
        /// convert sheet into data table
        /// </summary>
        /// <param name="sheet">
        /// Excel中的工作薄
        /// sheet
        /// </param>
        /// <param name="isFirstRowColumn">
        /// 第一行是否是DataTable的列名
        /// whether or not the first row is table's column name
        /// </param>
        /// <param name="data">
        /// 表
        /// table from the sheet object
        /// </param>
        private static void SheetToDataTable(ISheet sheet, bool isFirstRowColumn, ref DataTable data)
        {
            //最后一列的标号
            //get last row's position
            int rowCount = sheet.LastRowNum;
            IRow firstRow = sheet.GetRow(sheet.FirstRowNum);
            //一行最后一个cell的编号 即总的列数
            //get last cell's position
            int cellCount = firstRow.LastCellNum;
            int startRow = 0;
            if (isFirstRowColumn)
            {
                for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                {
                    if (firstRow.GetCell(i) == null) continue;
                    DataColumn column = new DataColumn(firstRow.GetCell(i).ToString());
                    data.Columns.Add(column);
                }
                startRow = sheet.FirstRowNum + 1;
            }
            else
            {
                startRow = sheet.FirstRowNum;
            }
            for (int i = startRow; i <= rowCount; ++i)
            {
                IRow row = sheet.GetRow(i);
                //没有数据的行为null　　　
                if (row == null) continue;
                DataRow dataRow = data.NewRow();
                for (int j = row.FirstCellNum; j < cellCount; ++j)
                {
                    //没有数据的单元格也为null
                    if (row.GetCell(j) != null)
                    {
                        //一般情况下row.FirstCellNum为0，但有时excel中的数据并不在A列，
                        //所以需减去，否则将导致溢出，出现异常。
                        dataRow[j - row.FirstCellNum] = row.GetCell(j).ToString();
                    }
                }
                data.Rows.Add(dataRow);
            }
        }
    }
}
