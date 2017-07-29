using System;
using System.Collections.Generic;
using System.Text;

namespace MergeExcel
{
    class GetSheet
    {
        #region Var

        /// <summary>
        /// 
        /// </summary>
        public Dictionary<string, string[]> ExceltoSheetDic = new Dictionary<string, string[]>();

        /// <summary>
        /// 
        /// </summary>
        public Dictionary<string, ExcelLoader> SheetToExcelDic = new Dictionary<string, ExcelLoader>();

        #endregion

        #region Fun

        /// <summary>
        /// 
        /// </summary>
        /// <param name="filepath"></param>
        public GetSheet(string filepath)
        {
            LoadFile(filepath);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="filepath"></param>
        public void LoadFile(string filepath)
        {
            string filename = filepath;

            ExcelLoader excel = new ExcelLoader();
            string[] sheets = null;
            excel.SqlStr = " * ";
            excel.setFilePath(filename);
            sheets = GetSheets(excel.getSheets());
            if (sheets == null)
            {
                return;
            }
            foreach (string sheet in sheets)
            {
                if (!SheetToExcelDic.ContainsKey(sheet))
                {
                    SheetToExcelDic.Add(sheet, excel);
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheets"></param>
        /// <returns></returns>
        string[] GetSheets(string[] sheets)
        {
            return sheets;
        }
        #endregion   
                        
    }
}
