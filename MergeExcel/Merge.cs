using System;
using System.Collections.Generic;
// using System.Linq
using System.Text;
using System.Data;
using System.Windows.Forms;
// using Microsoft.Office.Interop.Excel;

namespace MergeExcel
{
    class Merge
    {
        #region private variable define

        /// <summary>
        /// 
        /// </summary>
        private string szLeftExcel;

        /// <summary>
        /// 
        /// </summary>
        private string szRightExcel;       

        /// <summary>
        /// 把表格读入datatable
        /// </summary>
        private DataTable leftDt;

        /// <summary>
        /// 
        /// </summary>
        private DataTable rightDt;

        /// <summary>
        /// 
        /// </summary>
        private DataTable newDt;

        /// <summary>
        /// record station name
        /// </summary>
        List<string> staNameLst;

        List<DataRow> negtiveLst;

        #endregion

        #region public variable define

        

        #endregion

        #region public funcrion define

        /// <summary>
        /// 
        /// </summary>
        /// <param name="left"></param>
        /// <param name="right"></param>
        public Merge(string left, string right)
        {
            szLeftExcel = left;
            szRightExcel = right;            
            leftDt = new DataTable();
            rightDt = new DataTable();
            newDt = new DataTable();
            staNameLst = new List<string>();
            negtiveLst = new List<DataRow>();
        }

        /// <summary>
        /// 
        /// </summary>
        public void MergeEx()
        {
            GetSheet sheetL = new GetSheet(szLeftExcel);
            if (!LoadExcelToDataTable(sheetL, leftDt))
            {
                MessageBox.Show("load left excel file failed");
                return;
            }

            GetSheet sheetR = new GetSheet(szRightExcel);
            if (!LoadExcelToDataTable(sheetR, rightDt))
            {
                MessageBox.Show("load right excel file failed");
                return;
            }

            OldToNewDt(leftDt, newDt);
            OldToNewDt(rightDt, newDt);
            FigureTotal(newDt);
            AddNegtive();
            FigureTotal(negtiveLst);
            foreach (DataRow dr in negtiveLst)
            {
                newDt.Rows.Add(dr);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public DataTable GetDataTable()
        {
            return newDt;
        }              

        #endregion

        #region private function define

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        private bool LoadExcelToDataTable(GetSheet sheet, DataTable dt)
        {
            bool isBegin = false;
            DataTable dtTmp = null;

            if (sheet == null)
            {
                return false;
            }
            dtTmp = sheet.SheetToExcelDic["SHEET1$"].getDataTable("SHEET1$A1:AK300");
            if (dtTmp == null)
            {
                return false;
            }
            foreach (DataRow dr in dtTmp.Rows)
            {
                if (dr["F1"].ToString() == "分类")
                {
                    isBegin = true;
                    foreach (object colName in dr.ItemArray)
                    {
                        if (!string.IsNullOrEmpty(colName.ToString()))
                        {
                            dt.Columns.Add(colName.ToString());
                        }
                    }                  
                }
                else
                {
                    if (isBegin)
                    {
                        DataRow drNew = dt.NewRow();
                        for (int i = 0; i < drNew.ItemArray.Length; i++)
                        {
                            drNew[i] = dr[i].ToString();             
                        }
                        dt.Rows.Add(drNew);
                    }
                }
                if (string.IsNullOrEmpty(dr["F1"].ToString()))
                {
                    isBegin = false;
                }
            }
            return true;
        }
        
        /// <summary>
        /// 
        /// </summary>
        /// <param name="dtOld"></param>
        /// <param name="dtNew"></param>
        private void OldToNewDt(DataTable dtOld, DataTable dtNew)
        {
            string[] szNoCol =
            {
                "分类",
                "物品名称",
                "型号",
                "合计",
                "出库时间",
                "发货时间",
                "备注"
            };
            List<string> noColLst = new List<string>();
            for (int i = 0; i < szNoCol.Length; i++)
            {
                noColLst.Add(szNoCol[i]);
            }

            Dictionary<string, DataRow> restRowDic = new Dictionary<string, DataRow>();

            Dictionary<string, DataRow> allRowDic = new Dictionary<string, DataRow>();

            foreach (DataColumn col in dtOld.Columns)
            {                
                if (!dtNew.Columns.Contains(col.ColumnName))
                {
                    dtNew.Columns.Add(col.ColumnName);
                }
            }
            newDt.Columns["分类"].DataType = typeof(string);

            if (dtNew.Rows.Count < 1)
            {                
                foreach (DataRow dr in dtOld.Rows)
                {
                    DataRow drNew = dtNew.NewRow();
                    foreach (DataColumn col in dtOld.Columns)
                    {
                        drNew[col.ColumnName] = dr[col.ColumnName].ToString();
                    }
                    dtNew.Rows.Add(drNew);
                }

                foreach (DataRow dr in dtOld.Rows)
                {
                    foreach (DataColumn col in dtOld.Columns)
                    {
                        if (!staNameLst.Contains(col.ColumnName) &&
                            !noColLst.Contains(col.ColumnName))
                        {
                            staNameLst.Add(col.ColumnName);
                        }
                    }
                }
            }
            else
            {
                foreach (DataRow drN in dtNew.Rows)
                {
                    foreach (DataRow drO in dtOld.Rows)
                    {
                        if (drN["分类"].ToString() == drO["分类"].ToString() &&
                            drN["物品名称"].ToString() == drO["物品名称"].ToString()) //one row
                        {
                            foreach (DataColumn col in dtOld.Columns)
                            {
                                if (!noColLst.Contains(col.ColumnName) &&
                                    dtNew.Columns.Contains(col.ColumnName))    //one column
                                {
                                    int nTmpO = 0;
                                    int.TryParse(drO[col.ColumnName].ToString(), out nTmpO);
                                    int nTmpN = 0;
                                    int.TryParse(drN[col.ColumnName].ToString(), out nTmpN);
                                    drN[col.ColumnName] = (nTmpN + nTmpO).ToString();                                                                  
                                    if (!staNameLst.Contains(col.ColumnName))
                                    {
                                        staNameLst.Add(col.ColumnName);
                                    }
                                }
                                else if (!noColLst.Contains(col.ColumnName) &&
                                         !dtNew.Columns.Contains(col.ColumnName))  //tow column
                                {
                                    dtNew.Columns.Add(col.ColumnName);
                                    drN[col.ColumnName] = drO[col.ColumnName].ToString();
                                    if (!staNameLst.Contains(col.ColumnName))
                                    {
                                        staNameLst.Add(col.ColumnName);
                                    }
                                }
                            }                            
                            string key = 
                                drO["分类"].ToString() +
                                drO["物品名称"].ToString();
                            if (!allRowDic.ContainsKey(key))
                            {
                                allRowDic.Add(key, drO);
                            }                            
                        }
                        else   //tow row
                        {
                            DataRow drTmp = dtNew.NewRow();
                            drTmp = drO;
                            string key = drTmp["分类"].ToString() + drTmp["物品名称"].ToString();
                            if (!restRowDic.ContainsKey(key))
                            {
                                restRowDic.Add(key, drTmp);
                            }                            
                        }
                    }
                }
                foreach (string szKey in restRowDic.Keys)
                {
                    if (!allRowDic.ContainsKey(szKey))
                    {
                        DataRow drNew = dtNew.NewRow();
                        foreach (DataColumn col in dtOld.Columns)
                        {
                            drNew[col.ColumnName] = restRowDic[szKey][col.ColumnName].ToString();
                        }
                        dtNew.Rows.Add(drNew);
                    }
                }            
            }            
        }

        /// <summary>
        /// calculate the total 
        /// </summary>
        private void FigureTotal(DataTable dt)
        {
            foreach (DataRow dr in dt.Rows)
            {
                int nTotal = 0;
                foreach (string szStaName in staNameLst)
                {
                    int tmp = 0;
                    int.TryParse(dr[szStaName].ToString(), out tmp);
                    if (tmp < 0)
                    {
                        tmp = 0;
                    }
                    nTotal += tmp;
                }
                dr["合计"] = nTotal.ToString();
            }
        }


        /// <summary>
        /// calculate the total 
        /// </summary>
        private void FigureTotal(List<DataRow> list)
        {
            foreach (DataRow dr in list)
            {
                int nTotal = 0;
                foreach (string szStaName in staNameLst)
                {
                    int tmp = 0;
                    int.TryParse(dr[szStaName].ToString(), out tmp);
                    if (tmp < 0)
                    {
                        tmp = tmp * -1;                        
                    }
                    nTotal += tmp;
                }
                dr["合计"] = nTotal.ToString();
            }             
        }

        private void FigureTotal()
        {
            foreach (DataRow dr in newDt.Rows)
            {
                int nTotal = 0;
                foreach (string szStaName in staNameLst)
                {
                    int tmp = 0;
                    int.TryParse(dr[szStaName].ToString(), out tmp);
                    if (tmp < 0)
                    {
                        tmp = 0;
                    }
                    nTotal += tmp;
                }
                dr["合计"] = nTotal.ToString();
            }
        }

        /// <summary>
        /// 
        /// </summary>
        private void AddNegtive()
        {
            List<DataRow> rowLst = new List<DataRow>();
            foreach (DataRow dr in newDt.Rows)
            {
                foreach (string szStaName in staNameLst)
                {                    
                    int nTmp = 0;
                    if (int.TryParse(dr[szStaName].ToString(), out nTmp))
                    {
                        if (nTmp < 0)
                        {
                            rowLst.Add(dr);
                            break;
                        }                     
                    }
                }                
            }            
            
            foreach (DataRow drO in rowLst)
            {
                string keyO =
                    drO["分类"].ToString() +
                    drO["物品名称"].ToString();
                foreach (DataRow drL in rowLst)
                {                    
                    string keyL =
                        drL["分类"].ToString() +
                        drL["物品名称"].ToString();
                    if (keyL == keyO)
                    {
                        DataRow dr = newDt.NewRow();
                        foreach (DataColumn dc in newDt.Columns)
                        {
                            if (staNameLst.Contains(dc.ColumnName))
                            {
                                int nO = 0;
                                int nL = 0;
                                if (int.TryParse(drO[dc.ColumnName].ToString(), out nO))
                                {
                                    if (nO >= 0)
                                    {
                                        nO = 0;
                                    }
                                }
                                if (int.TryParse(drL[dc.ColumnName].ToString(), out nL))
                                {
                                    if (nL >= 0)
                                    {
                                        nL = 0;
                                    }
                                }
                                if (drO != drL)
                                {
                                    nO += nL;
                                }
                                dr[dc.ColumnName] = nO.ToString();                   
                            }
                            else
                            {
                                dr[dc.ColumnName] = drO[dc.ColumnName].ToString();
                            }                            
                        }
                        negtiveLst.Add(dr);                        
                    }                   
                }
            }            
        }
        #endregion
    }
}
