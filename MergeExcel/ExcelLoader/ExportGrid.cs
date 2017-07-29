/*
 * Copyright (C) 2011 河南辉煌科技股份有限公司
 * All rights reserved 
 *
 * 文件摘要：Grid导出Excel
 *
 * 当前版本： 1.0
 * 编写日期： 2011-6-27
 * 作    者： 陈永涛
 *
 * 修改记录： 
 * 
 **/

using System;
using System.Collections.Generic;
using System.Text; 
using Infragistics.Win.UltraWinGrid.ExcelExport;
using Infragistics.Win.UltraWinGrid;
using System.Windows.Forms;
using System.IO;
using System.Data;


namespace MergeExcel
{
    public class ExportGrid
    {
        private UltraGridExcelExporter ultraGridExcelExporter;

        public string fileName = "";

        private UltraGrid ultraGrid = null;

        public ExportGrid(UltraGrid ultragrid,string fileName)
        {
            ultraGridExcelExporter = new UltraGridExcelExporter();
            this.fileName = fileName;
            this.ultraGrid = ultragrid;
        }


        public void ExportToFile()
        {
            bool ishidden = true;
            foreach (UltraGridColumn c in this.ultraGrid.DisplayLayout.Bands[0].Columns)
            {
                if (!c.Hidden)
                {
                    ishidden = false;
                }
            }

            if (ishidden || this.ultraGrid.DisplayLayout.RowScrollRegions[0].VisibleRows.Count == 0 || 
                (this.ultraGrid.DisplayLayout.RowScrollRegions[0].VisibleRows.Count == 1 
                && this.ultraGrid.DisplayLayout.RowScrollRegions[0].VisibleRows[0].Row.Index == -1))
            {
                MessageBox.Show("没有可导出的数据!", "提示");
                return;
            }
            try
            {
                fileName = fileName + "_" + DateTime.Now.ToString("yyyy年MM月dd日HH时mm分ss秒") + ".xls";
                SaveFileDialog savefile = new SaveFileDialog();                
                if (savefile.ShowDialog() == DialogResult.OK)
                {
                    if (!savefile.FileName.EndsWith(".xls"))
                    {
                        savefile.FileName += ".xls";
                    }
                    this.ultraGridExcelExporter.Export(ultraGrid, savefile.FileName);
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("数据导出失败，失败信息如下：" + ex.ToString(), "错误");
            }
        }
    }
}
