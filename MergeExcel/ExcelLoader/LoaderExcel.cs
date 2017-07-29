using System;
using System.Collections.Generic;
using System.Text;
using System.Collections;
using System.Data.OleDb;
using System.Windows.Forms;
using System.Data;
using System.IO;

namespace MergeExcel
{
    public class ExcelLoader
    {
        #region 私有字段

        //连接字符串
        private string _connstr = string.Empty;
        //准备导入的excel文件路径
        private string _filePath = string.Empty;
        //excel中所有的sheet的名称集合
        private string[] _sheets;
        //sql转换语句
        private string _sqlStr = string.Empty;
        //过滤条件
        private string _filter = string.Empty;
        //是否获得了数据路径
        private bool _isGetPath = false;

        #endregion

        #region 属性

        /// <summary>
        /// 设置sql转换字符串
        /// </summary>
        public string SqlStr
        {
            get
            {
                return _sqlStr;
            }
            set
            {
                _sqlStr = value;
            }
        }
        /// <summary>
        /// 设置过滤条件
        /// </summary>
        public string Filter
        {
            get
            {
                return _filter;
            }
            set
            {
                _filter = value;
            }
        }
        /// <summary>
        /// 获得文件导入的路径
        /// </summary>
        public string FileName
        {
            get
            {
                return _filePath;
            }
        }

        #endregion

        #region 私有方法
        /// <summary>
        /// 获得连接字符串
        /// </summary>
        /// <param name="filePath"></param>
        private void getConnStr(string filePath)
        {
            if (filePath.Equals(string.Empty))
            {
                throw new Exception("请选择导入的Excel文件路径!");
            }
            _connstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +
                       filePath +
                       ";Extended Properties='Excel 8.0;HDR=NO;IMEX=1';";
        }

        /// <summary>
        /// 获得查询excel数据的sql语句
        /// </summary>
        /// <param name="tables"></param>
        /// <returns></returns>
        private string getSqlStr(string[] sheets, string sheetName)
        {
            if (sheetName == null || string.IsNullOrEmpty(sheetName))
            {
                return "";
            }
            string sql = string.Empty;
            string subsql = "select " + _sqlStr + " from ";
            string filter = "";
            if (!_filter.Equals(string.Empty))
            {
                filter = " where " + _filter;
            }
            sql += subsql + "[" + sheetName + "] " + filter;
            return sql.TrimEnd("union ".ToCharArray());
        }

        /// <summary>
        /// 把Excel中所有的sheet名，加载到checkListBox中
        /// </summary>
        /// <param name="sheets"></param>
        /// <param name="clb"></param>
        public void LoadCheckListBox(System.Windows.Forms.CheckedListBox clb)
        {
            clb.Items.Clear();
            getSheets();
            if (_sheets == null || _sheets.Length <= 0)
            {
                return;
            }
            clb.MultiColumn = true;
            clb.CheckOnClick = true;
            clb.Items.AddRange(_sheets);
        }

        /// <summary>
        /// 获得Excel所有sheet的名称
        /// </summary>
        /// <returns></returns>
        public string[] getSheets()
        {
            string[] names;
            DataTable dt;
            OleDbConnection conn = null;
            try
            {
                conn = new OleDbConnection(_connstr);
                conn.Open();
                dt = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            }
            catch (OleDbException e)
            {
                MessageBox.Show("Excel表导入数据出错！" + e.Message);
                return null;
            }
            finally
            {
                if (conn != null)
                {
                    conn.Close();
                }
            }
            names = new string[dt.Rows.Count];
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string name = dt.Rows[i][2].ToString();
                names[i] = name;
            }
            _sheets = names;
            return names;
        }
        #endregion

        #region 公共方法

        /// <summary>
        ///  返回所有符合条件的数据表
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public DataTable getDataTable(string sheetName)
        {
            DataTable dt = null;
            if (_isGetPath)
            {
                if (getSheets() != null)
                {
                    string sql = getSqlStr(_sheets, sheetName);
                    if (string.IsNullOrEmpty(sql))
                    {
                        return null;
                    }
                    try
                    {
                        OleDbDataAdapter adapter = new OleDbDataAdapter(sql, _connstr);
                        DataSet ds = new DataSet();
                        adapter.Fill(ds);
                        dt = ds.Tables[0];
                    }
                    catch (OleDbException e)
                    {
                        MessageBox.Show("Excel表导入数据出错！" + e.Message);
                    }
                }
            }
            return dt;
        }


        /// <summary>
        /// 获得Excel文件路径
        /// </summary>
        public string getFilePath()
        {
            string filePath = "";
            OpenFileDialog openDig = new OpenFileDialog();
            openDig.Filter = "Excel文件|*.xls";
            DialogResult dr = openDig.ShowDialog();
            if (dr == DialogResult.OK)
            {
                _isGetPath = true;
                if (openDig.FileName != null)
                {
                    _filePath = openDig.FileName;
                    getConnStr(_filePath);
                    filePath = _filePath;
                }
            }
            else
            {
                _isGetPath = false;
            }
            return filePath;
        }

        /// <summary>
        /// 获得Excel文件路径
        /// </summary>
        public bool setFilePath(string filePath)
        {
            if (File.Exists(filePath))
            {
                _isGetPath = true;
                _filePath = filePath;
                getConnStr(_filePath);
            }
            else
            {
                _isGetPath = false;
            }
            return _isGetPath;
        }
        #endregion
    }
}
