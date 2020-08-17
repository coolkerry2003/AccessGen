using ADOX;
using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AccessGen
{
    class Program
    {
        private static string mRootDir = @"D:\Tools\AccessGen\";
        private static string mSrcDir = Path.Combine(mRootDir, "Src");
        private static string mOutDir = Path.Combine(mRootDir, "Output");
        private static string mDataSource = Path.Combine(mOutDir, "LIS.mdb");
        private static StreamWriter pLogFile;
        static void Main(string[] args)
        {
            try
            {
                if (!Directory.Exists(mOutDir))
                    Directory.CreateDirectory(mOutDir);
                if (Directory.Exists(mSrcDir))
                {
                    OpenLogFile();
                    AccessGen aGen = new AccessGen();
                    if (aGen.Create(mDataSource))
                        Console.Write("建立資料庫成功！\n");
                    else
                        Console.Write("失敗\n");

                    aGen.Open();
                    ReadExcel(aGen);
                    aGen.Close();
                    CloseLogFile();
                }
            }
            catch (Exception ex)
            {
                string err = string.Format("{0}\n{1}\n", ex.Message, ex.StackTrace);
                WriteLog(err, true);
            }
            Console.Read();
        }
        /// <summary>
        /// 讀取Excel
        /// 2003舊版xls檔案
        /// </summary>
        /// <param name="aGen"></param>
        public static void ReadExcel(AccessGen aGen)
        {
            string[] mPaths = Directory.GetFiles(mSrcDir, "*.xls");
            Console.Write(string.Format("共{0}份檔案，開始將Excel資料匯入至DB...\n", mPaths.Length));
            for (int i = 0; i < mPaths.Length; i++)
            {
                string mPath = mPaths[i];
                string mFileName = Path.GetFileNameWithoutExtension(mPath);
                ExcelRead mER = new ExcelRead(mPath);
                mER.GetDataSet();
                if (mER.mCols.Count <= 0)
                    continue;

                //篩選欄位
                List<string> Cols = new List<string>();
                foreach (object col in mER.mCols)
                {
                    string str = col.ToString().Replace("\'", "");//替換特殊字元
                    Cols.Add(string.Format("[{0}]  Text", str));
                }
                //建立資料表
                string mTName = mFileName;
                string mSQL = string.Join(",", Cols);
                Console.Write("建立資料表中...");
                if (aGen.CreateTable(mTName, mSQL))
                    Console.Write("成功！\n");
                else
                    Console.Write("失敗！\n");

                for (int j = 0; j < mER.mDataRows.Count; j++)
                {
                    DataRowCollection mDataRowC = mER.mDataRows[j];
                    for (int k = 0; k < mDataRowC.Count; k++)
                    {
                        Console.Write(string.Format("[{0}/{1}] > ",i+1, mPaths.Length));
                        Console.Write(string.Format("輸入第{0}/{1}張表 > [{2}] > ", j + 1, mER.mDataRows.Count, mTName));
                        Console.Write(string.Format("輸入資料{0}/{1}...進度{2}%...", k + 1, mDataRowC.Count, Math.Round((double)k * 100 / (double)mDataRowC.Count, 1, MidpointRounding.AwayFromZero)));

                        DataRow mDataRow = mDataRowC[k];
                        //篩選資料
                        List<string> Rows = new List<string>();
                        foreach (object row in mDataRow.ItemArray.ToList())
                        {
                            string str = row.ToString().Replace("\'", "");//替換特殊字元
                            Rows.Add(string.Format("\'{0}\'", str));
                        }
                        //輸入資料
                        string aInsetSQL = string.Format("INSERT INTO {0} VALUES ({1})", mTName, string.Join(",", Rows));
                        if (aGen.ExecuteNonQuery(aInsetSQL))
                            Console.Write("成功\n");
                        else
                            Console.Write("失敗\n");
                    }
                }
            }
            Console.Write("完成！\n");
        }
        /// <summary>
        /// 開啟Log檔案
        /// </summary>
        public static void OpenLogFile()
        {
            string path = string.Format("{0}\\.AccessGen.log", mOutDir);
            //清空
            File.WriteAllText(path, string.Empty, Encoding.Default);
            FileStream fs = new FileStream(path, FileMode.OpenOrCreate);
            pLogFile = new StreamWriter(fs, Encoding.Default);
        }
        /// <summary>
        /// 關閉Log檔案
        /// </summary>
        public static void CloseLogFile()
        {
            if (pLogFile != null)
            {
                pLogFile.Close();
            }
        }
        /// <summary>
        /// 寫入Log
        /// </summary>
        /// <param name="_Str"></param>
        public static void WriteLog(string a_Message, bool a_ShowTime)
        {
            string msg = a_ShowTime ? string.Format("[{0}] {1}", DateTime.Now.ToString("HH:mm:ss"), a_Message) : a_Message;
            Console.Write(msg);
            pLogFile.Write(msg);
            pLogFile.Flush();
        }
    }
    class AccessGen
    {
        CatalogClass mCat = new CatalogClass();
        OleDbConnection mConn = new OleDbConnection();
        OleDbTransaction mTrans;
        string m_ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Jet OLEDB:Engine Type=5";
        string mErrorMsg = string.Empty;
        public AccessGen()
        {

        }
        public bool Create(string aDataSource)
        {
            bool success = false;
            try
            {
            Begin:
                if (!File.Exists(aDataSource))
                {
                    m_ConnStr = string.Format(m_ConnStr, aDataSource);
                    mCat.Create(m_ConnStr);
                    mConn.ConnectionString = m_ConnStr;
                    success = true;
                }
                else
                {
                    Console.Write(string.Format("此資料庫檔案已存在：{0}\n", aDataSource));

                Back:
                    Console.Write("是否刪除此檔案(Y/N)?");
                    string mCmd = Console.ReadLine();
                    switch (mCmd.ToUpper())
                    {
                        case "Y":
                            File.Delete(aDataSource);
                            goto Begin;
                            break;
                        case "N":
                            m_ConnStr = string.Format(m_ConnStr, aDataSource);
                            mConn.ConnectionString = m_ConnStr;
                            return false;
                            break;
                        default:
                            goto Back;
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                mErrorMsg = ex.Message;
                success = false;
            }
            return success;
        }
        public void Open()
        {
            if(mConn.State != ConnectionState.Open)
                mConn.Open();
        }
        public void Close()
        {
            mConn.Close();
        }
        public bool CreateTable(string aTName,string aSQL)
        {
            bool success = false;
            try
            {
                if (!ExistTable(aTName))
                {
                    OleDbCommand myCommand = new OleDbCommand();
                    myCommand.Connection = mConn;
                    myCommand.CommandText = string.Format("CREATE TABLE {0} ({1})", aTName, aSQL);
                    myCommand.ExecuteNonQuery();
                }
                else
                {
                    Console.Write("資料表已存在！");
                }
            }
            catch (Exception ex)
            {
                mErrorMsg = ex.Message;
                success = false;
            }
            return success;
        }
        public bool ExistTable(string _Table)
        {
            bool success = false;
            string sql = string.Format("Select Count(*) From {0}", _Table);
            OleDbDataReader r = ExecuteQuery(sql);
            if (r != null)
            {
                if (r.Read())
                {
                    success = true;
                }
            }
            return success;
        }
        public bool ExecuteNonQuery(string SQLStr)
        {
            bool success = false;
            try
            {
                OleDbCommand Command = mConn.CreateCommand();
                mTrans = mConn.BeginTransaction();
                Command.Transaction = mTrans;
                Command.CommandText = SQLStr;
                Command.ExecuteNonQuery();
                mTrans.Commit();
                success = true;
            }
            catch (Exception ex)
            {
                mTrans.Rollback();
                mErrorMsg = ex.Message;
                success = false;
            }
            return success;
        }
        public OleDbDataReader ExecuteQuery(string SQLStr)
        {
            OleDbDataReader Rtn = null;
            try
            {
                OleDbCommand Command = mConn.CreateCommand();
                Command.CommandText = SQLStr;
                Rtn = Command.ExecuteReader();
            }
            catch (Exception ex)
            {
                mErrorMsg = ex.Message;
            }
            return Rtn;
        }
    }
    class ExcelRead
    {
        string mPath = string.Empty;
        bool isFirstField = true;
        public List<object> mCols = new List<object>();
        public List<DataRowCollection> mDataRows = new List<DataRowCollection>();
        public ExcelRead(string aPath)
        {
            mPath = aPath;
        }
        public void GetDataSet()
        {
            try
            {
                Console.Write(string.Format("取得資料表：[{0}]...", mPath));
                using (var stream = File.Open(mPath, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        Console.Write("成功！\n");

                        DataSet mSheet = reader.AsDataSet();
                        foreach (DataTable mTable in mSheet.Tables)
                        {
                            Console.Write(string.Format("讀取[{0}]表...", mTable.TableName));
                            if (mTable.Rows.Count > 0)
                            {
                                if (isFirstField)
                                {
                                    mCols = mTable.Rows[0].ItemArray.ToList();
                                    mTable.Rows.RemoveAt(0);
                                }
                                mDataRows.Add(mTable.Rows);
                                Console.Write(string.Format("成功！共{0}個資料列\n", mTable.Rows.Count));
                            }
                            else
                            {
                                Console.Write("無資料\n");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
                //string err = string.Format("{0}\n", ex.ToString());
                //Console.WriteLine(err);
            }
        }
    }
}
