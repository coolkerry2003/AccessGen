using ADOX;
using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DAO = Microsoft.Office.Interop.Access.Dao;

namespace AccessGen
{
    class Program
    {
        private static string mRootDir = ConfigurationSettings.AppSettings["mRootDir"].ToString();
        private static string mSrcDir = Path.Combine(mRootDir, ConfigurationSettings.AppSettings["mSrcDir"].ToString());
        private static string mOutDir = Path.Combine(mRootDir, ConfigurationSettings.AppSettings["mOutDir"].ToString());
        private static string mDataSource = Path.Combine(mOutDir, ConfigurationSettings.AppSettings["mDataSource"].ToString());
        private static StreamWriter pLogFile;
        static void Main(string[] args)
        {
            try
            {
                Directory.CreateDirectory(mOutDir);
                if (Directory.Exists(mSrcDir))
                {
                    OpenLogFile();
                    Stopwatch sw = new Stopwatch();
                    sw.Restart();
                    sw.Start();

                    DAOGen aGen = new DAOGen();
                    if (aGen.Create(mDataSource))
                    {
                        //aGen.Open();
                        ReadExcel(aGen);
                    }
                    else
                    {
                        Console.Write("失敗\n");
                        WriteLog(aGen.GetLastError(), true);
                    }

                    WriteLog(string.Format("花費：{0}分", sw.Elapsed.TotalMinutes.ToString()), true);
                    aGen.Close();
                    CloseLogFile();
                }
                else
                {
                    Console.Write("來源資料不存在\n");
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
                //取得Excel資料
                mER.GetDataSet();
                if (mER.mCols.Count <= 0) continue;
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
                {
                    Console.Write("成功！\n");
                    //依據表單建立資料
                    for (int j = 0; j < mER.mDataRows.Count; j++)
                    {
                        DataRowCollection mDataRowC = mER.mDataRows[j];
                        for (int k = 0; k < mDataRowC.Count; k++)
                        {
                            Console.Write($"[{i + 1}/{mPaths.Length}] > ");
                            Console.Write($"[{mTName}] > 輸入第{j + 1}/{mER.mDataRows.Count}張表 > ");
                            Console.Write($"輸入資料{k + 1}/{mDataRowC.Count}...進度{Math.Round((double)k * 100 / (double)mDataRowC.Count, 1, MidpointRounding.AwayFromZero)}%...");

                            //篩選資料
                            List<string> Rows = new List<string>();
                            DataRow mDataRow = mDataRowC[k];
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
                            {
                                Console.Write("失敗\n");
                                WriteLog(aGen.GetLastError(), true);
                            }
                        }
                    }
                }
                else
                {
                    Console.Write("失敗！\n");
                    WriteLog(aGen.GetLastError(), true);
                }
            }
            Console.Write($"資料庫建置完成！\n");
        }
        /// <summary>
        /// 讀取Excel
        /// 2003舊版xls檔案
        /// </summary>
        /// <param name="aGen"></param>
        public static void ReadExcel(DAOGen aGen)
        {
            string[] mPaths = Directory.GetFiles(mSrcDir, "*.xls");
            Console.Write(string.Format("共{0}份檔案，開始將Excel資料匯入至DB...\n", mPaths.Length));
            for (int i = 0; i < mPaths.Length; i++)
            {
                string mPath = mPaths[i];
                string mFileName = Path.GetFileNameWithoutExtension(mPath);
                ExcelRead mER = new ExcelRead(mPath);
                //取得Excel資料
                mER.GetDataSet();
                if (mER.mCols.Count <= 0) continue;
                //篩選欄位
                List<string> Cols = new List<string>();
                foreach (object col in mER.mCols)
                {
                    string str = col.ToString().Replace("\'", "");//替換特殊字元
                    Cols.Add(str);
                }
                //建立資料表
                string mTName = mFileName;
                DAO.Recordset mRs = null;
                DAO.Field[] mFields = null;
                Console.Write("建立資料表中...");
                if (aGen.CreateTable(mTName, Cols.ToArray()) && aGen.OpenTable(mTName, Cols.ToArray(), ref mRs, ref mFields))
                {
                    Console.Write("成功！\n");

                    Console.Write("建立欄位資料中...");
                    //依據表單建立資料
                    for (int j = 0; j < mER.mDataRows.Count; j++)
                    {
                        DataRowCollection mDataRowC = mER.mDataRows[j];
                        for (int k = 0; k < mDataRowC.Count; k++)
                        {
                            //Console.Write($"[{i + 1}/{mPaths.Length}] > ");
                            //Console.Write($"[{mTName}] > 輸入第{j + 1}/{mER.mDataRows.Count}張表 > ");
                            //Console.Write($"輸入資料{k + 1}/{mDataRowC.Count}...進度{Math.Round((double)k * 100 / (double)mDataRowC.Count, 1, MidpointRounding.AwayFromZero)}%...");

                            //篩選資料
                            List<string> Rows = new List<string>();
                            DataRow mDataRow = mDataRowC[k];
                            foreach (object row in mDataRow.ItemArray.ToList())
                            {
                                string str = row.ToString().Replace("\'", "");//替換特殊字元
                                str = string.IsNullOrEmpty(str) ? " " : str;
                                Rows.Add(str);
                            }

                            //輸入資料
                            if (!aGen.Insert(mRs, mFields, Rows))
                            {
                                WriteLog(aGen.GetLastError(), true);
                            }
                        }
                    }

                    //關閉資料表
                    if (mRs != null)
                    {
                        mRs.Close();
                    }
                    Console.Write("成功！\n");
                }
                else
                {
                    Console.Write("失敗！\n");
                    WriteLog(aGen.GetLastError(), true);
                }
            }
            Console.Write($"資料庫建置完成！\n");
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
        /// <summary>
        /// 建立資料庫
        /// </summary>
        /// <param name="aDataSource"></param>
        /// <returns></returns>
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
                            return true;
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
        /// <summary>
        /// 開啟資料庫
        /// </summary>
        public void Open()
        {
            if(mConn.State != ConnectionState.Open)
                mConn.Open();
        }
        /// <summary>
        /// 關閉資料庫
        /// </summary>
        public void Close()
        {
            mConn.Close();
        }
        /// <summary>
        /// 建立資料表
        /// </summary>
        /// <param name="aTName"></param>
        /// <param name="aSQL"></param>
        /// <returns></returns>
        public bool CreateTable(string aTName,string aSQL)
        {
            bool success = true;
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
        /// <summary>
        /// 檢核資料表是否存在
        /// </summary>
        /// <param name="_Table"></param>
        /// <returns></returns>
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
        /// <summary>
        /// 執行SQL語法
        /// </summary>
        /// <param name="SQLStr"></param>
        /// <returns></returns>
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
        /// <summary>
        /// 執行SQL語法
        /// </summary>
        /// <param name="SQLStr"></param>
        /// <returns></returns>
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
        /// <summary>
        /// 取得最後錯誤紀錄
        /// </summary>
        /// <returns></returns>
        public string GetLastError()
        {
            return mErrorMsg;
        }
    }
    class DAOGen
    {
        string mDbLangGeneral = ";LANGID=0x0409;CP=1252;COUNTRY=0";
        DAO.Database mDB = null;
        string mErrorMsg = string.Empty;
        /// <summary>
        /// 建立資料庫
        /// </summary>
        /// <param name="aDataSource"></param>
        /// <returns></returns>
        public bool Create(string aDataSource)
        {
            bool success = false;
            try
            {
                Begin:
                if (!File.Exists(aDataSource))
                {
                    DAO.DBEngine dbEngine = new DAO.DBEngine();
                    mDB = dbEngine.CreateDatabase(aDataSource, mDbLangGeneral);
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
                            DAO.DBEngine dbEngine = new DAO.DBEngine();
                            mDB = dbEngine.OpenDatabase(aDataSource);
                            return true;
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
        /// <summary>
        /// 關閉資料庫
        /// </summary>
        public void Close()
        {
            if (mDB != null)
            {
                mDB.Close();
            }
        }
        /// <summary>
        /// 建立資料表
        /// </summary>
        /// <param name="aTName"></param>
        /// <param name="aCols"></param>
        public bool CreateTable(string aTName, string[] aCols)
        {
            bool success = true;
            try
            {
                if (!ExistTable(aTName))
                {
                    DAO.TableDef t1 = mDB.CreateTableDef(aTName);
                    foreach (string col in aCols)
                    {
                        DAO.Field f1 = t1.CreateField(col, DAO.DataTypeEnum.dbText);
                        t1.Fields.Append(f1);
                    }
                    mDB.TableDefs.Append(t1);
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
        /// <summary>
        /// 檢核資料表是否存在
        /// </summary>
        /// <param name="_Table"></param>
        /// <returns></returns>
        public bool ExistTable(string aTName)
        {
            try
            {
                DAO.Recordset rs = mDB.OpenRecordset(aTName);
                return (rs != null);
            }
            catch
            {
                return false;
            }
        }
        /// <summary>
        /// 開啟資料表
        /// </summary>
        /// <param name="aTName"></param>
        /// <param name="aCols"></param>
        /// <returns></returns>
        public bool OpenTable(string aTName, string[] aCols,ref DAO.Recordset aRs, ref DAO.Field[] aField)
        {
            try
            {
                DAO.Recordset rs = mDB.OpenRecordset(aTName);
                DAO.Field[] fileds = new DAO.Field[aCols.Length];
                for (int i = 0; i < aCols.Length; i++)
                {
                    fileds[i] = rs.Fields[aCols[i]];
                }
                aRs = rs;
                aField = fileds;
                return true;
            }
            catch (Exception ex)
            {
                mErrorMsg = ex.Message;
                return false;
            }
        }
        /// <summary>
        /// 插入資料
        /// </summary>
        /// <param name="aRs"></param>
        /// <param name="aField"></param>
        /// <param name="aRows"></param>
        public bool Insert(DAO.Recordset aRs, DAO.Field[] aField, List<string> aRows)
        {
            bool success = true;
            try
            {
                aRs.AddNew();
                for (int i = 0; i < aField.Length; i++)
                {
                    aField[i].Value = aRows[i];
                }
                aRs.Update();
            }
            catch (Exception ex)
            {
                mErrorMsg = ex.Message;
                success = false;
            }
            return success;
        }
        /// <summary>
        /// 取得最後錯誤紀錄
        /// </summary>
        /// <returns></returns>
        public string GetLastError()
        {
            return mErrorMsg;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="aPath"></param>
        /// <returns></returns>
        public bool test(string aPath)
        {
            if (File.Exists(aPath))
                File.Delete(aPath);
            string databasePath = aPath;
            string dbLangGeneral = ";LANGID=0x0409;CP=1252;COUNTRY=0";
            DAO.DBEngine dbEngine = new DAO.DBEngine();
            DAO.Database db = dbEngine.CreateDatabase(databasePath, dbLangGeneral);

            DAO.TableDef t1 = db.CreateTableDef("F_Table");
            DAO.Field f1 = t1.CreateField("F_TYPE", DAO.DataTypeEnum.dbText, 255);
            DAO.Field f2 = t1.CreateField("F_TAG", DAO.DataTypeEnum.dbText, 255);

            t1.Fields.Append(f1);
            t1.Fields.Append(f2);
            db.TableDefs.Append(t1);


            //
            //DAO.Database db = dbEngine.OpenDatabase(databasePath);

            DAO.Recordset rs = db.OpenRecordset("F_Table");
            DAO.Field[] myFields = new DAO.Field[2];
            myFields[0] = rs.Fields["F_TYPE"];
            myFields[1] = rs.Fields["F_TAG"];

            for (int i = 0; i < 1000; i++)
            {
                rs.AddNew();
                myFields[0].Value = "test" + i;
                myFields[1].Value = "test" + i;
                rs.Update();
            }

            rs.Close();
            db.Close();

            return true;
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
        /// <summary>
        /// 取得Excel資料內容
        /// </summary>
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
