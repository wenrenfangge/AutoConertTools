using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Data.SqlTypes;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace AutoConvertToText
{
    public enum OffceVersion
    {
        Office2007 = 1,
        Office2010 = 2,
        Office2016 = 3,
    }
    class ExcelToTxt
    {
        private static ExcelToTxt _instance;
        public static ExcelToTxt Instance
        {
            get 
            {
                if(_instance == null)
                {
                    _instance = new ExcelToTxt();
                }
                return _instance;
            }
        }
        public string[] selectionItems = { "Office2007", "Office2010", "Office2016" };
        static bool _isVerbose = false;
        static bool _isAllian = false;

        // 获得字段的实际最大长度
        static int GetMaxLength(DataTable dt, string captionName)
        {
            DataColumn maxLengthColumn = new DataColumn();
            maxLengthColumn.ColumnName = "MaxLength";
            maxLengthColumn.Expression = String.Format("len(convert({0},'System.String'))", captionName);
            dt.Columns.Add(maxLengthColumn);
            object maxLength = dt.Compute("max(MaxLength)", "true");
            if (maxLength == DBNull.Value)
            {
                return 0;
            }
            dt.Columns.Remove(maxLengthColumn);

            return Convert.ToInt32(maxLength);
        }

        static void convertExcelToTxt(string inputFile, string outputPath)
        {
            if (Path.GetExtension(inputFile) != ".xls" && Path.GetExtension(inputFile) != ".xlsx")
            {
                return;
            }

            if (!Directory.Exists(outputPath))
            {
                Directory.CreateDirectory(outputPath);
            }

            //string newFileNameNoExt = Path.GetFileNameWithoutExtension(inputFile);
            //string newFileNoExt = outputPath + "\\" + newFileNameNoExt;
            //string newFile = newFileNoExt + ".txt";
            //Console.WriteLine("Convert file[{0}] to [{1}]", inputFile, newFile);

            var conn = new OleDbConnection();

            conn.ConnectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;" +
                @"Data Source={0}" +
                ";Extended Properties=\"Excel 12.0 Xml;HDR=No;IMEX=1\"", inputFile);

            //修复不能全部读取的问题            
            //conn.ConnectionString = "Provider = Microsoft.Jet.OLEDB.4.0;" + "Data Source= " + inputFile + " ;"
            //+ " Extended Properties = Excel 8.0;";
            conn.Open();
            DataTable sheetTb = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            foreach (DataRow sheet in sheetTb.Rows)
            {
                string tableName = sheet["TABLE_NAME"].ToString();

                string sql = String.Format("select * from [{0}]", tableName);
                OleDbDataAdapter da = new OleDbDataAdapter(sql, conn);

                var ds = new DataSet();
                da.Fill(ds);

                var tb1 = ds.Tables[0];

                if (tb1.Rows.Count == 0)
                {
                    continue; // 空表
                }
                if (tb1.Rows.Count == 1 && tb1.Columns.Count == 1)
                {
                    if (tb1.Rows[0][0] == DBNull.Value)
                    {
                        continue; // 空表
                    }
                }
                //modify
                string newFileNoExt = outputPath + "\\" + tableName;
                string newFile = newFileNoExt + ".txt";
                Console.WriteLine("Convert file[{0}] to [{1}]", inputFile, newFile);
                FileStream fs = new FileStream(newFileNoExt.Trim('$') + ".txt", FileMode.OpenOrCreate);
                StreamWriter sw = new StreamWriter(fs);

                int[] colMaxLen = new int[tb1.Columns.Count];

                if (_isAllian)
                {
                    for (int i = 0; i < tb1.Columns.Count; ++i)
                    {
                        colMaxLen[i] = 0;
                        for (int j = 0; j < tb1.Rows.Count; ++j)
                        {
                            string s = tb1.Rows[j][i].ToString();
                            int len = System.Text.Encoding.Default.GetBytes(s).Length;
                            if (len > colMaxLen[i])
                            {
                                colMaxLen[i] = len;
                            }
                        }

                    }
                }


                foreach (DataRow row in tb1.Rows)
                {
                    for (int j = 0; j < tb1.Columns.Count; ++j)
                    {
                        DataColumn col = tb1.Columns[j];
                        string content = row[j].ToString();

                        bool hasYinhao = false;
                        if (-1 != content.IndexOf("\r") || -1 != content.IndexOf("\n"))
                        {
                            hasYinhao = true;
                        }

                        string fmt;
                        if (_isAllian)
                        {
                            int realLen = colMaxLen[j] - (System.Text.Encoding.Default.GetBytes(content).Length - content.Length);
                            // "{0,-10}"\t
                            fmt = String.Format("{0}{1}0,-{2}{3}{4}{5}", hasYinhao ? "\"" : "",
                            "{", realLen, "}", hasYinhao ? "\"" : "", j + 1 == tb1.Columns.Count ? "" : "\t");
                        }
                        else
                        {
                            // "{0}"\t
                            fmt = String.Format("{0}{1}0{2}{3}{4}", hasYinhao ? "\"" : "",
                            "{", "}", hasYinhao ? "\"" : "", j + 1 == tb1.Columns.Count ? "" : "\t");
                        }

                        sw.Write(fmt, row[j]);
                        if (_isVerbose)
                        {
                            Console.Write(fmt, row[j]);
                        }
                    }
                    sw.WriteLine();
                    if (_isVerbose)
                    {
                        Console.WriteLine();
                    }

                }
                sw.Close();
            }
            conn.Close();
        }

        static void loopDir(string inputDir, string outDir)
        {
            DirectoryInfo di = new DirectoryInfo(inputDir);
            var files = di.GetFiles();
            foreach (var file in files)
            {
                string srcPath = file.FullName;
                string dstPath = outDir;
                // Console.WriteLine(srcPath + "         " + dstPath);
                convertExcelToTxt(srcPath, dstPath);
            }

            var infos = di.GetDirectories();
            foreach (var info in infos)
            {
                string srcPath = info.FullName;
                string dstPath = outDir + "\\" + info.Name;
                loopDir(srcPath, dstPath);
            }
        }

        public static int RunConvertTool(string[] args)
        {
            if (args.Length < 1)
            {
                Console.WriteLine("Help: \r\n" +
                    " excel_convert InputDir [OutputDir] /v /a\r\n" +
                    " InputDir 必选，输入目录。 将遍历此目录下的所有.xls 和.xlsx文件进行转换 \r\n" +
                    " Output   可选，输出目录。 如果没有使用和InputDir同一目录 \r\n" +
                    " /v       可选，显示转换信息。" +
                    " /a       可选，对齐列打印。");
                return -1;
            }
            string inputDir = args[0];
            if (!Directory.Exists(inputDir))
            {
                Console.WriteLine(inputDir);
                Console.WriteLine("No input directory exist, {0}", inputDir);
                return -1;
            }

            string outputPath;
            if (args.Length > 1)
            {
                outputPath = args[1];
            }
            else
            {
                outputPath = inputDir;
            }

            for (int i = 2; i < args.Length; ++i)
            {
                string verbose = args[i];
                verbose = verbose.ToLower();
                if (verbose == "/v")
                {
                    _isVerbose = true;
                }
                if (verbose == "/a")
                {
                    _isAllian = true;
                }
            }
            //outputPath = inputDir;
            loopDir(inputDir, outputPath);
            return 0;
        }
    }
}
