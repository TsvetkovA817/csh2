using System;
using System.Data;
using System.Data.Odbc;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace ex2dbf
{
    internal class Program
    {
        static void Main(string[] args)
        {

            Console.WriteLine("Hello World!");
            Console.WriteLine("Обработка xls-dbf");
            Console.WriteLine("Продолжить (y/д/1)?");
            string isNext1 = Console.ReadLine();
            isNext1 = isNext1 + "   ";
            isNext1 = isNext1.Trim();
            if (isNext1 == "y" || isNext1 == "д" || isNext1 == "1")
            {
                String p1 = "";
                String p2 = "";
                String p3 = "";
                String p4 = "";
                String p5 = "";

                Excel.Application appShab = new Excel.Application();
                Excel.Workbook shabBook = appShab.Workbooks.Open("c:\\test\\sh.xlsx");
                Excel.Worksheet shab = (Excel.Worksheet)appShab.Worksheets.get_Item(1);

                clEx le = new clEx();
                le.sheet = shab;
                le.nn = true;      

                le.msg = "Данные";
                string dataXlsFile = le.GetVs(4, 4);     
                le.msg = "Файл";
                string dataDbfFile = le.GetVs(5, 4);     
                le.msg = "Бланк";
                string pathBlankFile = le.GetVs(14, 4); 

                if (pathBlankFile.Trim() != "")
                {
                    p4 = Path.GetFileName(dataDbfFile);
                    FileInfo fileInf = new FileInfo(pathBlankFile + "\\" + p4.Trim());
                    if (fileInf.Exists)
                    {
                        fileInf.CopyTo(dataDbfFile, true);
                    }
                    else
                    {
                        Console.WriteLine($"Не найден бланк {p4} в папке {pathBlankFile}");
                    }
                }
                else
                {
                    Console.WriteLine("Не указанна папка с бланком..");
                }

                le.msg = "Кол-во ячеек";
                int maxKolScanCells = le.GetVi(16, 4);

                le.msg = "Поля";
                int mRowCName = le.GetVi(7, 4);

                le.msg = "Начало диапазона";
                int mRowFrom = le.GetVi(8, 4);

                le.msg = "Конец диапазона";
                int mRowTo = le.GetVi(9, 4);

                le.msg = "от";
                int mColFrom = le.GetVi(10, 4);

                le.msg = "до";
                int mColTo = le.GetVi(11, 4);

                le.msg = "Лист";
                int mListData = le.GetVi(12, 4);

                Console.WriteLine("Открываем таблицу");
                WorkDBF WDBF = new WorkDBF();

                Console.WriteLine("Открываем данные");

                Excel.Application app = new Excel.Application();
                Excel.Workbook workBook = app.Workbooks.Open(dataXlsFile);
                Excel.Worksheet sheet = (Excel.Worksheet)app.Worksheets.get_Item(mListData);
                Console.WriteLine(sheet.Name);

                clEx ld = new clEx();
                ld.sheet = sheet;
                ld.nn = false;  

                le.msg = "авто поиск строк";
                p5 = le.GetVs(7, 6); 
                le.msg = "название столбца";
                p3 = le.GetVs(7, 7); 


                if (p5 == "auto")
                {
                    le.msg = "maxKolNullCells";
                    int maxKolNullCells = le.GetVi(17, 4);

                    le.msg = "maxKolNullDataCells";
                    int maxKolNullDataCells = le.GetVi(18, 4);

                    ld.msg = "Нет заголовка";
                    ld.nn = true;
                    adrCell adrH = new adrCell();
                    adrH = ld.GetVAdr(p3, mRowFrom, mColFrom, maxKolScanCells, maxKolNullCells);
                    mRowCName = adrH.row;
                    mRowFrom = mRowCName + 1;
                    mColFrom = adrH.col;

                    ld.msg = "Нет данных";
                    ld.nn = true;
                    mRowTo = ld.GetREnd(mRowFrom, mColFrom, maxKolScanCells, maxKolNullDataCells);
                    mColTo = ld.GetCEnd(mRowCName - 1, mColFrom, maxKolScanCells, maxKolNullDataCells);

                }
                appShab.Quit();

                Console.WriteLine($"N поля {mRowCName}");
                Console.WriteLine($"N диапазон {mRowFrom}");
                Console.WriteLine($"N до {mRowTo}");
                Console.WriteLine($"N {mColFrom}");
                Console.WriteLine($"N {mColTo}");

                Console.WriteLine("Чтобы продолжить нажмите клавишу (y/д/1)");
                string isNext2 = Console.ReadLine();
                isNext2 = isNext2 + "   ";
                isNext2 = isNext2.Trim();
                if (isNext2 != "y" && isNext2 != "д" && isNext2 != "1")
                {
                    app.Quit();
                    Environment.Exit(0);    
                }

                int mKolF = 0;      
                mKolF = ld.GetHcount(mRowCName, mColFrom, mColTo);
                var aFName = new string[mKolF];
                var aCNum = new string[mKolF];

                ld.GetHArr(ref aFName, ref aCNum, mRowCName, mColFrom, mColTo);

                Console.WriteLine("удаление предыдущие записей");
                var dt6 = WDBF.DelAll(dataDbfFile);

                Console.WriteLine("Вставляем данные");

                int c = ld.InsDt2dbf(mRowFrom, mRowTo, dataDbfFile, aFName, aCNum, ld, WDBF);
                Console.WriteLine($"Обработано: {c} записей");

                Console.WriteLine("Завершено");
                app.Quit();
                Console.ReadKey();
            }
        }

        class clEx
        {
            public Excel.Worksheet sheet;
            public string msg;
            public bool nn; 
            public string GetVs(int mR, int mC)  
            {
                string mr, p1;
                mr = ""; p1 = "";
                p1 = Convert.ToString(sheet.Cells[mR, mC].Value + "   ");
                mr = p1.Trim();
                if (mr == "" && nn == true)
                {
                    Console.WriteLine($"Не верно указано: {msg}");
                    closeApp();
                }

                return mr;
            }

            public int GetVi(int mR, int mC) 
            {
                string p1;
                int mr;
                bool n = false;
                mr = 0; p1 = "";
                p1 = Convert.ToString(sheet.Cells[mR, mC].Value + "   ");
                p1 = p1.Trim();
                if (p1 == "" && nn == true) n = true;
                if (p1 == "" && nn == false) { n = false; p1 = "0"; }
                bool success = int.TryParse(p1, out mr);
                if (!success || n == true || (nn == true && mr == 0))
                {
                    Console.WriteLine($"Не верно указано: {msg}");
                    closeApp();
                }
                return mr;
            }
            public int GetREnd(int mRowFrom, int mColFrom, int maxKolScanCells, int maxKolNullDataCells) 
            {
                string p1;
                int mret = 0; p1 = "";
                int i = 0;
                for (int mR = mRowFrom; mR <= maxKolScanCells; mR++)
                {
                    p1 = Convert.ToString(sheet.Cells[mR, mColFrom].Value + "   ");  //
                    p1 = p1.Trim();
                    if (p1 == "")
                    {
                        i = i + 1;
                        if (i >= maxKolNullDataCells)
                        {
                            i = 0;
                            mret = mR - maxKolNullDataCells;
                            break;
                        }
                    }
                    else
                    {
                        i = 0;
                    }

                } 
                if ((mret == 0 || mret <= mRowFrom || mret - mRowFrom <= 0) && nn == true)
                {
                    Console.WriteLine($"{msg}");
                    closeApp();
                }
                return mret;
            }


            public int GetCEnd(int mRowFrom, int mColFrom, int maxKolScanCells, int maxKolNullDataCells) 
            {
                string p1;
                int mret = 0; p1 = "";
                int i = 0;
                for (int mC = mColFrom; mC <= maxKolScanCells; mC++)
                {
                    p1 = Convert.ToString(sheet.Cells[mRowFrom, mC].Value + "   ");  //
                    p1 = p1.Trim();
                    if (p1 == "")
                    {
                        i = i + 1;
                        if (i >= maxKolNullDataCells)
                        {
                            i = 0;
                            mret = mC - maxKolNullDataCells;
                            break;
                        }
                    }
                    else
                    {
                        i = 0;
                    }
                } 
                if ((mret == 0 || mret <= mColFrom || mret - mColFrom <= 0) && nn == true)
                {
                    Console.WriteLine($"{msg}");
                    closeApp();
                }
                return mret;
            }

            public adrCell GetVAdr(string mFindStr, int mRowFrom, int mColFrom, int maxKolScanCells, int maxKolNullCells)  
            {
                string p1;
                string p3;
                p3 = mFindStr.Trim();
                adrCell mret = new adrCell(); p1 = "";
                mret.row = 0;
                mret.col = 0;
                bool emp = true;
                bool mex = false;
                int i = 0;
                for (int mC = 1; mC <= maxKolScanCells; mC++)
                {
                    for (int mR = 1; mR <= maxKolScanCells; mR++)
                    {
                        p1 = Convert.ToString(sheet.Cells[mR, mC].Value + "   ");  //
                        p1 = p1.Trim();
                        if (p1 == "")
                        {
                            i = i + 1;
                            if (i > maxKolNullCells)
                            {
                                i = 0;
                                break;
                            }
                        }
                        else
                        {
                            i = 0;
                        }
                        if (p1 != "" && p1 == p3)
                        {
                            mret.row = mR;
                            mret.col = mC;
                            mex = true;
                            emp = false;
                            break;
                        }
                    }
                    if (mex)
                    {
                        break;
                    }
                }
                if (emp && nn == true)
                {
                    Console.WriteLine($"{msg}");
                    closeApp();
                }
                return mret;
            }

            public void closeApp()
            {
                Console.WriteLine("Завершено");
                Console.ReadKey();
                Environment.Exit(0);
            }

            public int GetHcount(int mRowCName, int mColFrom, int mColTo)  
            {
                string p1;
                int mr = 0; p1 = "";
                for (int i = mColFrom; i <= mColTo; i++)
                {
                    p1 = Convert.ToString(sheet.Cells[mRowCName, i].Value + "   ");
                    p1 = p1.Trim();
                    if (p1 != "" && p1 != "end") mr = mr + 1;
                }
                if (mr == 0 && nn == true)
                {
                    Console.WriteLine($"Не верно указано: {msg}");
                    closeApp();
                }

                return mr;
            }

            public string GetSqlIns(string[] aFName, string[] aCNum, int mNumRecData, string dataDbfFile)
            {
                string p1, p3, p2;
                string mr = "";
                string ml = "";
                string mz = "";
                p1 = ""; p3 = ""; p2 = "";
                ml = "";
                mz = "";
                int iFInArr = 0;
                int mNumCol = 0;

                for (int j = 0; j < aFName.Length; j++)
                {
                    p3 = aFName[j] + "    ";
                    if (p3.Trim() != "")
                    {
                        p3 = p3.Trim();
                        iFInArr = Array.IndexOf(aFName, p3);

                        mNumCol = Convert.ToInt32(aCNum[iFInArr]);
                        p1 = Convert.ToString(sheet.Cells[mNumRecData, mNumCol].Value);
                        if (p1 != null)
                        {
                            if (ml == "") ml = p3;
                            else ml = ml + "," + p3;
                            p2 = "'" + p1 + "'";
                            if (mz == "") mz = mz + p2;
                            else mz = mz + "," + p2;
                        }
                    }
                }
                mr = $"INSERT INTO {dataDbfFile} ({ml}) VALUES({mz})";
                return mr;
            }

            public int InsDt2dbf(int mRowFrom, int mRowTo, string dataDbfFile, string[] aFName, string[] aCNum, clEx ld, WorkDBF wdbf)  
            {
                string p1, mq;
                int mr = 0; p1 = "";
                for (int i = mRowFrom; i <= mRowTo; i++)
                {
                    mq = ld.GetSqlIns(aFName, aCNum, i, dataDbfFile);
                    var dt5 = wdbf.Execute(mq);
                    mr = mr + 1;
                    Console.WriteLine($"Строка: {i}");
                }
                return mr;
            }

            public void GetHArr(ref string[] aFName, ref string[] aCNum, int mRowCName, int mColFrom, int mColTo)
            {
                string p1 = "";
                int l = 0;
                for (int k = mColFrom; k < mColTo; k++)
                {
                    p1 = Convert.ToString(sheet.Cells[mRowCName, k].Value + "   ");
                    p1 = p1.Trim();
                    if (p1 != "" && p1 != "end")
                    {
                        aFName[l] = p1.ToString();
                        aCNum[l] = k.ToString();
                        l = l + 1;
                    }
                }
            }
        }

        struct adrCell
        {
            public int row;
            public int col;
        }

        class WorkDBF
        {
            private OdbcConnection Conn = null;

            public WorkDBF()
            {
                this.Conn = new System.Data.Odbc.OdbcConnection();
                Conn.ConnectionString = @"Driver={Microsoft dBase Driver (*.dbf)};" +
                       "SourceType=DBF;Exclusive=No;" +
                       "Collate=Machine;NULL=NO;DELETED=NO;" +
                       "BACKGROUNDFETCH=NO;";

            }

            public DataTable Execute(string Command)
            {
                DataTable dt = null;
                if (Conn != null)
                {
                    try
                    {
                        Conn.Open();
                        dt = new DataTable();
                        System.Data.Odbc.OdbcCommand oCmd = Conn.CreateCommand();
                        oCmd.CommandText = Command;
                        dt.Load(oCmd.ExecuteReader());
                        Conn.Close();
                    }
                    catch (Exception e)
                    {
                        if (Conn != null) Conn.Close();
                        Console.WriteLine(e.Message);
                    }
                }
                return dt;
            }
            public DataTable GetAll(string DB_path)
            {
                return Execute("SELECT * FROM " + DB_path);
            }

            public DataTable DelAll(string DB_path)
            {
                string mq = $"DELETE FROM {DB_path}";
                return Execute(mq);
            }
        }
    }
}


