using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data;
using SD = System.Data;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;


namespace test1
{
    internal class Pke
    {
        private readonly List<ParamObject> _listParamObject;

        public string MessageException{ get; set; }


        public int Length => _listParamObject.Count;


        public static SD.DataTable CreateParamTable()
        {
            SD.DataTable dt = new SD.DataTable("ParamObject");

            DataColumn colNameObject = new DataColumn("Имя объекта", typeof(string));
            DataColumn colTimeStart = new DataColumn("Время старка", typeof(string));
            DataColumn colTimeStop = new DataColumn("Время остановки", typeof(string));
            DataColumn colActiveCxema = new DataColumn("Схема проверки", typeof(int));
            DataColumn colAveragingIntervalTime = new DataColumn("Интервал времени", typeof(string));

            dt.Columns.Add(colNameObject);
            dt.Columns.Add(colTimeStart);
            dt.Columns.Add(colTimeStop);
            dt.Columns.Add(colActiveCxema);
            dt.Columns.Add(colAveragingIntervalTime);

            return dt;
        }


        private SD.DataTable CreateResultTable(int activeCxema)
        {
            SD.DataTable dt = new SD.DataTable("ResultTable");

            if (activeCxema == 1)
            {
                DataColumn colTimeTek = new DataColumn("Дата / время", typeof(string));
                DataColumn colUa = new DataColumn("UA", typeof(string));
                DataColumn colIa = new DataColumn("IA", typeof(string));
                DataColumn colPa = new DataColumn("PA", typeof(string));
                DataColumn colQa = new DataColumn("QA", typeof(string));
                DataColumn colSa = new DataColumn("SA", typeof(string));
                DataColumn colFreq = new DataColumn("Freq", typeof(string));
                DataColumn colSigmaUy = new DataColumn("sigmaUy", typeof(string));

                dt.Columns.Add(colTimeTek);
                dt.Columns.Add(colUa);
                dt.Columns.Add(colIa);
                dt.Columns.Add(colPa);
                dt.Columns.Add(colQa);
                dt.Columns.Add(colSa);
                dt.Columns.Add(colFreq);
                dt.Columns.Add(colSigmaUy);
            }
            if (activeCxema == 2)
            {
                DataColumn colTimeTek = new DataColumn("Дата / время", typeof(string));
                DataColumn colUab = new DataColumn("UAB", typeof(string));
                DataColumn colUbc = new DataColumn("UBC", typeof(string));
                DataColumn colUca = new DataColumn("UCA", typeof(string));
                DataColumn colIab = new DataColumn("IAB", typeof(string));
                DataColumn colIbc = new DataColumn("IBC", typeof(string));
                DataColumn colIca = new DataColumn("ICA", typeof(string));
                DataColumn colIa = new DataColumn("IA", typeof(string));
                DataColumn colIb = new DataColumn("IB", typeof(string));
                DataColumn colIc = new DataColumn("IC", typeof(string));
                DataColumn colPo = new DataColumn("PO", typeof(string));
                DataColumn colPp = new DataColumn("PP", typeof(string));
                DataColumn colQo = new DataColumn("QO", typeof(string));
                DataColumn colQp = new DataColumn("QP", typeof(string));
                DataColumn colSo = new DataColumn("SO", typeof(string));
                DataColumn colSp = new DataColumn("SP", typeof(string));
                DataColumn colUo = new DataColumn("UO", typeof(string));
                DataColumn colUp = new DataColumn("UP", typeof(string));
                DataColumn colIo = new DataColumn("IO", typeof(string));
                DataColumn colIp = new DataColumn("IP", typeof(string));
                DataColumn colKo = new DataColumn("KO", typeof(string));
                DataColumn colFreq = new DataColumn("Freq", typeof(string));
                DataColumn colSigmaUy = new DataColumn("SigmaUy", typeof(string));
                DataColumn colSigmaUyAb = new DataColumn("SigmaUyAB", typeof(string));
                DataColumn colSigmaUyBc = new DataColumn("SigmaUyBC", typeof(string));
                DataColumn colSigmaUyCa = new DataColumn("SigmaUyCA", typeof(string));

                dt.Columns.Add(colTimeTek);
                dt.Columns.Add(colUab);
                dt.Columns.Add(colUbc);
                dt.Columns.Add(colUca);
                dt.Columns.Add(colIab);
                dt.Columns.Add(colIbc);
                dt.Columns.Add(colIca);
                dt.Columns.Add(colIa);
                dt.Columns.Add(colIb);
                dt.Columns.Add(colIc);
                dt.Columns.Add(colPo);
                dt.Columns.Add(colPp);
                dt.Columns.Add(colQo);
                dt.Columns.Add(colQp);
                dt.Columns.Add(colSo);
                dt.Columns.Add(colSp);
                dt.Columns.Add(colUo);
                dt.Columns.Add(colUp);
                dt.Columns.Add(colIo);
                dt.Columns.Add(colIp);
                dt.Columns.Add(colKo);
                dt.Columns.Add(colFreq);
                dt.Columns.Add(colSigmaUy);
                dt.Columns.Add(colSigmaUyAb);
                dt.Columns.Add(colSigmaUyBc);
                dt.Columns.Add(colSigmaUyCa);
            }
            if (activeCxema == 3)
            {
                DataColumn colTimeTek = new DataColumn("Дата / время", typeof(string));
                DataColumn colUab = new DataColumn("UAB", typeof(string));
                DataColumn colUbc = new DataColumn("UBC", typeof(string));
                DataColumn colUca = new DataColumn("UCA", typeof(string));
                DataColumn colIa = new DataColumn("IA", typeof(string));
                DataColumn colIb = new DataColumn("IB", typeof(string));
                DataColumn colIc = new DataColumn("IC", typeof(string));
                DataColumn colUa = new DataColumn("UA", typeof(string));
                DataColumn colUb = new DataColumn("UB", typeof(string));
                DataColumn colUc = new DataColumn("UC", typeof(string));
                DataColumn colPo = new DataColumn("PO", typeof(string));
                DataColumn colPp = new DataColumn("PP", typeof(string));
                DataColumn colPh = new DataColumn("PH", typeof(string));
                DataColumn colQo = new DataColumn("QO", typeof(string));
                DataColumn colQp = new DataColumn("QP", typeof(string));
                DataColumn colQh = new DataColumn("QH", typeof(string));
                DataColumn colSo = new DataColumn("SO", typeof(string));
                DataColumn colSp = new DataColumn("SP", typeof(string));
                DataColumn colSh = new DataColumn("SH", typeof(string));
                DataColumn colUo = new DataColumn("UO", typeof(string));
                DataColumn colUp = new DataColumn("UP", typeof(string));
                DataColumn colUh = new DataColumn("UH", typeof(string));
                DataColumn colIo = new DataColumn("IO", typeof(string));
                DataColumn colIp = new DataColumn("IP", typeof(string));
                DataColumn colIh = new DataColumn("IH", typeof(string));
                DataColumn colKo = new DataColumn("KO", typeof(string));
                DataColumn colKh = new DataColumn("KH", typeof(string));
                DataColumn colFreq = new DataColumn("Freq", typeof(string));
                DataColumn colSigmaUy = new DataColumn("SigmaUy", typeof(string));
                DataColumn colSigmaUyA = new DataColumn("SigmaUyA", typeof(string));
                DataColumn colSigmaUyB = new DataColumn("SigmaUyB", typeof(string));
                DataColumn colSigmaUyC = new DataColumn("SigmaUyC", typeof(string));

                dt.Columns.Add(colTimeTek);
                dt.Columns.Add(colUab);
                dt.Columns.Add(colUbc);
                dt.Columns.Add(colUca);
                dt.Columns.Add(colIa);
                dt.Columns.Add(colIb);
                dt.Columns.Add(colIc);
                dt.Columns.Add(colUa);
                dt.Columns.Add(colUb);
                dt.Columns.Add(colUc);
                dt.Columns.Add(colPo);
                dt.Columns.Add(colPp);
                dt.Columns.Add(colPh);
                dt.Columns.Add(colQo);
                dt.Columns.Add(colQp);
                dt.Columns.Add(colQh);
                dt.Columns.Add(colSo);
                dt.Columns.Add(colSp);
                dt.Columns.Add(colSh);
                dt.Columns.Add(colUo);
                dt.Columns.Add(colUp);
                dt.Columns.Add(colUh);
                dt.Columns.Add(colIo);
                dt.Columns.Add(colIp);
                dt.Columns.Add(colIh);
                dt.Columns.Add(colKo);
                dt.Columns.Add(colKh);
                dt.Columns.Add(colFreq);
                dt.Columns.Add(colSigmaUy);
                dt.Columns.Add(colSigmaUyA);
                dt.Columns.Add(colSigmaUyB);
                dt.Columns.Add(colSigmaUyC);
            }

            return dt;
        }


        public SD.DataTable GetParamTable
        {
            get
            {
                SD.DataTable dt = CreateParamTable();

                if (_listParamObject.Count > 0)
                {
                    foreach (ParamObject paramObject in _listParamObject)
                    {
                        DataRow row = dt.NewRow();

                        row["Имя объекта"] = paramObject.NameObject;
                        row["Время старка"] = paramObject.TimeStart.ToString("dd/MM/yyyy H:mm");
                        row["Время остановки"] = paramObject.TimeStop.ToString("dd/MM/yyyy H:mm");
                        row["Схема проверки"] = paramObject.ActiveCxema;

                        int seconds = paramObject.AveragingIntervalTime.Seconds;
                        if (seconds < 60) row["Интервал времени"] = string.Format("{0} c", seconds);
                        else
                        {
                            row["Интервал времени"] = string.Format("{0} минут", paramObject.AveragingIntervalTime.Minutes);
                        }
                        
                        dt.Rows.Add(row);
                    }
                }

                return dt;
            }
        }


        public SD.DataTable GetResultTable(int index)
        {
            SD.DataTable dt = CreateResultTable(_listParamObject[index].ActiveCxema);

            if (_listParamObject[index].ActiveCxema == 1)
            {
                foreach (ResultObjectSchema1 resultObject in _listParamObject[index].ListResultObject)
                {
                    DataRow row = dt.NewRow();

                    row["Дата / время"] = resultObject.TimeTek.ToString("dd/MM/yyyy H:mm");
                    row["UA"] = resultObject.Ua;
                    row["IA"] = resultObject.Ia;
                    row["PA"] = resultObject.Pa;
                    row["QA"] = resultObject.Qa;
                    row["SA"] = resultObject.Sa;
                    row["Freq"] = resultObject.Freq;
                    row["sigmaUy"] = resultObject.SigmaUy;

                    dt.Rows.Add(row);
                }
            }
            else if (_listParamObject[index].ActiveCxema == 2)
            {
                foreach (ResultObjectSchema2 resultObject in _listParamObject[index].ListResultObject)
                {
                    DataRow row = dt.NewRow();

                    row["Дата / время"] = resultObject.TimeTek;
                    row["UAB"] = resultObject.Uab;
                    row["UBC"] = resultObject.Ubc;
                    row["UCA"] = resultObject.Uca;
                    row["IAB"] = resultObject.Iab;
                    row["IBC"] = resultObject.Ubc;
                    row["ICA"] = resultObject.Ica;
                    row["IA"] = resultObject.Ia;
                    row["IB"] = resultObject.Ib;
                    row["IC"] = resultObject.Ic;
                    row["PO"] = resultObject.Po;
                    row["PP"] = resultObject.Pp;
                    row["QO"] = resultObject.Qo;
                    row["QP"] = resultObject.Qp;
                    row["SO"] = resultObject.So;
                    row["SP"] = resultObject.Sp;
                    row["UO"] = resultObject.Uo;
                    row["UP"] = resultObject.Up;
                    row["IO"] = resultObject.Io;
                    row["IP"] = resultObject.Ip;
                    row["KO"] = resultObject.Ko;
                    row["Freq"] = resultObject.Freq;
                    row["SigmaUy"] = resultObject.SigmaUy;
                    row["SigmaUyAB"] = resultObject.SigmaUyAb;
                    row["SigmaUyBC"] = resultObject.SigmaUyBc;
                    row["SigmaUyCA"] = resultObject.SigmaUyCa;

                    dt.Rows.Add(row);
                }
            }
            else if (_listParamObject[index].ActiveCxema == 3)
            {
                foreach (ResultObjectSchema3 resultObject in _listParamObject[index].ListResultObject)
                {
                    DataRow row = dt.NewRow();

                    row["Дата / время"] = resultObject.TimeTek;
                    row["UAB"] = resultObject.Uab;
                    row["UBC"] = resultObject.Ubc;
                    row["UCA"] = resultObject.Uca;
                    row["IA"] = resultObject.Ia;
                    row["IB"] = resultObject.Ib;
                    row["IC"] = resultObject.Ic;
                    row["UA"] = resultObject.Ua;
                    row["UB"] = resultObject.Ub;
                    row["UC"] = resultObject.Uc;
                    row["PO"] = resultObject.Po;
                    row["PP"] = resultObject.Pp;
                    row["PH"] = resultObject.Ph;
                    row["QO"] = resultObject.Qo;
                    row["QP"] = resultObject.Qp;
                    row["QH"] = resultObject.Qh;
                    row["SO"] = resultObject.So;
                    row["SP"] = resultObject.Sp;
                    row["SH"] = resultObject.Sh;
                    row["UO"] = resultObject.Uo;
                    row["UP"] = resultObject.Up;
                    row["UH"] = resultObject.Uh;
                    row["IO"] = resultObject.Io;
                    row["IP"] = resultObject.Ip;
                    row["IH"] = resultObject.Ih;
                    row["KO"] = resultObject.Ko;
                    row["KH"] = resultObject.Kh;
                    row["Freq"] = resultObject.Freq;
                    row["SigmaUy"] = resultObject.SigmaUy;
                    row["SigmaUyA"] = resultObject.SigmaUyA;
                    row["SigmaUyB"] = resultObject.SigmaUyB;
                    row["SigmaUyC"] = resultObject.SigmaUyC;

                    dt.Rows.Add(row);
                }
            }

            return dt;
        }


        public Pke(string pkePath)
        {
            _listParamObject = new List<ParamObject>();
            MessageException = string.Empty;

            List<string> fileNames = GetFilesFromDirection(pkePath);

            foreach (string fileName in fileNames)
            {
                ParamObject paramObject;

                try
                {
                    paramObject = new ParamObject(fileName);
                }
                catch(Exception e)
                {
                    MessageException += $"{e.Message}\n";

                    continue;
                }

                int i = 0;
                for (i = 0; i < _listParamObject.Count; i++)
                {
                    ParamObject pObject = _listParamObject[i];

                    if (pObject.Uid != paramObject.Uid) continue;
                    pObject.ListResultObject.Add(paramObject.ListResultObject[0]);

                    break;
                }

                if (i == _listParamObject.Count) _listParamObject.Add(paramObject);

                // Сортировка.
                foreach(ParamObject pObject in _listParamObject)
                {
                    pObject.SortByTimeTek();
                }
            }
        }


        private List<string> GetFilesFromDirection(string path)
        {
            List<string> allFiels = new List<string>();

            string[] directories = Directory.GetDirectories(path);

            if (directories.Length > 0)
            {
                foreach (string directory in directories)
                {
                    allFiels.AddRange(GetFilesFromDirection(directory));
                }
            }

            string[] fiels = Directory.GetFiles(path);
            allFiels.AddRange(fiels);

            return allFiels;
        }


        private void WriteToExcel(Worksheet workSheet, SD.DataTable dt, ref int indexDoc)
        {
            const int indent = 5;

            // Заголовок.
            for (int j = 0; j < dt.Columns.Count; j++)
            {
                workSheet.Cells[indexDoc + 1, j + 1] = dt.Columns[j].ColumnName;
            }

            // Содержимое.
            int i;
            for (i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    workSheet.Cells[indexDoc + i + 2, j + 1] = dt.Rows[i].ItemArray[j];
                }
            }

            indexDoc = indexDoc + i + indent;
        }


        public void ExportToExcel(string excelFile)
        {
            Excel.Application excelApp = new Excel.Application();
            excelApp.Workbooks.Add();

            Worksheet workSheet = (Worksheet) excelApp.ActiveSheet;
            workSheet.Name = "ParamObjects";

            SD.DataTable dt = this.GetParamTable;

            int indexDoc = 0;
            
            // ParamObjects.
            WriteToExcel(workSheet, dt, ref indexDoc);

            // ResultObjects.
            for (int k = 0; k < _listParamObject.Count; k++)
            {
                dt = this.GetResultTable(k);

                WriteToExcel(workSheet, dt, ref indexDoc);
            }

            workSheet.SaveAs(excelFile);

            excelApp.Quit();
        }

    }
}
