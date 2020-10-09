using System;
using System.Collections.Generic;
using System.Diagnostics.Eventing.Reader;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.IO;


namespace test1
{
    internal abstract class ResultObject
    {
        public abstract int PkeCxema { get; set; }
        public abstract DateTime TimeTek { get; set; }
    }

    sealed class ResultObjectSchema1 : ResultObject
    {
        public override int PkeCxema { get; set; }
        public override DateTime TimeTek { get; set; }
        public string Ua;
        public string Ia;
        public string Pa;
        public string Qa;
        public string Sa;
        public string Freq;
        public string SigmaUy;

        public ResultObjectSchema1(int pkeCxema, DateTime timeTek, string ua, string ia, string pa, string qa, string sa, string freq, string sigmaUy)
        {
            PkeCxema = pkeCxema;
            TimeTek = timeTek;
            Ua = ua;
            Ia = ia;
            Pa = pa;
            Qa = qa;
            Sa = sa;
            Freq = freq;
            SigmaUy = sigmaUy;
        }
    }

    sealed class ResultObjectSchema2 : ResultObject
    {
        public override int PkeCxema { get; set; }
        public override DateTime TimeTek { get; set; }
        public string Uab;
        public string Ubc;
        public string Uca;
        public string Iab;
        public string Ibc;
        public string Ica;
        public string Ia;
        public string Ib;
        public string Ic;
        public string Po;
        public string Pp;
        public string Qo;
        public string Qp;
        public string So;
        public string Sp;
        public string Uo;
        public string Up;
        public string Io;
        public string Ip;
        public string Ko;
        public string Freq;
        public string SigmaUy;
        public string SigmaUyAb;
        public string SigmaUyBc;
        public string SigmaUyCa;

        public ResultObjectSchema2(int pkeCxema, DateTime timeTek, string uab, string ubc, string uca, string iab, string ibc, string ica, string ia, string ib, string ic, string po, string pp, string qo, string qp, string so, string sp, string uo, string up, string io, string ip, string ko, string freq, string sigmaUy, string sigmaUyAb, string sigmaUyBc, string sigmaUyCa)
        {
            PkeCxema = pkeCxema;
            TimeTek = timeTek;
            Uab = uab;
            Ubc = ubc;
            Uca = uca;
            Iab = iab;
            Ibc = ibc;
            Ica = ica;
            Ia = ia;
            Ib = ib;
            Ic = ic;
            Po = po;
            Pp = pp;
            Qo = qo;
            Qp = qp;
            So = so;
            Sp = sp;
            Uo = uo;
            Up = up;
            Io = io;
            Ip = ip;
            Ko = ko;
            Freq = freq;
            SigmaUy = sigmaUy;
            SigmaUyAb = sigmaUyAb;
            SigmaUyBc = sigmaUyBc;
            SigmaUyCa = sigmaUyCa;
        }
    }

    sealed class ResultObjectSchema3 : ResultObject
    {
        public override int PkeCxema { get; set; }
        public override DateTime TimeTek { get; set; }
        public string Uab;
        public string Ubc;
        public string Uca;
        public string Ia;
        public string Ib;
        public string Ic;
        public string Ua;
        public string Ub;
        public string Uc;
        public string Po;
        public string Pp;
        public string Ph;
        public string Qo;
        public string Qp;
        public string Qh;
        public string So;
        public string Sp;
        public string Sh;
        public string Uo;
        public string Up;
        public string Uh;
        public string Io;
        public string Ip;
        public string Ih;
        public string Ko;
        public string Kh;
        public string Freq;
        public string SigmaUy;
        public string SigmaUyA;
        public string SigmaUyB;
        public string SigmaUyC;

        public ResultObjectSchema3(int pkeCxema, DateTime timeTek, string uab, string ubc, string uca, string ia, string ib, string ic, string ua, string ub, string uc, string po, string pp, string ph, string qo, string qp, string qh, string so, string SP, string sh, string uo, string up, string uh, string io, string ip, string ih, string ko, string kh, string freq, string sigmaUy, string sigmaUyA, string sigmaUyB, string sigmaUyC)
        {
            PkeCxema = pkeCxema;
            TimeTek = timeTek;
            Uab = uab;
            Ubc = ubc;
            Uca = uca;
            Ia = ia;
            Ib = ib;
            Ic = ic;
            Ua = ua;
            Ub = ub;
            Uc = uc;
            Po = po;
            Pp = pp;
            Ph = ph;
            Qo = qo;
            Qp = qp;
            Qh = qh;
            So = so;
            Sp = SP;
            Sh = sh;
            Uo = uo;
            Up = up;
            Uh = uh;
            Io = io;
            Ip = ip;
            Ih = ih;
            Ko = ko;
            Kh = kh;
            Freq = freq;
            SigmaUy = sigmaUy;
            SigmaUyA = sigmaUyA;
            SigmaUyB = sigmaUyB;
            SigmaUyC = sigmaUyC;
        }
    }


    class ParamObject
    {
        public long Uid;
        public DateTime TimeStart { get; set; }
        public DateTime TimeStop { get; set; }
        public string NameObject { get; set; }
        public TimeSpan AveragingIntervalTime { get; set; }
        public int ActiveCxema { get; set; }

        public List<ResultObject> ListResultObject;


        public ParamObject(string fileName)
        {
            XDocument xDoc = XDocument.Load(fileName);

            // RM3.
            if (xDoc.Descendants("RM3_ПКЭ").Any())
            {
                foreach (XElement xelement in xDoc.Descendants("RM3_ПКЭ"))
                {
                    if (xelement.HasAttributes && xelement.Attribute("UID") != null)
                    {
                        string sUid = xelement.Attribute("UID")?.Value;
                        if (sUid != null) Uid = Convert.ToInt64(sUid.Substring(1, sUid.IndexOf('-', 0, 10) - 1), 16);
                        else throw new Exception($"Файл {Path.GetFileName(fileName)} имеет неверные параметры RM3_ПКЭ.");
                    }
                }
            }
            else
            {
                throw new Exception($"Файл {Path.GetFileName(fileName)} не имеет элемента RM3_ПКЭ.");
            }

            // Param_Check_PKE.
            if (xDoc.Descendants("Param_Check_PKE").Any())
            {
                foreach (XElement xelement in xDoc.Descendants("Param_Check_PKE"))
                {
                    try
                    {
                        TimeStart = new DateTime(1970, 1, 1) + TimeSpan.FromMilliseconds(double.Parse(xelement.Attribute("TimeStart")?.Value ?? throw new InvalidOperationException()));
                        TimeStop = new DateTime(1970, 1, 1) + TimeSpan.FromMilliseconds(double.Parse(xelement.Attribute("TimeStop")?.Value ?? throw new InvalidOperationException()));
                        NameObject = xelement.Attribute("nameObject")?.Value;
                        AveragingIntervalTime = TimeSpan.FromMilliseconds(double.Parse(xelement.Attribute("averaging_interval_time")?.Value ?? throw new InvalidOperationException()));
                        ActiveCxema = int.Parse(xelement.Attribute("active_cxema")?.Value ?? throw new InvalidOperationException());
                    }
                    catch(Exception)
                    {
                        throw new Exception(
                            $"Файл {Path.GetFileName(fileName)} имеет неверные параметры Param_Check_PKE.");
                    }
                }
            }
            else
            {
                throw new Exception($"Файл {Path.GetFileName(fileName)} не имеет элемента Param_Check_PKE.");
            }

            // Result_Check_PKE.
            ListResultObject = new List<ResultObject>();

            if (xDoc.Descendants("Result_Check_PKE").Any())
            {
                foreach (XElement xelement in xDoc.Descendants("Result_Check_PKE"))
                {
                    try
                    {
                        if (ActiveCxema == 1)
                        {
                            int pkeCxema = int.Parse(xelement.Attribute("pke_cxema")?.Value ?? throw new InvalidOperationException());
                            DateTime timeTek = new DateTime(1970, 1, 1) + TimeSpan.FromMilliseconds(double.Parse(xelement.Attribute("TimeTek")?.Value ?? throw new InvalidOperationException()));
                            string ua = xelement.Attribute("UA")?.Value ?? throw new InvalidOperationException();
                            string ia = xelement.Attribute("IA")?.Value ?? throw new InvalidOperationException();
                            string pa = xelement.Attribute("PA")?.Value ?? throw new InvalidOperationException();
                            string qa = xelement.Attribute("QA")?.Value ?? throw new InvalidOperationException();
                            string sa = xelement.Attribute("SA")?.Value ?? throw new InvalidOperationException();
                            string freq = xelement.Attribute("Freq")?.Value ?? throw new InvalidOperationException();
                            string sigmaUy = xelement.Attribute("sigmaUy")?.Value ?? throw new InvalidOperationException();

                            ListResultObject.Add(new ResultObjectSchema1(pkeCxema, timeTek, ua, ia, pa, qa, sa, freq, sigmaUy));
                        }
                        else if (ActiveCxema == 2)
                        {
                            int pkeCxema = int.Parse(xelement.Attribute("pke_cxema")?.Value ?? throw new InvalidOperationException());
                            DateTime timeTek = new DateTime(1970, 1, 1) + TimeSpan.FromMilliseconds(double.Parse(xelement.Attribute("TimeTek")?.Value ?? throw new InvalidOperationException()));
                            string uab = xelement.Attribute("UAB")?.Value ?? throw new InvalidOperationException();
                            string ubc = xelement.Attribute("UBC")?.Value ?? throw new InvalidOperationException();
                            string uca = xelement.Attribute("UCA")?.Value ?? throw new InvalidOperationException();
                            string iab = xelement.Attribute("IAB")?.Value ?? throw new InvalidOperationException();
                            string ibc = xelement.Attribute("IBC")?.Value ?? throw new InvalidOperationException();
                            string ica = xelement.Attribute("ICA")?.Value ?? throw new InvalidOperationException();
                            string ia = xelement.Attribute("IA")?.Value ?? throw new InvalidOperationException();
                            string ib = xelement.Attribute("IB")?.Value ?? throw new InvalidOperationException();
                            string ic = xelement.Attribute("IC")?.Value ?? throw new InvalidOperationException();
                            string po = xelement.Attribute("PO")?.Value ?? throw new InvalidOperationException();
                            string pp = xelement.Attribute("PP")?.Value ?? throw new InvalidOperationException();
                            string qo = xelement.Attribute("QO")?.Value ?? throw new InvalidOperationException();
                            string qp = xelement.Attribute("QP")?.Value ?? throw new InvalidOperationException();
                            string so = xelement.Attribute("SO")?.Value ?? throw new InvalidOperationException();
                            string sp = xelement.Attribute("SP")?.Value ?? throw new InvalidOperationException();
                            string uo = xelement.Attribute("UO")?.Value ?? throw new InvalidOperationException();
                            string up = xelement.Attribute("UP")?.Value ?? throw new InvalidOperationException();
                            string io = xelement.Attribute("IO")?.Value ?? throw new InvalidOperationException();
                            string ip = xelement.Attribute("IP")?.Value ?? throw new InvalidOperationException();
                            string ko = xelement.Attribute("KO")?.Value ?? throw new InvalidOperationException();
                            string freq = xelement.Attribute("Freq")?.Value ?? throw new InvalidOperationException();
                            string sigmaUy = xelement.Attribute("sigmaUy")?.Value ?? throw new InvalidOperationException();
                            string sigmaUyAb = xelement.Attribute("sigmaUyAB")?.Value ?? throw new InvalidOperationException();
                            string sigmaUyBc = xelement.Attribute("sigmaUyBC")?.Value ?? throw new InvalidOperationException();
                            string sigmaUyCa = xelement.Attribute("sigmaUyCA")?.Value ?? throw new InvalidOperationException();

                            ListResultObject.Add(new ResultObjectSchema2(pkeCxema, timeTek, uab, ubc, uca, iab, ibc, ica, ia, ib, ic, po, pp, qo, qp, so, sp, uo, up, io, ip, ko, freq, sigmaUy, sigmaUyAb, sigmaUyBc, sigmaUyCa));
                        }
                        else if (ActiveCxema == 3)
                        {
                            int pkeCxema = int.Parse(xelement.Attribute("pke_cxema")?.Value ?? throw new InvalidOperationException());
                            DateTime timeTek = new DateTime(1970, 1, 1) + TimeSpan.FromMilliseconds(double.Parse(xelement.Attribute("TimeTek")?.Value ?? throw new InvalidOperationException()));
                            string uab = xelement.Attribute("UAB")?.Value ?? throw new InvalidOperationException();
                            string ubc = xelement.Attribute("UBC")?.Value ?? throw new InvalidOperationException();
                            string uca = xelement.Attribute("UCA")?.Value ?? throw new InvalidOperationException();
                            string ia = xelement.Attribute("IA")?.Value ?? throw new InvalidOperationException();
                            string ib = xelement.Attribute("IB")?.Value ?? throw new InvalidOperationException();
                            string ic = xelement.Attribute("IC")?.Value ?? throw new InvalidOperationException();
                            string ua = xelement.Attribute("UA")?.Value ?? throw new InvalidOperationException();
                            string ub = xelement.Attribute("UB")?.Value ?? throw new InvalidOperationException();
                            string uc = xelement.Attribute("UC")?.Value ?? throw new InvalidOperationException();
                            string po = xelement.Attribute("PO")?.Value ?? throw new InvalidOperationException();
                            string pp = xelement.Attribute("PP")?.Value ?? throw new InvalidOperationException();
                            string ph = xelement.Attribute("PH")?.Value ?? throw new InvalidOperationException();
                            string qo = xelement.Attribute("QO")?.Value ?? throw new InvalidOperationException();
                            string qp = xelement.Attribute("QP")?.Value ?? throw new InvalidOperationException();
                            string qh = xelement.Attribute("QH")?.Value ?? throw new InvalidOperationException();
                            string so = xelement.Attribute("SO")?.Value ?? throw new InvalidOperationException();
                            string sp = xelement.Attribute("SP")?.Value ?? throw new InvalidOperationException();
                            string sh = xelement.Attribute("SH")?.Value ?? throw new InvalidOperationException();
                            string uo = xelement.Attribute("UO")?.Value ?? throw new InvalidOperationException();
                            string up = xelement.Attribute("UP")?.Value ?? throw new InvalidOperationException();
                            string uh = xelement.Attribute("UH")?.Value ?? throw new InvalidOperationException();
                            string io = xelement.Attribute("IO")?.Value ?? throw new InvalidOperationException();
                            string ip = xelement.Attribute("IP")?.Value ?? throw new InvalidOperationException();
                            string ih = xelement.Attribute("IH")?.Value ?? throw new InvalidOperationException();
                            string ko = xelement.Attribute("KO")?.Value ?? throw new InvalidOperationException();
                            string kh = xelement.Attribute("KH")?.Value ?? throw new InvalidOperationException();
                            string freq = xelement.Attribute("Freq")?.Value ?? throw new InvalidOperationException();
                            string sigmaUy = xelement.Attribute("sigmaUy")?.Value ?? throw new InvalidOperationException();
                            string sigmaUyA = xelement.Attribute("sigmaUyA")?.Value ?? throw new InvalidOperationException();
                            string sigmaUyB = xelement.Attribute("sigmaUyB")?.Value ?? throw new InvalidOperationException();
                            string sigmaUyC = xelement.Attribute("sigmaUyC")?.Value ?? throw new InvalidOperationException();

                            ListResultObject.Add(new ResultObjectSchema3(pkeCxema, timeTek, uab, ubc, uca, ia, ib, ic, ua, ub, uc, po, pp, ph, qo, qp, qh, so, sp, sh, uo, up, uh, io, ip, ih, ko, kh, freq, sigmaUy, sigmaUyA, sigmaUyB, sigmaUyC));
                        }
                    }
                    catch (Exception)
                    {
                        throw new Exception(
                            $"Файл {Path.GetFileName(fileName)} имеет неверные параметры Result_Check_PKE.");
                    }
                } // foreach (XElement xelement in xDoc.Descendants("Result_Check_PKE "))
            } // if (xDoc.Descendants("Result_Check_PKE").Any())
            else
            {
                throw new Exception($"Файл {Path.GetFileName(fileName)} не имеет элемента Result_Check_PKE.");
            }
        } // public ParamObject(string fileName)


        public void SortByTimeTek()
        {
            for (int i = 0; i < ListResultObject.Count - 1; i++)
            {
                bool isSwapped = false;

                for (int j = 0; j < ListResultObject.Count - 1; j++)
                {
                    if (ListResultObject[j + 1].TimeTek > ListResultObject[j].TimeTek)
                    {
                        ResultObject tmp = ListResultObject[j];
                        ListResultObject[j] = ListResultObject[j + 1];
                        ListResultObject[j + 1] = tmp;

                        isSwapped = true;
                    }
                }

                if (!isSwapped) break;
            }
        }

    }
}
