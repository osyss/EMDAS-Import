using OfficeOpenXml;
using System;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;
using System.Xml;

namespace JYSK_EMDAS_XML //version 1.2
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Converting Excel....");
            string currentDirectory = Directory.GetCurrentDirectory();
            ExcelWorksheet worksheet = new ExcelPackage(new FileInfo(currentDirectory + "\\IM.xlsx")).Workbook.Worksheets["IM sagatave"];
            CultureInfo cultureInfo = (CultureInfo)Thread.CurrentThread.CurrentCulture.Clone();
            cultureInfo.NumberFormat.NumberDecimalSeparator = ".";
            Thread.CurrentThread.CurrentCulture = cultureInfo;
            int index1 = 89;
            Random random = new Random();
            random.Next(10);
            random.Next(10);
            random.Next(10);
            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Encoding = (Encoding)new UTF8Encoding(false);
            MemoryStream memoryStream = new MemoryStream();
            XmlWriter xmlWriter = XmlWriter.Create((Stream)memoryStream, settings);
            xmlWriter.WriteStartElement("IcsIE");
            xmlWriter.WriteElementString("MesTypMES20", "I01");
            xmlWriter.WriteStartElement("HEAHEA");
            xmlWriter.WriteElementString("RefNumEPT1", (worksheet.Cells["B36"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("CusOffENT730", (worksheet.Cells["B37"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("DecWorDatHEA999", DateTime.Now.ToString("yyyy-MM-ddT00:00:00.000"));
            xmlWriter.WriteElementString("DecPlaHEA394", (worksheet.Cells["G12"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("TypOfOpeHEA994", (worksheet.Cells["B4"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("OpeCodHEA995", (worksheet.Cells["C4"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("TotGroMasHEA307", (worksheet.Cells["B68"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("CouOfDisCodHEA55", (worksheet.Cells["B32"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("TypOfPriGHEA1003", (worksheet.Cells["B8"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("IdeOfMeaOfTraAtDHEA78", (worksheet.Cells["B50"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("DelReqCodHEA992", (worksheet.Cells["B46"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("DesOfDelReqHEA993", (worksheet.Cells["B47"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("NatOfMeaOfTraCroHEA87", (worksheet.Cells["B52"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("TraModAtBorHEA76", (worksheet.Cells["B54"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("InlTraModHEA75", (worksheet.Cells["B53"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("TypOfDeaHEA995", (worksheet.Cells["B66"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("Cur126", (worksheet.Cells["C72"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("DefPayVGA48", (worksheet.Cells["E66"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("CusProCodGDS379", (worksheet.Cells["C1"].Value ?? (object)"").ToString());
            xmlWriter.WriteEndElement();
            xmlWriter.WriteStartElement("TRACONCO1");
            xmlWriter.WriteElementString("NamCO17", (worksheet.Cells["B18"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("StrAndNumCO122", (worksheet.Cells["B19"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("PosCodCO123", (worksheet.Cells["B21"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("CitCO124", (worksheet.Cells["B20"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("CouCO125", (worksheet.Cells["B22"].Value ?? (object)"").ToString());
            xmlWriter.WriteEndElement();
            xmlWriter.WriteStartElement("TRACONCE1");
            xmlWriter.WriteElementString("NamCE17", (worksheet.Cells["B25"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("StrAndNumCE122", (worksheet.Cells["B26"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("PosCodCE123", (worksheet.Cells["B28"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("CitCE124", (worksheet.Cells["B27"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("CouCE125", (worksheet.Cells["B29"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("TINCE159", (worksheet.Cells["B24"].Value ?? (object)"").ToString());
            xmlWriter.WriteEndElement();
            xmlWriter.WriteStartElement("TRAREP");
            xmlWriter.WriteElementString("NamTRE1", (worksheet.Cells["B11"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("StrAndNumTRE1", (worksheet.Cells["B12"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("PosCodTRE1", (worksheet.Cells["B14"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("CitTRE1", (worksheet.Cells["B13"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("CouCodTRE1", (worksheet.Cells["B15"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("TINTRE1", (worksheet.Cells["B10"].Value ?? (object)"").ToString());
            xmlWriter.WriteEndElement();
            xmlWriter.WriteStartElement("NOTPAR670");
            xmlWriter.WriteElementString("NamNOTPAR672", (worksheet.Cells["G10"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("StrNumNOTPAR673", (worksheet.Cells["G11"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("PosCodNOTPAR676", (worksheet.Cells["G13"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("CitNOTPAR674", (worksheet.Cells["G12"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("CouCodNOTPAR675", (worksheet.Cells["G14"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("TINNOTPAR671", (worksheet.Cells["G9"].Value ?? (object)"").ToString());
            xmlWriter.WriteEndElement();
            xmlWriter.WriteStartElement("GDSLOC");
            xmlWriter.WriteStartElement("GDSLOCStr");
            xmlWriter.WriteElementString("CitGdsLoc", (worksheet.Cells["B57"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("StrGdsLoc", (worksheet.Cells["B58"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("StrNumLoc", (worksheet.Cells["F56"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("PosCodGdsLoc", (worksheet.Cells["F57"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("CouCodGdsLoc", (worksheet.Cells["F58"].Value ?? (object)"").ToString());
            xmlWriter.WriteEndElement();
            xmlWriter.WriteEndElement();
            xmlWriter.WriteStartElement("WAR");
            xmlWriter.WriteElementString("WarTypWARTyp", (worksheet.Cells["B61"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("WarTypWARId", (worksheet.Cells["B62"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("WarTypWARCou", (worksheet.Cells["B63"].Value ?? (object)"").ToString());
            xmlWriter.WriteEndElement();
            for (int index2 = 89; index2 > 50 && worksheet.Cells[index2, 1].Value != null; ++index2)
            {
                xmlWriter.WriteStartElement("GOOITEGDS");
                xmlWriter.WriteElementString("GooDesGDS23", (worksheet.Cells[index1, 20].Value ?? (object)"").ToString());
                xmlWriter.WriteElementString("IteNumGDS7", (worksheet.Cells[index1, 1].Value ?? (object)"").ToString());
                xmlWriter.WriteElementString("IteNumB32F1", (worksheet.Cells[index1, 1].Value ?? (object)"").ToString());
                xmlWriter.WriteElementString("ComCodTarCodGDS10", (worksheet.Cells[index1, 6].Value ?? (object)"").ToString());
                xmlWriter.WriteElementString("CouOfOriCodGDS63", (worksheet.Cells[index1, 8].Value ?? (object)"").ToString());
                xmlWriter.WriteElementString("GroMasGDS46", (worksheet.Cells[index1, 4].Value ?? (object)"").ToString());
                xmlWriter.WriteElementString("NetMasGDS48", (worksheet.Cells[index1, 3].Value ?? (object)"").ToString());
                xmlWriter.WriteElementString("CusProCodGDS379", (worksheet.Cells[index1, 24].Value ?? (object)"").ToString() + (worksheet.Cells[index1, 25].Value ?? (object)"").ToString());
                xmlWriter.WriteElementString("AddCusProCodGDS340", (worksheet.Cells[index1, 21].Value ?? (object)"").ToString());
                xmlWriter.WriteElementString("GooInvB42", (worksheet.Cells[index1, 5].Value ?? (object)"").ToString());
                xmlWriter.WriteElementString("AppRecMet", (worksheet.Cells[index1, 28].Value ?? (object)"").ToString());
                xmlWriter.WriteElementString("PreB36", (worksheet.Cells[index1, 27].Value ?? (object)"").ToString());
                if (worksheet.Cells[index1, 11].Value != null)
                {
                    xmlWriter.WriteStartElement("OutGooFreAmSTA");
                    xmlWriter.WriteElementString("AmSTAElm", (worksheet.Cells[index1, 11].Value ?? (object)"").ToString());
                    xmlWriter.WriteElementString("AmSTAElmNat", (worksheet.Cells[index1, 13].Value ?? (object)"").ToString());
                    xmlWriter.WriteEndElement();
                }
                xmlWriter.WriteStartElement("PACGS2");
                xmlWriter.WriteElementString("KinOfPacGS23", (worksheet.Cells[index1, 30].Value ?? (object)"").ToString());
                xmlWriter.WriteElementString("MarNumOfPacGS21", (worksheet.Cells[index1, 31].Value ?? (object)"").ToString());
                xmlWriter.WriteElementString("NumOfPacGS24", (worksheet.Cells[index1, 2].Value ?? (object)"").ToString());
                xmlWriter.WriteEndElement();
                if (worksheet.Cells[index1, 29].Value != null)
                {
                    xmlWriter.WriteStartElement("CONNR2");
                    xmlWriter.WriteElementString("ConNumNR21", (worksheet.Cells[index1, 29].Value ?? (object)"").ToString());
                    xmlWriter.WriteEndElement();
                }
                xmlWriter.WriteStartElement("PREADMREFAR2");
                xmlWriter.WriteElementString("PreDocTyp", (worksheet.Cells[index1, 22].Value ?? (object)"").ToString());
                xmlWriter.WriteElementString("TitAR212", (worksheet.Cells[index1, 23].Value ?? (object)"").ToString());
                xmlWriter.WriteElementString("PreDocRefAR26", (worksheet.Cells[index1, 7].Value ?? (object)"").ToString());
                xmlWriter.WriteEndElement();
                int index3 = 34;
                while (index3 > 10 && worksheet.Cells[index1, index3].Value != null)
                {
                    xmlWriter.WriteStartElement("PRODOCDC2");
                    xmlWriter.WriteElementString("DocCoDC28", (worksheet.Cells[index1, index3].Value ?? (object)"").ToString());
                    xmlWriter.WriteElementString("TitDC29", (worksheet.Cells[index1, index3 + 1].Value ?? (object)"").ToString());
                    xmlWriter.WriteElementString("ComOfInfDC25", (worksheet.Cells[index1, index3 + 2].Value ?? (object)"").ToString());
                    xmlWriter.WriteEndElement();
                    index3 += 3;
                }
                xmlWriter.WriteEndElement();
                ++index1;
            }
            xmlWriter.WriteStartElement("STA");
            xmlWriter.WriteStartElement("InvAmSTA");
            xmlWriter.WriteElementString("CurSTAElm", (worksheet.Cells["C72"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("AmSTAElm", (worksheet.Cells["B72"].Value ?? (object)"").ToString());
            xmlWriter.WriteEndElement();
            xmlWriter.WriteStartElement("OutFreAmSTA");
            xmlWriter.WriteElementString("CurSTAElm", (worksheet.Cells["C73"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("AmSTAElm", (worksheet.Cells["B73"].Value ?? (object)"").ToString());
            if ((worksheet.Cells["G73"].Value ?? (object)"").ToString() == "Manuāli")
                xmlWriter.WriteElementString("DivMetAmSTA", "1");
            else if ((worksheet.Cells["G73"].Value ?? (object)"").ToString() == "Pēc svara")
                xmlWriter.WriteElementString("DivMetAmSTA", "2");
            else if ((worksheet.Cells["G73"].Value ?? (object)"").ToString() == "Pēc vērtības")
                xmlWriter.WriteElementString("DivMetAmSTA", "3");
            xmlWriter.WriteEndElement();
            xmlWriter.WriteStartElement("InsAmSTA");
            xmlWriter.WriteElementString("CurSTAElm", (worksheet.Cells["C74"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("AmSTAElm", (worksheet.Cells["B74"].Value ?? (object)"").ToString());
            if ((worksheet.Cells["G74"].Value ?? (object)"").ToString() == "Manuāli")
                xmlWriter.WriteElementString("DivMetAmSTA", "1");
            else if ((worksheet.Cells["G74"].Value ?? (object)"").ToString() == "Pēc svara")
                xmlWriter.WriteElementString("DivMetAmSTA", "2");
            else if ((worksheet.Cells["G74"].Value ?? (object)"").ToString() == "Pēc vērtības")
                xmlWriter.WriteElementString("DivMetAmSTA", "3");
            xmlWriter.WriteEndElement();
            xmlWriter.WriteStartElement("OthCstAmSTA");
            xmlWriter.WriteElementString("CurSTAElm", (worksheet.Cells["C75"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("AmSTAElm", (worksheet.Cells["B75"].Value ?? (object)"").ToString());
            if ((worksheet.Cells["G75"].Value ?? (object)"").ToString() == "Manuāli")
                xmlWriter.WriteElementString("DivMetAmSTA", "1");
            else if ((worksheet.Cells["G75"].Value ?? (object)"").ToString() == "Pēc svara")
                xmlWriter.WriteElementString("DivMetAmSTA", "2");
            else if ((worksheet.Cells["G75"].Value ?? (object)"").ToString() == "Pēc vērtības")
                xmlWriter.WriteElementString("DivMetAmSTA", "3");
            xmlWriter.WriteEndElement();
            xmlWriter.WriteStartElement("InnFreAmSTA");
            xmlWriter.WriteElementString("CurSTAElm", (worksheet.Cells["C76"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("AmSTAElm", (worksheet.Cells["B76"].Value ?? (object)"").ToString());
            if ((worksheet.Cells["G76"].Value ?? (object)"").ToString() == "Manuāli")
                xmlWriter.WriteElementString("DivMetAmSTA", "1");
            else if ((worksheet.Cells["G76"].Value ?? (object)"").ToString() == "Pēc svara")
                xmlWriter.WriteElementString("DivMetAmSTA", "2");
            else if ((worksheet.Cells["G76"].Value ?? (object)"").ToString() == "Pēc vērtības")
                xmlWriter.WriteElementString("DivMetAmSTA", "3");
            xmlWriter.WriteEndElement();
            xmlWriter.WriteStartElement("DedAmSTA");
            xmlWriter.WriteElementString("CurSTAElm", (worksheet.Cells["C77"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("AmSTAElm", (worksheet.Cells["B77"].Value ?? (object)"").ToString());
            if ((worksheet.Cells["G77"].Value ?? (object)"").ToString() == "Manuāli")
                xmlWriter.WriteElementString("DivMetAmSTA", "1");
            else if ((worksheet.Cells["G77"].Value ?? (object)"").ToString() == "Pēc svara")
                xmlWriter.WriteElementString("DivMetAmSTA", "2");
            else if ((worksheet.Cells["G77"].Value ?? (object)"").ToString() == "Pēc vērtības")
                xmlWriter.WriteElementString("DivMetAmSTA", "3");
            xmlWriter.WriteEndElement();
            xmlWriter.WriteStartElement("InlTAXAmSTA");
            xmlWriter.WriteElementString("CurSTAElm", (worksheet.Cells["C79"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("AmSTAElm", (worksheet.Cells["B79"].Value ?? (object)"").ToString());
            if ((worksheet.Cells["G79"].Value ?? (object)"").ToString() == "Manuāli")
                xmlWriter.WriteElementString("DivMetAmSTA", "1");
            else if ((worksheet.Cells["G79"].Value ?? (object)"").ToString() == "Nav jāsadala")
                xmlWriter.WriteElementString("DivMetAmSTA", "4");
            else if ((worksheet.Cells["G79"].Value ?? (object)"").ToString() == "Pēc svara")
                xmlWriter.WriteElementString("DivMetAmSTA", "2");
            else if ((worksheet.Cells["G79"].Value ?? (object)"").ToString() == "Pēc vērtības")
                xmlWriter.WriteElementString("DivMetAmSTA", "3");
            xmlWriter.WriteEndElement();
            xmlWriter.WriteEndElement();
            xmlWriter.Flush();
            xmlWriter.Close();

            File.WriteAllText(currentDirectory + "\\XML_IM_" + DateTime.Now.ToString("yyyyMMddHHmm") + ".xml", new UTF8Encoding(false).GetString(memoryStream.ToArray()));
        }
    }
}
