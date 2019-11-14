using OfficeOpenXml;
using System;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;
using System.Xml;

namespace JYSK_EMDAS_XML //version 2.8
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Converting Excel....");
            string currentDirectory = Directory.GetCurrentDirectory();
            ExcelWorksheet worksheet = new ExcelPackage(new FileInfo(currentDirectory + "\\IM.xlsx")).Workbook.Worksheets["IM sagatave NEW"];
            CultureInfo cultureInfo = (CultureInfo)Thread.CurrentThread.CurrentCulture.Clone();
            cultureInfo.NumberFormat.NumberDecimalSeparator = ".";
            Thread.CurrentThread.CurrentCulture = cultureInfo;
            int index1 = 104;
            Random random = new Random();
            random.Next(10);
            random.Next(10);
            random.Next(10);
            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Indent = true;
            settings.Encoding = (Encoding)new UTF8Encoding(false);
            settings.OmitXmlDeclaration = true;
            settings.NewLineOnAttributes = true;
            MemoryStream memoryStream = new MemoryStream();
            XmlWriter xmlWriter = XmlWriter.Create((Stream)memoryStream, settings);
            xmlWriter.WriteRaw("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n");
            xmlWriter.WriteStartElement("IcsIE");

            if ((worksheet.Cells["C6"].Value ?? (object)"").ToString() == "A")
                xmlWriter.WriteElementString("MesTypMES20", "I01.1");
            else if ((worksheet.Cells["C6"].Value ?? (object)"").ToString() == "C")
                xmlWriter.WriteElementString("MesTypMES20", "I01.2");
            else if ((worksheet.Cells["C6"].Value ?? (object)"").ToString() == "EIDR")
                xmlWriter.WriteElementString("MesTypMES20", "I01.3");
            else if ((worksheet.Cells["C6"].Value ?? (object)"").ToString() == "Y, Z")
                xmlWriter.WriteElementString("MesTypMES20", "I01.4");

            xmlWriter.WriteStartElement("HEAHEA");

            xmlWriter.WriteElementString("RefNumEPT1", (worksheet.Cells["B7"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("TypOfDecHEA24", (worksheet.Cells["B6"].Value ?? (object)"").ToString() + (worksheet.Cells["C6"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("CusProCodGDS379", (worksheet.Cells["B5"].Value ?? (object)"").ToString());
            if ((worksheet.Cells["B10"].Value ?? (object)"").ToString() == "Nav pārstāvības")
                xmlWriter.WriteElementString("TypOfPriGHEA1003", "1");
            else if ((worksheet.Cells["B10"].Value ?? (object)"").ToString() == "Tieša pārstāvība")
                xmlWriter.WriteElementString("TypOfPriGHEA1003", "2");
            else if ((worksheet.Cells["B10"].Value ?? (object)"").ToString() == "Netiešā pārstāvība")
                xmlWriter.WriteElementString("TypOfPriGHEA1003", "3");
            xmlWriter.WriteStartElement("TRAREP");
            xmlWriter.WriteElementString("NamTRE1", (worksheet.Cells["B13"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("StrAndNumTRE1", (worksheet.Cells["B14"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("PosCodTRE1", (worksheet.Cells["B16"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("CitTRE1", (worksheet.Cells["B15"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("CouCodTRE1", (worksheet.Cells["B17"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("TINTRE1", (worksheet.Cells["B12"].Value ?? (object)"").ToString());
            xmlWriter.WriteEndElement();

            xmlWriter.WriteStartElement("TRACONCO1");
            xmlWriter.WriteElementString("NamCO17", (worksheet.Cells["B21"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("StrAndNumCO122", (worksheet.Cells["B22"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("PosCodCO123", (worksheet.Cells["B24"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("CitCO124", (worksheet.Cells["B23"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("CouCO125", (worksheet.Cells["B25"].Value ?? (object)"").ToString());
            xmlWriter.WriteEndElement();

            xmlWriter.WriteStartElement("SELLER1");
            xmlWriter.WriteElementString("NamSEL1", (worksheet.Cells["B30"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("StrAndNumSEL1", (worksheet.Cells["B31"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("PosCodSEL1", (worksheet.Cells["B33"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("CitSEL1", (worksheet.Cells["B32"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("CouCodSEL1", (worksheet.Cells["B34"].Value ?? (object)"").ToString());
            xmlWriter.WriteEndElement();

            xmlWriter.WriteStartElement("TRACONCE1");
            xmlWriter.WriteElementString("NamCE17", (worksheet.Cells["B38"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("StrAndNumCE122", (worksheet.Cells["B39"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("PosCodCE123", (worksheet.Cells["B41"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("CitCE124", (worksheet.Cells["B40"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("CouCE125", (worksheet.Cells["B42"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("TINCE159", (worksheet.Cells["B37"].Value ?? (object)"").ToString());
            xmlWriter.WriteEndElement();

            xmlWriter.WriteElementString("DelReqCodHEA992", (worksheet.Cells["B47"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("DesOfDelCou", (worksheet.Cells["B48"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("DesOfDelReqHEA993", (worksheet.Cells["B49"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("InlTraModHEA75", (worksheet.Cells["B53"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("IdeOfMeaOfTraAtDTYPE", (worksheet.Cells["B54"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("IdeOfMeaOfTraAtDHEA78", (worksheet.Cells["B55"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("TraModAtBorHEA76", (worksheet.Cells["B56"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("NatOfMeaOfTraCroHEA87", (worksheet.Cells["B57"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("CouOfDisCodHEA55", (worksheet.Cells["B61"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("CouOfDesCodHEA30", (worksheet.Cells["B62"].Value ?? (object)"").ToString());
            //xmlWriter.WriteElementString("Cur126", (worksheet.Cells["C81"].Value ?? (object)"").ToString());

            xmlWriter.WriteStartElement("STA");
            xmlWriter.WriteStartElement("InvAmSTA"); // 4/11 Rēķina kopsumma
            xmlWriter.WriteElementString("AmSTAElm", (worksheet.Cells["B81"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("CurSTAElm", (worksheet.Cells["C81"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("DivMetAmSTA", "1");
            xmlWriter.WriteEndElement();
            if (worksheet.Cells["B83"].Value != null)
            {
                xmlWriter.WriteStartElement("AddCstAmSTA"); // 4/9 Pieskaitāmās izmaksas
                xmlWriter.WriteElementString("PozID", "206928335");
                xmlWriter.WriteElementString("AmSTAElm", (worksheet.Cells["B83"].Value ?? (object)"").ToString());
                xmlWriter.WriteElementString("CurSTAElm", (worksheet.Cells["C83"].Value ?? (object)"").ToString());
                if ((worksheet.Cells["G83"].Value ?? (object)"").ToString() == "Manuāli")
                    xmlWriter.WriteElementString("DivMetAmSTA", "1");
                else if ((worksheet.Cells["G83"].Value ?? (object)"").ToString() == "Pēc svara")
                    xmlWriter.WriteElementString("DivMetAmSTA", "2");
                else if ((worksheet.Cells["G83"].Value ?? (object)"").ToString() == "Pēc vērtības")
                    xmlWriter.WriteElementString("DivMetAmSTA", "3");
                else if ((worksheet.Cells["G83"].Value ?? (object)"").ToString() == "Nav jāsadala")
                    xmlWriter.WriteElementString("DivMetAmSTA", "4");
                xmlWriter.WriteEndElement();
            }

            if (worksheet.Cells["B84"].Value != null)
            {
                xmlWriter.WriteStartElement("AddCstAmSTA"); // 4/9 Pieskaitāmās izmaksas
                xmlWriter.WriteElementString("PozID", "206928340");
                xmlWriter.WriteElementString("AmSTAElm", (worksheet.Cells["B84"].Value ?? (object)"").ToString());
                xmlWriter.WriteElementString("CurSTAElm", (worksheet.Cells["C84"].Value ?? (object)"").ToString());
                if ((worksheet.Cells["G84"].Value ?? (object)"").ToString() == "Manuāli")
                    xmlWriter.WriteElementString("DivMetAmSTA", "1");
                else if ((worksheet.Cells["G84"].Value ?? (object)"").ToString() == "Pēc svara")
                    xmlWriter.WriteElementString("DivMetAmSTA", "2");
                else if ((worksheet.Cells["G84"].Value ?? (object)"").ToString() == "Pēc vērtības")
                    xmlWriter.WriteElementString("DivMetAmSTA", "3");
                else if ((worksheet.Cells["G84"].Value ?? (object)"").ToString() == "Nav jāsadala")
                    xmlWriter.WriteElementString("DivMetAmSTA", "4");
                xmlWriter.WriteEndElement();
            }

            if (worksheet.Cells["B86"].Value != null)
            {
                xmlWriter.WriteStartElement("DedAmSTA"); // 4/9 Atskaitāmās izmaksas
                xmlWriter.WriteElementString("PozID", "206928336");
                xmlWriter.WriteElementString("AmSTAElm", (worksheet.Cells["B86"].Value ?? (object)"").ToString());
                xmlWriter.WriteElementString("CurSTAElm", (worksheet.Cells["C86"].Value ?? (object)"").ToString());
                if ((worksheet.Cells["G86"].Value ?? (object)"").ToString() == "Manuāli")
                    xmlWriter.WriteElementString("DivMetAmSTA", "1");
                else if ((worksheet.Cells["G86"].Value ?? (object)"").ToString() == "Pēc svara")
                    xmlWriter.WriteElementString("DivMetAmSTA", "2");
                else if ((worksheet.Cells["G86"].Value ?? (object)"").ToString() == "Pēc vērtības")
                    xmlWriter.WriteElementString("DivMetAmSTA", "3");
                else if ((worksheet.Cells["G86"].Value ?? (object)"").ToString() == "Nav jāsadala")
                    xmlWriter.WriteElementString("DivMetAmSTA", "4");
                xmlWriter.WriteEndElement();
            }

            if (worksheet.Cells["B88"].Value != null)
            {
                xmlWriter.WriteStartElement("OthCstAmSTA"); // Pārējās izmaksas
                xmlWriter.WriteElementString("PozID", "206928337");
                xmlWriter.WriteElementString("AmSTAElm", (worksheet.Cells["B88"].Value ?? (object)"").ToString());
                xmlWriter.WriteElementString("CurSTAElm", (worksheet.Cells["C88"].Value ?? (object)"").ToString());
                if ((worksheet.Cells["G88"].Value ?? (object)"").ToString() == "Manuāli")
                    xmlWriter.WriteElementString("DivMetAmSTA", "1");
                else if ((worksheet.Cells["G88"].Value ?? (object)"").ToString() == "Pēc svara")
                    xmlWriter.WriteElementString("DivMetAmSTA", "2");
                else if ((worksheet.Cells["G88"].Value ?? (object)"").ToString() == "Pēc vērtības")
                    xmlWriter.WriteElementString("DivMetAmSTA", "3");
                else if ((worksheet.Cells["G88"].Value ?? (object)"").ToString() == "Nav jāsadala")
                    xmlWriter.WriteElementString("DivMetAmSTA", "4");
                xmlWriter.WriteEndElement();
            }

            if (worksheet.Cells["B90"].Value != null)
            {
                xmlWriter.WriteStartElement("InlTAXAmSTA"); // Izmaksas par pakalpojumiem (PVN bāzei)
                xmlWriter.WriteElementString("PozID", "206928338");
                xmlWriter.WriteElementString("AmSTAElm", (worksheet.Cells["B90"].Value ?? (object)"").ToString());
                xmlWriter.WriteElementString("CurSTAElm", (worksheet.Cells["C90"].Value ?? (object)"").ToString());
                if ((worksheet.Cells["G90"].Value ?? (object)"").ToString() == "Manuāli")
                    xmlWriter.WriteElementString("DivMetAmSTA", "1");
                else if ((worksheet.Cells["G90"].Value ?? (object)"").ToString() == "Pēc svara")
                    xmlWriter.WriteElementString("DivMetAmSTA", "2");
                else if ((worksheet.Cells["G90"].Value ?? (object)"").ToString() == "Pēc vērtības")
                    xmlWriter.WriteElementString("DivMetAmSTA", "3");
                else if ((worksheet.Cells["G90"].Value ?? (object)"").ToString() == "Nav jāsadala")
                    xmlWriter.WriteElementString("DivMetAmSTA", "4");
                xmlWriter.WriteEndElement();
            }
            
            if (worksheet.Cells["B91"].Value != null)
            {
                xmlWriter.WriteStartElement("InlTAXAmSTA"); // Izmaksas par pakalpojumiem (PVN bāzei)
                xmlWriter.WriteElementString("PozID", "206928339");
                xmlWriter.WriteElementString("AmSTAElm", (worksheet.Cells["B91"].Value ?? (object)"").ToString());
                xmlWriter.WriteElementString("CurSTAElm", (worksheet.Cells["C91"].Value ?? (object)"").ToString());
                if ((worksheet.Cells["G91"].Value ?? (object)"").ToString() == "Manuāli")
                    xmlWriter.WriteElementString("DivMetAmSTA", "1");
                else if ((worksheet.Cells["G91"].Value ?? (object)"").ToString() == "Pēc svara")
                    xmlWriter.WriteElementString("DivMetAmSTA", "2");
                else if ((worksheet.Cells["G91"].Value ?? (object)"").ToString() == "Pēc vērtības")
                    xmlWriter.WriteElementString("DivMetAmSTA", "3");
                else if ((worksheet.Cells["G91"].Value ?? (object)"").ToString() == "Nav jāsadala")
                    xmlWriter.WriteElementString("DivMetAmSTA", "4");
                xmlWriter.WriteEndElement();
            } 
            xmlWriter.WriteEndElement();
            

            xmlWriter.WriteStartElement("GUARANTEE");
            xmlWriter.WriteElementString("GuaTypGUATyp", (worksheet.Cells["B94"].Value ?? (object)"").ToString());
            xmlWriter.WriteElementString("GuaTypGUAId", (worksheet.Cells["B95"].Value ?? (object)"").ToString());
            xmlWriter.WriteEndElement();

            xmlWriter.WriteEndElement();



            for (int index2 = 104; index2 > 50; index2++)
            {
                if (worksheet.Cells[index2, 1].Value == null)
                {
                    goto afterGoods;
                }
                
                    xmlWriter.WriteStartElement("GOOITEGDS");
                    xmlWriter.WriteElementString("IteNumB32F1", (worksheet.Cells[index1, 1].Value ?? (object)"").ToString());
                    xmlWriter.WriteElementString("ComCodTarCodGDS10", (worksheet.Cells[index1, 6].Value ?? (object)"").ToString());
                    xmlWriter.WriteElementString("GooDesGDS23", (worksheet.Cells[index1, 20].Value ?? (object)"").ToString());
                    xmlWriter.WriteElementString("CusProCodGDS379", (worksheet.Cells[index1, 26].Value ?? (object)"").ToString());
                    xmlWriter.WriteElementString("AddCusProCodGDS340", (worksheet.Cells[index1, 27].Value ?? (object)"").ToString());
                    xmlWriter.WriteElementString("GroMasGDS46", (worksheet.Cells[index1, 4].Value ?? (object)"").ToString());
                    xmlWriter.WriteElementString("NetMasGDS48", (worksheet.Cells[index1, 3].Value ?? (object)"").ToString());
                if (worksheet.Cells[index1, 18].Value != null)
                {
                    xmlWriter.WriteElementString("QuaOfGooGDS376", (worksheet.Cells[index1, 18].Value ?? (object)"").ToString());
                }
                    xmlWriter.WriteElementString("CouOfOriCodGDS63", (worksheet.Cells[index1, 8].Value ?? (object)"").ToString());
                    xmlWriter.WriteElementString("PreB36", (worksheet.Cells[index1, 28].Value ?? (object)"").ToString());
                    xmlWriter.WriteElementString("AppRecMet", (worksheet.Cells[index1, 29].Value ?? (object)"").ToString());
                    xmlWriter.WriteStartElement("GooInvSTA");
                xmlWriter.WriteElementString("AmSTAElm", (worksheet.Cells[index1, 5].Value ?? (object)"").ToString());
                xmlWriter.WriteElementString("CurSTAElm", (worksheet.Cells["C81"].Value ?? (object)"").ToString());   
                xmlWriter.WriteElementString("DivMetAmSTA", "1");
                xmlWriter.WriteEndElement();
                if (worksheet.Cells["B83"].Value != null)
                {
                    xmlWriter.WriteStartElement("AddGooAmSTA");
                    xmlWriter.WriteElementString("PozID", "206928335");
                    xmlWriter.WriteElementString("AmSTAElm", (worksheet.Cells[index1, 12].Value ?? (object)"").ToString());
                    xmlWriter.WriteElementString("CurSTAElm", (worksheet.Cells["C83"].Value ?? (object)"").ToString());
                    if ((worksheet.Cells["G83"].Value ?? (object)"").ToString() == "Manuāli")
                        xmlWriter.WriteElementString("DivMetAmSTA", "1");
                    else if ((worksheet.Cells["G83"].Value ?? (object)"").ToString() == "Pēc svara")
                        xmlWriter.WriteElementString("DivMetAmSTA", "2");
                    else if ((worksheet.Cells["G83"].Value ?? (object)"").ToString() == "Pēc vērtības")
                        xmlWriter.WriteElementString("DivMetAmSTA", "3");
                    else if ((worksheet.Cells["G83"].Value ?? (object)"").ToString() == "Nav jāsadala")
                        xmlWriter.WriteElementString("DivMetAmSTA", "4");
                    xmlWriter.WriteEndElement();
                }

                if (worksheet.Cells["B84"].Value != null)
                {
                    xmlWriter.WriteStartElement("AddGooAmSTA");
                    xmlWriter.WriteElementString("PozID", "206928340");
                    xmlWriter.WriteElementString("AmSTAElm", (worksheet.Cells[index1, 13].Value ?? (object)"").ToString());
                    xmlWriter.WriteElementString("CurSTAElm", (worksheet.Cells["C84"].Value ?? (object)"").ToString());
                    if ((worksheet.Cells["G84"].Value ?? (object)"").ToString() == "Manuāli")
                        xmlWriter.WriteElementString("DivMetAmSTA", "1");
                    else if ((worksheet.Cells["G84"].Value ?? (object)"").ToString() == "Pēc svara")
                        xmlWriter.WriteElementString("DivMetAmSTA", "2");
                    else if ((worksheet.Cells["G84"].Value ?? (object)"").ToString() == "Pēc vērtības")
                        xmlWriter.WriteElementString("DivMetAmSTA", "3");
                    else if ((worksheet.Cells["G84"].Value ?? (object)"").ToString() == "Nav jāsadala")
                        xmlWriter.WriteElementString("DivMetAmSTA", "4");
                    xmlWriter.WriteEndElement();
                }

                if (worksheet.Cells["B90"].Value != null)
                {
                    xmlWriter.WriteStartElement("InlGooTAXAmSTA");
                    xmlWriter.WriteElementString("PozID", "206928338");
                    xmlWriter.WriteElementString("AmSTAElm", (worksheet.Cells[index1, 17].Value ?? (object)"").ToString());
                    xmlWriter.WriteElementString("CurSTAElm", (worksheet.Cells["C90"].Value ?? (object)"").ToString());
                    if ((worksheet.Cells["G90"].Value ?? (object)"").ToString() == "Manuāli")
                        xmlWriter.WriteElementString("DivMetAmSTA", "1");
                    else if ((worksheet.Cells["G90"].Value ?? (object)"").ToString() == "Pēc svara")
                        xmlWriter.WriteElementString("DivMetAmSTA", "2");
                    else if ((worksheet.Cells["G90"].Value ?? (object)"").ToString() == "Pēc vērtības")
                        xmlWriter.WriteElementString("DivMetAmSTA", "3");
                    else if ((worksheet.Cells["G90"].Value ?? (object)"").ToString() == "Nav jāsadala")
                        xmlWriter.WriteElementString("DivMetAmSTA", "4");
                    xmlWriter.WriteEndElement();
                }
                if (worksheet.Cells["B91"].Value != null)
                {
                    xmlWriter.WriteStartElement("InlGooTAXAmSTA");
                    xmlWriter.WriteElementString("PozID", "206928339");
                    xmlWriter.WriteElementString("AmSTAElm", (worksheet.Cells[index1, 16].Value ?? (object)"").ToString());
                    xmlWriter.WriteElementString("CurSTAElm", (worksheet.Cells["C91"].Value ?? (object)"").ToString());
                                   if ((worksheet.Cells["G91"].Value ?? (object)"").ToString() == "Manuāli")
                    xmlWriter.WriteElementString("DivMetAmSTA", "1");
                else if ((worksheet.Cells["G91"].Value ?? (object)"").ToString() == "Pēc svara")
                    xmlWriter.WriteElementString("DivMetAmSTA", "2");
                else if ((worksheet.Cells["G91"].Value ?? (object)"").ToString() == "Pēc vērtības")
                    xmlWriter.WriteElementString("DivMetAmSTA", "3");
                else if ((worksheet.Cells["G91"].Value ?? (object)"").ToString() == "Nav jāsadala")
                    xmlWriter.WriteElementString("DivMetAmSTA", "4");
                    xmlWriter.WriteEndElement();
                }

                for (int index3 = 35; index3 > 10; index3=index3+3)
                    { 
                        if (worksheet.Cells[index1, index3].Value == null)
                        { 
                            goto afterProDoc;
                         }
                        xmlWriter.WriteStartElement("PRODOCDC2");
                        xmlWriter.WriteElementString("DocCoDC28", (worksheet.Cells[index1, index3].Value ?? (object)"").ToString());
                        xmlWriter.WriteElementString("TitDC29", (worksheet.Cells[index1, index3 + 1].Value ?? (object)"").ToString());
                        xmlWriter.WriteElementString("ComOfInfDC25", (worksheet.Cells[index1, index3 + 2].Value ?? (object)"").ToString());
                        xmlWriter.WriteEndElement();
                    
                }
                afterProDoc:
                   xmlWriter.WriteStartElement("PACGS2");
                    xmlWriter.WriteElementString("KinOfPacGS23", (worksheet.Cells[index1, 31].Value ?? (object)"").ToString());
                    xmlWriter.WriteElementString("MarNumOfPacGS21", (worksheet.Cells[index1, 32].Value ?? (object)"").ToString());
                    xmlWriter.WriteElementString("NumOfPacGS24", (worksheet.Cells[index1, 2].Value ?? (object)"").ToString());
                    xmlWriter.WriteEndElement();
                    xmlWriter.WriteStartElement("PREADMREFAR2");
                    xmlWriter.WriteElementString("PreDocTyp", (worksheet.Cells[index1, 22].Value ?? (object)"").ToString());
                    xmlWriter.WriteElementString("TitAR212", (worksheet.Cells[index1, 23].Value ?? (object)"").ToString());
                    xmlWriter.WriteElementString("PreDocRefAR26", (worksheet.Cells[index1, 7].Value ?? (object)"").ToString());
                    xmlWriter.WriteElementString("PreDocItemNumb", (worksheet.Cells[index1, 10].Value ?? (object)"").ToString());
                    xmlWriter.WriteElementString("AllGoodsFlag", "1");
                    xmlWriter.WriteEndElement();
                    xmlWriter.WriteEndElement();
                    index1++;
                }
        afterGoods:



            /*

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

            if (worksheet.Cells[index1, 29].Value != null)
            {
                xmlWriter.WriteStartElement("CONNR2");
                xmlWriter.WriteElementString("ConNumNR21", (worksheet.Cells[index1, 29].Value ?? (object)"").ToString());
                xmlWriter.WriteEndElement();
            }

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
        xmlWriter.WriteEndElement();


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
        */

            xmlWriter.Flush();
                xmlWriter.Close();

                File.WriteAllText(currentDirectory + "\\XML_IM_" + DateTime.Now.ToString("yyyyMMddHHmm") + ".xml", new UTF8Encoding(false).GetString(memoryStream.ToArray()));
            }
        }
    }
