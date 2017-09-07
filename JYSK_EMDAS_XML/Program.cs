using OfficeOpenXml;
using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace JYSK_EMDAS_XML //Version 0.1.2 //
{
    class Program
    {
        static void Main(string[] args)
        {
            
            Console.WriteLine("Converting Excel....");
            string curr = Directory.GetCurrentDirectory();
                FileInfo fileName = new FileInfo(curr+"\\EX1.xlsx");
                ExcelPackage pck = new ExcelPackage(fileName);
                var ws = pck.Workbook.Worksheets["EX1 sagatave"];

            System.Globalization.CultureInfo customCulture = (System.Globalization.CultureInfo)System.Threading.Thread.CurrentThread.CurrentCulture.Clone();
            customCulture.NumberFormat.NumberDecimalSeparator = ".";

            System.Threading.Thread.CurrentThread.CurrentCulture = customCulture;

            var row = 76;
                var n = 76;
                var i = 22;

                var rng = new Random();
                int first = rng.Next(10);
                int second = rng.Next(10);
                int third = rng.Next(10);

                XmlWriterSettings xmlWriterSettings = new XmlWriterSettings();
            xmlWriterSettings.Encoding = new UTF8Encoding(false);
                MemoryStream ms = new MemoryStream();
                XmlWriter writer = XmlWriter.Create(ms, xmlWriterSettings);
            // MESSAGE //

            writer.WriteStartElement("CC515A");
                //writer.WriteElementString("SynIdeMES1", "MESSAGE.Syntax identifier - a4");
                //writer.WriteElementString("SynVerNumMES2", "MESSAGE.Syntax version number - n1");
                //writer.WriteElementString("MesSenMES3", " MESSAGE.Message sender  - an..35");
                //writer.WriteElementString("MesRecMES6", " MESSAGE.Message recipient - an..35");
                writer.WriteElementString("DatOfPreMES9", DateTime.Now.ToString("yyMMdd"));
                writer.WriteElementString("TimOfPreMES10", DateTime.Now.ToString("HHmm"));
                //writer.WriteElementString("IntConRefMES11", "MESSAGE.Interchange control reference - an..14");
                //writer.WriteElementString("MesIdeMES19", "MESSAGE.Message identification  - an..14");
                writer.WriteElementString("MesTypMES20", "CC515A");

                // HEADER //

                writer.WriteStartElement("HEAHEA");
                //writer.WriteElementString("RefNumHEA4", "MESSAGE - HEADER.Reference number  - an..22");
                writer.WriteElementString("TypOfDecHEA24", (ws.Cells["B4"].Value ?? "").ToString());
                writer.WriteElementString("CouOfDesCodHEA30", (ws.Cells["B28"].Value ?? "").ToString());
                writer.WriteElementString("AgrLocOfGooCodHEA38", (ws.Cells["B53"].Value ?? "").ToString());
                writer.WriteElementString("AgrLocOfGooHEA39", (ws.Cells["B52"].Value ?? "").ToString());
                writer.WriteElementString("AgrLocOfGooHEA39LNG", (ws.Cells["B14"].Value ?? "").ToString());
                writer.WriteElementString("CouOfDisCodHEA55", (ws.Cells["B14"].Value ?? "").ToString());
                writer.WriteElementString("InlTraModHEA75", (ws.Cells["B48"].Value ?? "").ToString());
                writer.WriteElementString("TraModAtBorHEA76", (ws.Cells["B49"].Value ?? "").ToString());
                //writer.WriteElementString("ConIndHEA96", "a");
                writer.WriteElementString("ECSAccDocHEA601", "LV");
            //writer.WriteElementString("TotNumOfIteHEA305", "a");
            writer.WriteElementString("DecPlaHEA394", (ws.Cells["B71"].Value ?? "").ToString());
            //writer.WriteElementString("DecDatHEA383", "a");
                writer.WriteElementString("TypOfDecBx12HEA651", (ws.Cells["C4"].Value ?? "").ToString());
                writer.WriteEndElement();

                // TRADER EXPORTER //

                writer.WriteStartElement("TRAEXPEX1");
                writer.WriteElementString("NamEX17", (ws.Cells["B17"].Value ?? "").ToString());
                writer.WriteElementString("StrAndNumEX122", (ws.Cells["B18"].Value ?? "").ToString());
                writer.WriteElementString("PosCodEX123", (ws.Cells["B20"].Value ?? "").ToString());
                writer.WriteElementString("CitEX124", (ws.Cells["B19"].Value ?? "").ToString());
                writer.WriteElementString("CouEX125", (ws.Cells["B21"].Value ?? "").ToString());
                writer.WriteElementString("TINEX159", (ws.Cells["B16"].Value ?? "").ToString());
                writer.WriteEndElement();

                // TRADER CONSIGNEE //

                writer.WriteStartElement("TRACONCE1");
                writer.WriteElementString("NamCE17", (ws.Cells["B24"].Value ?? "").ToString());
                writer.WriteElementString("StrAndNumCE122", (ws.Cells["B25"].Value ?? "").ToString());
                writer.WriteElementString("PosCodCE123", (ws.Cells["B27"].Value ?? "").ToString());
                writer.WriteElementString("CitCE124", (ws.Cells["B26"].Value ?? "").ToString());
                writer.WriteElementString("CouCE125", (ws.Cells["B28"].Value ?? "").ToString());
                writer.WriteEndElement();

                // CUSTOMS OFFICE - EXPORT //

                writer.WriteStartElement("CUSOFFEXPERT");
                writer.WriteElementString("RefNumERT1", (ws.Cells["B35"].Value ?? "").ToString());
                writer.WriteEndElement();

                // CUSTOMS OFFICE - EXIT //

                writer.WriteStartElement("CUSOFFEXIEXT");
                writer.WriteElementString("RefNumEXT1", (ws.Cells["B36"].Value ?? "").ToString());
                writer.WriteEndElement();

                // CONTROL RESULTS //

                //writer.WriteStartElement("CONRESERS");
                //writer.WriteElementString("ConResCodERS16", "MESSAGE - CONTROL RESULT.Control result code - an2");
                //writer.WriteElementString("DatLimERS69", "MESSAGE - CONTROL RESULT.Date limit - n8");
                //writer.WriteEndElement();

                // SEALS INFO //

               // writer.WriteStartElement("SEAINFSLI");
               // writer.WriteElementString("SeaNumSLI2", "MESSAGE - SEALS INFO.Seals number - n..4 ");
                // SEALS ID //
                //writer.WriteStartElement("SEAIDSID");
                //writer.WriteElementString("SeaIdeSID1", "MESSAGE - SEALS INFO - SEALS ID.Seals identity - an..20");
                //writer.WriteEndElement();
                //writer.WriteEndElement();


                // GOODS ITEM //

                for (n = 76; n > 50; n++)
                {
                    if (ws.Cells[n, 1].Value == null)
                    {
                        goto afterGoods;
                    }

                    writer.WriteStartElement("GOOITEGDS");
                    writer.WriteElementString("IteNumGDS7", (ws.Cells[row, 1].Value ?? "").ToString());
                    writer.WriteElementString("GooDesGDS23", (ws.Cells[row, 13].Value ?? "").ToString());
                    //writer.WriteElementString("GooDesGDS23LNG", (ws.Cells[row, 8].Value ?? "").ToString());
                    writer.WriteElementString("GroMasGDS46", (ws.Cells[row, 4].Value ?? "").ToString());
                    writer.WriteElementString("NetMasGDS48", (ws.Cells[row, 3].Value ?? "").ToString());
                    writer.WriteElementString("ProReqGDI1", (ws.Cells[row, 17].Value ?? "").ToString());
                    writer.WriteElementString("PreProGDI1", (ws.Cells[row, 18].Value ?? "").ToString());
                    writer.WriteElementString("ComNatProGIM1", (ws.Cells[row, 19].Value ?? "").ToString());
                    writer.WriteElementString("StaValAmoGDI1", (ws.Cells[row, 5].Value ?? "").ToString());
                    writer.WriteElementString("CouOfOriGDI1", (ws.Cells[row, 8].Value ?? "").ToString());
                    writer.WriteElementString("SupUniGDI1", (ws.Cells[row, 25].Value ?? "").ToString());
                    // PREVIOUS ADMINISTRATIVE REFERENCES //
                    writer.WriteStartElement("PREADMREFAR2");
                    writer.WriteElementString("PreDocTypAR21", (ws.Cells[row, 16].Value ?? "").ToString());
                    writer.WriteElementString("PreDocRefAR26", (ws.Cells[row, 7].Value ?? "").ToString());
                    //writer.WriteElementString("PreDocRefLNG", (ws.Cells[row, 2].Value ?? "").ToString());
                    writer.WriteElementString("PreDocCatPREADMREF21", (ws.Cells[row, 15].Value ?? "").ToString());
                    writer.WriteEndElement();
                    // PRODUCED DOCUMENTS/CERTIFICATES //
                    for (i = 26; i > 10; i=i+3)
                    {
                    if (ws.Cells[row, i].Value == null)
                    {
                        goto afterProDoc;
                    }
                    writer.WriteStartElement("PRODOCDC2");
                    writer.WriteElementString("DocTypDC21", (ws.Cells[row, i].Value ?? "").ToString());
                    writer.WriteElementString("DocRefDC23", (ws.Cells[row, i+1].Value ?? "").ToString());
                    //writer.WriteElementString("DocRefDCLNG", (ws.Cells[row, 2].Value ?? "").ToString());
                    writer.WriteEndElement();
                    }
                    afterProDoc:
                    // SPECIAL MENTIONS //
                    //writer.WriteStartElement("SPEMENMT2");
                    //writer.WriteElementString("AddInfCodMT23", "MESSAGE - GOODS ITEM - SPECIAL MENTIONS.Additional information coded - an..5");
                    //writer.WriteEndElement();
                    // TRADER EXPORTER //
                    //writer.WriteStartElement("TRACONEX2");
                    //writer.WriteElementString("TINEX259", "MESSAGE - GOODS ITEM - (EXPORTER) TRADER.TIN - an..17");
                    //writer.WriteEndElement();
                    // COMMODITY CODE //
                    writer.WriteStartElement("COMCODGODITM");
                    writer.WriteElementString("ComNomCMD1", (ws.Cells[row, 6].Value ?? "").ToString());
                    writer.WriteElementString("TARCodCMD1", (ws.Cells[row, 10].Value ?? "").ToString());
                    writer.WriteElementString("TARFirAddCodCMD1", (ws.Cells[row, 23].Value ?? "").ToString());
                    //writer.WriteElementString("TARSecAddCodCMD1", (ws.Cells[row, 2].Value ?? "").ToString());
                    writer.WriteEndElement();
                    // CALCULATION TAXES //
                    //writer.WriteStartElement("CALTAXGOD");
                    //writer.WriteElementString("TypOfTaxCTX1", "MESSAGE - GOODS ITEM - (TAXES) CALCULATION.Type of tax - an3");
                    //writer.WriteElementString("TaxBasCTX1", "MESSAGE - GOODS ITEM - (TAXES) CALCULATION.Tax base - n..15,2");
                    //writer.WriteElementString("RatOfTaxCTX1", "MESSAGE - GOODS ITEM - (TAXES) CALCULATION.Rate of tax - an..15");
                    //writer.WriteElementString("AmoOfTaxTCL1", "MESSAGE - GOODS ITEM - (TAXES) CALCULATION.Amount of tax - n..15,2");
                    //writer.WriteElementString("MetOfPayCTX1", "MESSAGE - GOODS ITEM - (TAXES) CALCULATION.Method of payment - a1");
                    //writer.WriteEndElement();
                    // TRADER CONSIGNEE //
                    //writer.WriteStartElement("TRACONCE2");
                    //writer.WriteElementString("NamCE27", (ws.Cells[row, 2].Value ?? "").ToString());
                    //writer.WriteElementString("StrAndNumCE222", (ws.Cells[row, 2].Value ?? "").ToString());
                    //writer.WriteElementString("PosCodCE223", (ws.Cells[row, 2].Value ?? "").ToString());
                    //writer.WriteElementString("CitCE224", (ws.Cells[row, 2].Value ?? "").ToString());
                    //writer.WriteElementString("CouCE225", (ws.Cells[row, 2].Value ?? "").ToString());
                    //writer.WriteElementString("TINCE259", (ws.Cells[row, 2].Value ?? "").ToString());
                    //writer.WriteEndElement();
                    // CONTAINERS //
                    //writer.WriteStartElement("CONNR2");
                    //writer.WriteElementString("ConNumNR21", "MESSAGE - GOODS ITEM - CONTAINERS.Container number - an..17");
                    //writer.WriteElementString("TaxBasCTX1", "MESSAGE - GOODS ITEM - (TAXES) CALCULATION.Tax base - n..15,2");
                    //writer.WriteElementString("RatOfTaxCTX1", "MESSAGE - GOODS ITEM - (TAXES) CALCULATION.Rate of tax - an..15");
                    //writer.WriteElementString("AmoOfTaxTCL1", "MESSAGE - GOODS ITEM - (TAXES) CALCULATION.Amount of tax - n..15,2");
                    //writer.WriteElementString("MetOfPayCTX1", "MESSAGE - GOODS ITEM - (TAXES) CALCULATION.Method of payment - a1");
                    //writer.WriteEndElement();
                    // CONTAINERS //
                    writer.WriteStartElement("PACGS2");
                    writer.WriteElementString("MarNumOfPacGS21", (ws.Cells[row, 22].Value ?? "").ToString());
                    //writer.WriteElementString("MarNumOfPacGS21LNG", (ws.Cells[row, 2].Value ?? "").ToString());
                    writer.WriteElementString("KinOfPacGS23", (ws.Cells[row, 21].Value ?? "").ToString());
                    writer.WriteElementString("NumOfPacGS24", (ws.Cells[row, 2].Value ?? "").ToString());
                    writer.WriteEndElement();

                    writer.WriteEndElement();


                    //do not edit
                    row++;
                }
            afterGoods:
                // ITINERARY //

                //writer.WriteStartElement("ITI");
                //writer.WriteElementString("CouOfRouCodITI1", " MESSAGE - ITINERARY.Country of routing code - a2");
                //writer.WriteEndElement();

                // TRADE DECLARANT //

                writer.WriteStartElement("TRADEC");
                writer.WriteElementString("NamTDE1", (ws.Cells["B10"].Value ?? "").ToString());
                writer.WriteElementString("StrAndNumTDE1", (ws.Cells["B11"].Value ?? "").ToString());
                writer.WriteElementString("PosTDE1", (ws.Cells["B13"].Value ?? "").ToString());
                writer.WriteElementString("CiTDE1", (ws.Cells["B12"].Value ?? "").ToString());
                writer.WriteElementString("CouCodTDE1", (ws.Cells["B14"].Value ?? "").ToString());
                writer.WriteElementString("TINTDE1", (ws.Cells["B9"].Value ?? "").ToString());
                writer.WriteEndElement();

                // DELIVERY TERMS //

                //writer.WriteStartElement("DELTER");
                //writer.WriteElementString("IncCodTDL1", "MESSAGE - (TERMS) DELIVERY.Incoterm Code - a3");
                //writer.WriteElementString("ComInfDELTER387", "MESSAGE - (TERMS) DELIVERY.Complement of info - an..35");
                //writer.WriteEndElement();

                // TRANSACTION DATA //

                writer.WriteStartElement("TRADAT");
                writer.WriteElementString("CurTRD1", (ws.Cells["B58"].Value ?? "").ToString());
                //writer.WriteElementString("TotAmoInvTRD1", (ws.Cells["B58"].Value ?? "").ToString());
                //writer.WriteElementString("ExcRatTRD1", (ws.Cells["B58"].Value ?? "").ToString());
                writer.WriteEndElement();

                // DEFERRED OR POSTPONED PAYMENT //

                //writer.WriteStartElement("DEFPOSPAY");
                //writer.WriteElementString("AutRefDPP1", "MESSAGE - (PAYMENT) DEFERRED OR POSTPONED.Authorisation Reference - an..35");
                //writer.WriteEndElement();

                // IDENTIFICATION WAREHOUSE //

                writer.WriteStartElement("IDEWAR");
                writer.WriteElementString("WarTypWID1", (ws.Cells["B66"].Value ?? "").ToString());
                writer.WriteElementString("AutCouWID1", (ws.Cells["B68"].Value ?? "").ToString()); //NULL
                writer.WriteElementString("WarIdeWID1", (ws.Cells["B67"].Value ?? "").ToString());
                writer.WriteEndElement();

                // REPRESENTIVE STATUS //

                writer.WriteStartElement("STATREP385");
                writer.WriteElementString("RepStatCodSTATREP386", (ws.Cells["B7"].Value ?? "").ToString());
                writer.WriteEndElement();

                // WRITE XML //

                writer.Flush();
                writer.Close();
            
                var xmlOut = curr+"\\XML_" + DateTime.Now.ToString("yyyyMMddHHmm") + ".xml";
                File.WriteAllText(@xmlOut, new UTF8Encoding(false).GetString(ms.ToArray()));

            // CONSOLE OUT //

            //Console.WriteLine(Encoding.UTF8.GetString(ms.ToArray()));
                
                
            }
        }
    }
