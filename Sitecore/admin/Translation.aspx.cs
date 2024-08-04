using UtilityProject.Schema;
using ClosedXML.Excel;
using Sitecore.Diagnostics;
using Sitecore.sitecore.admin;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;

using SC = Sitecore;

namespace UtilityProject
{
    public partial class Translation : AdminPage
    {
        public string RootPath { get; set; }
        public string XMLFileName { get; set; }

        public static readonly string[] FIELD_TYPE = { "Pages","Supported Languages","Supported Time Zone","Location Distance","Cities","Job Type","Onboarding Pages",
            "Components","Role","Values","Pronouns","Email Type","Righttrack Persona","Option Type","Question Design Type","Salary Ranges","Supported Currencies","Child Pages",
            "Focus Area", "Resource Type","Resource Pages","Assessment Type","Image", "Answers","Flag","City","Skill","Component Title"};

        protected void Page_Load(object sender, EventArgs e)
        {
            RootPath = Server.MapPath("~/" + SC.Configuration.Settings.GetSetting("path"));
            XMLFileName = SC.Configuration.Settings.GetSetting("fileName");
            if (!Page.IsPostBack)
            {
                FileUpload1 = new System.Web.UI.WebControls.FileUpload();
                fuTranslator = new System.Web.UI.WebControls.FileUpload();

            }
            else
            {
                lblMsg.Text = string.Empty;
                btnDownload.UseSubmitBehavior = false;
                btnUpdateTranslation.UseSubmitBehavior = false;
            }
        }

        protected void btnUpdateTranslation_Click(object sender, EventArgs e)
        {
            try
            {

                Helper serialize = new Helper();

                if (!string.IsNullOrEmpty(RootPath))
                {
                    if (!string.IsNullOrEmpty(fuTranslator.FileName))
                    {
                        string uploadedFilePath = fuTranslator.FileName;
                        XMLFileName = uploadedFilePath.Substring(0, uploadedFilePath.LastIndexOf("_"));

                        string translationXML = RootPath + XMLFileName + ".xml";
                        string translationXLS = RootPath + fuTranslator.FileName;

                        fuTranslator.SaveAs(translationXLS);

                        var data = serialize.ConvertExcelToDataTable(translationXLS, true);

                        if (data != null)
                        {
                            Sitecore1 sitecore = serialize.DeserializeToObject<Sitecore1>(translationXML);
                            for (int i = 0; i < data.Rows.Count; i++)
                            {
                                DataRow dr = data.Rows[i];
                                string itemId = dr["Itemid"].ToString();
                                string field = dr["Field"].ToString();
                                var scItem = sitecore.Phrase.Where(x => x.Itemid.ToString().ToLower().Equals(itemId.ToLower()) &&
                                x.Fieldid.ToLower().Equals(field.ToLower())).FirstOrDefault();
                                if (scItem != null)
                                {
                                    Log.Info("File successfully uploaded to Azure Blob Storage: " + scItem.ItemName, this);
                                    if (!(dr["CanadaEnglish"] is DBNull))
                                        scItem.EnCA = dr["CanadaEnglish"]?.ToString();
                                    if (!(dr["CanadaFrench"] is DBNull))
                                        scItem.FRCA = dr["CanadaFrench"]?.ToString();
                                    if (!(dr["English"] is DBNull))
                                        scItem.En = dr["English"]?.ToString();
                                    if (!(dr["USEnglish"] is DBNull))
                                        scItem.EnUS = dr["USEnglish"]?.ToString();

                                    if (!(dr["EnglishAustralia"] is DBNull))
                                        scItem.ENAustralia = dr["EnglishAustralia"]?.ToString();
                                    if (!(dr["EnglishUnitedKingdom"] is DBNull))
                                        scItem.ENUnitedKingdom = dr["EnglishUnitedKingdom"]?.ToString();

                                    if (!(dr["FRFrench"] is DBNull))
                                        scItem.FRFrench = dr["FRFrench"]?.ToString();

                                    if (!(dr["BelgiumDutch"] is DBNull))
                                        scItem.BelgiumDutch = dr["BelgiumDutch"]?.ToString();
                                    if (!(dr["BelgiumFrench"] is DBNull))
                                        scItem.BelgiumFrench = dr["BelgiumFrench"]?.ToString();

                                    if (!(dr["GermanyGerman"] is DBNull))
                                        scItem.GermanyGerman = dr["GermanyGerman"]?.ToString();

                                    if (!(dr["JapanJapanese"] is DBNull))
                                        scItem.JapanJapanese = dr["JapanJapanese"]?.ToString();

                                    if (!(dr["NetherlandDutch"] is DBNull))
                                        scItem.NetherlandDutch = dr["NetherlandDutch"]?.ToString();

                                    if (!(dr["SwitzerlandFrench"] is DBNull))
                                        scItem.SwitzerlandFrench = dr["SwitzerlandFrench"]?.ToString();
                                    if (!(dr["SwitzerlandGerman"] is DBNull))
                                        scItem.SwitzerlandGerman = dr["SwitzerlandGerman"]?.ToString();

                                    if (!(dr["MexicoSpanish"] is DBNull))
                                        scItem.MexicoSpanish = dr["MexicoSpanish"]?.ToString();

                                    if (!(dr["BrazilPortugese"] is DBNull))
                                        scItem.BrazilPortugese = dr["BrazilPortugese"]?.ToString();
                                    if (!(dr["PortugalPortugese"] is DBNull))
                                        scItem.PortugalPortugese = dr["PortugalPortugese"]?.ToString();

                                    if (!(dr["ItalyItalian"] is DBNull))
                                        scItem.ItalyItalian = dr["ItalyItalian"]?.ToString();

                                    if (!(dr["SpainSpainish"] is DBNull))
                                        scItem.SpainSpainish = dr["SpainSpainish"]?.ToString();

                                    if (!(dr["DenmarkDanish"] is DBNull))
                                        scItem.DenmarkDanish = dr["DenmarkDanish"]?.ToString();

                                    if (!(dr["FinlandFinnish"] is DBNull))
                                        scItem.FinlandFinnish = dr["FinlandFinnish"]?.ToString();

                                    if (!(dr["PolandPolish"] is DBNull))
                                        scItem.PolandPolish = dr["PolandPolish"]?.ToString();
                                    if (!(dr["SwedenSwedish"] is DBNull))
                                        scItem.SwedenSwedish = dr["SwedenSwedish"]?.ToString();

                                    if (!(dr["NorwayNorwegian"] is DBNull))
                                        scItem.NorwayNorwegian = dr["NorwayNorwegian"]?.ToString();
                                    if (!(dr["ArgentinaSpanish"] is DBNull))
                                        scItem.ArgentinaSpanish = dr["ArgentinaSpanish"]?.ToString();
                                    if (!(dr["KoreaKorean"] is DBNull))
                                        scItem.KoreaKorean = dr["KoreaKorean"]?.ToString();
                                }
                            }

                            serialize.SerializeToXml<Sitecore1>(sitecore, translationXML);

                            lblMsg.Text = "Successfully Updated";
                            lblMsg.ForeColor = Color.DarkGreen;
                        }
                        else
                        {
                            lblMsg.ForeColor = Color.Red;
                            lblMsg.Text = "No Translation Data";
                        }
                    }
                }
                else
                {
                    lblMsg.ForeColor = Color.Red;
                    lblMsg.Text = "File paths are not specified in configuration file";
                }

            }
            catch (Exception ex)
            {
                lblMsg.Text = ex.Message;
                lblMsg.ForeColor = Color.Red;
            }
        }

        protected void btnDownload_Click(object sender, EventArgs e)
        {
            Helper serialize = new Helper();
            //string rootPath = System.Configuration.ConfigurationManager.AppSettings["path"];
            //string fileName = System.Configuration.ConfigurationManager.AppSettings["fileName"];
            //string translationXML = RootPath + XMLFileName;
            try
            {
                string uploadedFilePath = FileUpload1.FileName;

                if (!string.IsNullOrWhiteSpace(uploadedFilePath))
                {

                    string fileNameWithoutExt = uploadedFilePath.Contains('.') ? uploadedFilePath.Substring(0, uploadedFilePath.LastIndexOf(".")) : uploadedFilePath;

                    uploadedFilePath = uploadedFilePath.Contains('.') ? uploadedFilePath : uploadedFilePath + ".xml";
                    string translationXML = RootPath + uploadedFilePath;
                    //string translationXLS = RootPath + fuTranslator.FileName;
                    FileUpload1.SaveAs(translationXML);

                    var content = serialize.DeserializeToObject<Sitecore1>(translationXML);
                    List<TranslationData> translatonList = new List<TranslationData>();
                    int count = 0;
                    foreach (var item in content.Phrase)
                    {
                        Log.Info(string.Format("File successfully uploaded to Azure Blob Storage:{0},{1},{2}", item.Itemid, item.ItemName, ++count), this);
                        Log.Info(string.Format("Cell Value EN:{0}", item.En), this);
                        Log.Info(string.Format("Cell Value EN-CA:{0}", item.EnCA), this);

                        TranslationData data = new TranslationData();
                        data.Itemid = item.Itemid;
                        data.Path = item.Path;
                        data.Fieldid = item.Fieldid;
                        data.English = item.En;
                        data.USEnglish = item.EnUS;
                        data.CanadaEnglish = item.EnCA;
                        data.CanadaFrench = item.FRCA;
                        //new 
                        data.EnglishAustralia = item.ENAustralia;
                        data.EnglishUnitedKingdom = item.ENUnitedKingdom;
                        //French
                        data.FRFrench = item.FRFrench;
                        //Belgium
                        data.BelgiumDutch = item.BelgiumDutch;
                        data.BelgiumFrench = item.BelgiumFrench;
                        //Germany
                        data.GermanyGerman = item.GermanyGerman;
                        //Japan
                        data.JapanJapanese = item.JapanJapanese;
                        //Netherland
                        data.NetherlandDutch = item.NetherlandDutch;
                        //Switzerland
                        data.SwitzerlandFrench = item.SwitzerlandFrench;
                        data.SwitzerlandGerman = item.SwitzerlandGerman;

                        data.MexicoSpanish = item.MexicoSpanish;
                        data.BrazilPortugese = item.BrazilPortugese;
                        data.PortugalPortugese = item.PortugalPortugese;

                        data.ItalyItalian = item.ItalyItalian;

                        data.SpainSpainish = item.SpainSpainish;
                        data.DenmarkDanish = item.DenmarkDanish;
                        data.FinlandFinnish = item.FinlandFinnish;
                        data.PolandPolish = item.PolandPolish;
                        data.SwedenSwedish = item.SwedenSwedish;
                        data.NorwayNorwegian = item.NorwayNorwegian;
                        data.ArgentinaSpanish = item.ArgentinaSpanish;
                        data.KoreaKorean = item.KoreaKorean;

                        translatonList.Add(data);
                    }
                    var workbook = new XLWorkbook();
                    workbook.AddWorksheet("Translation");
                    var ws = workbook.Worksheet("Translation");

                    SetColumnWidthAndWrapText(ws);

                    // Protect a sheet
                    ws.Protect().AllowElement(XLSheetProtectionElements.FormatCells)   // Cell Formatting
                      .AllowElement(XLSheetProtectionElements.InsertColumns) // Inserting Columns
                      .AllowElement(XLSheetProtectionElements.DeleteColumns) // Deleting Columns
                      .AllowElement(XLSheetProtectionElements.DeleteRows);   // Deleting Rows

                    int row = 1;
                    ws.Cell("A" + row.ToString()).Value = "Itemid";
                    ws.Cell("B" + row.ToString()).Value = "Path";
                    ws.Cell("C" + row.ToString()).Value = "Field";
                    ws.Cell("D" + row.ToString()).Value = "English";
                    ws.Cell("E" + row.ToString()).Value = "USEnglish";
                    ws.Cell("F" + row.ToString()).Value = "CanadaFrench";
                    ws.Cell("G" + row.ToString()).Value = "CanadaEnglish";

                    ws.Cell("H" + row.ToString()).Value = "EnglishAustralia";
                    ws.Cell("I" + row.ToString()).Value = "EnglishUnitedKingdom";

                    ws.Cell("J" + row.ToString()).Value = "FRFrench";

                    ws.Cell("K" + row.ToString()).Value = "BelgiumDutch";
                    ws.Cell("L" + row.ToString()).Value = "BelgiumFrench";

                    ws.Cell("M" + row.ToString()).Value = "GermanyGerman";

                    ws.Cell("N" + row.ToString()).Value = "JapanJapanese";

                    ws.Cell("O" + row.ToString()).Value = "NetherlandDutch";

                    ws.Cell("P" + row.ToString()).Value = "SwitzerlandFrench";
                    ws.Cell("Q" + row.ToString()).Value = "SwitzerlandGerman";

                    ws.Cell("R" + row.ToString()).Value = "MexicoSpanish";
                    ws.Cell("S" + row.ToString()).Value = "BrazilPortugese";
                    ws.Cell("T" + row.ToString()).Value = "PortugalPortugese";

                    ws.Cell("U" + row.ToString()).Value = "ItalyItalian";
                    ws.Cell("V" + row.ToString()).Value = "SpainSpainish";

                    ws.Cell("W" + row.ToString()).Value = "DenmarkDanish";
                    ws.Cell("X" + row.ToString()).Value = "FinlandFinnish";
                    ws.Cell("Y" + row.ToString()).Value = "PolandPolish";
                    ws.Cell("Z" + row.ToString()).Value = "SwedenSwedish";
                    ws.Cell("AA" + row.ToString()).Value = "NorwayNorwegian";
                    ws.Cell("AB" + row.ToString()).Value = "ArgentinaSpanish";
                    ws.Cell("AC" + row.ToString()).Value = "KoreaKorean";
                    SetLockOntheCells(ws, row);
                    //Guid result = new Guid();
                    foreach (var trans in translatonList)
                    {
                        row++;
                        ws.Cell("A" + row.ToString()).Value = trans.Itemid;
                        ws.Cell("B" + row.ToString()).Value = trans.Path;
                        ws.Cell("C" + row.ToString()).Value = trans.Fieldid;
                        ws.Cell("D" + row.ToString()).Value = trans.English;
                        ws.Cell("E" + row.ToString()).Value = trans.USEnglish;
                        ws.Cell("F" + row.ToString()).Value = trans.CanadaFrench;
                        ws.Cell("G" + row.ToString()).Value = trans.CanadaEnglish;

                        ws.Cell("H" + row.ToString()).Value = trans.EnglishAustralia;
                        ws.Cell("I" + row.ToString()).Value = trans.EnglishUnitedKingdom;

                        // ws.Cell("J" + row.ToString()).Value = trans.FrenchEnglish;
                        ws.Cell("J" + row.ToString()).Value = trans.FRFrench;

                        ws.Cell("K" + row.ToString()).Value = trans.BelgiumDutch;
                        //ws.Cell("M" + row.ToString()).Value = trans.BelgiumEnglish;
                        ws.Cell("L" + row.ToString()).Value = trans.BelgiumFrench;

                        //ws.Cell("O" + row.ToString()).Value = trans.GermanyEnglish;
                        ws.Cell("M" + row.ToString()).Value = trans.GermanyGerman;

                        //ws.Cell("Q" + row.ToString()).Value = trans.JapanEnglish;
                        ws.Cell("N" + row.ToString()).Value = trans.JapanJapanese;

                        ws.Cell("O" + row.ToString()).Value = trans.NetherlandDutch;
                        //ws.Cell("T" + row.ToString()).Value = trans.NetherlandEnglish;

                        //ws.Cell("U" + row.ToString()).Value = trans.SwitzerlandEnglish;
                        ws.Cell("P" + row.ToString()).Value = trans.SwitzerlandFrench;
                        ws.Cell("Q" + row.ToString()).Value = trans.SwitzerlandGerman;

                        ws.Cell("R" + row.ToString()).Value = trans.MexicoSpanish;
                        ws.Cell("S" + row.ToString()).Value = trans.BrazilPortugese;
                        ws.Cell("T" + row.ToString()).Value = trans.PortugalPortugese;

                        ws.Cell("U" + row.ToString()).Value = trans.ItalyItalian;
                        ws.Cell("V" + row.ToString()).Value = trans.SpainSpainish;
                        ws.Cell("W" + row.ToString()).Value = trans.DenmarkDanish;
                        ws.Cell("X" + row.ToString()).Value = trans.FinlandFinnish;
                        ws.Cell("Y" + row.ToString()).Value = trans.PolandPolish;
                        ws.Cell("Z" + row.ToString()).Value = trans.SwedenSwedish;
                        ws.Cell("AA" + row.ToString()).Value = trans.NorwayNorwegian;
                        ws.Cell("AB" + row.ToString()).Value = trans.ArgentinaSpanish;
                        ws.Cell("AC" + row.ToString()).Value = trans.KoreaKorean;
                        UnlockCellsForEditing(ws, row);

                        if (FIELD_TYPE.Contains(trans.Fieldid))
                        // if(Guid.TryParse(trans.Fieldid, out result))
                        {
                            ChangeCellsbackgroundColor(ws, row);
                        }
                    }


                    string fileName = fileNameWithoutExt + "_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".xlsx";
                    var filePath = $"{RootPath}\\{fileName}";
                    workbook.SaveAs(filePath);
                    DownloadOutPut(filePath);
                }
            }
            catch (Exception ex)
            {
                lblMsg.Text = ex.Message;
                lblMsg.ForeColor = Color.Red;
            }
        }

        private static void SetLockOntheCells(IXLWorksheet ws, int row)
        {
            ws.Cell("D" + row.ToString()).Style.Protection.SetLocked(false);
            ws.Cell("E" + row.ToString()).Style.Protection.SetLocked(false);
            ws.Cell("F" + row.ToString()).Style.Protection.SetLocked(false);
            ws.Cell("G" + row.ToString()).Style.Protection.SetLocked(false);

            ws.Cell("H" + row.ToString()).Style.Protection.SetLocked(false);
            ws.Cell("I" + row.ToString()).Style.Protection.SetLocked(false);
            ws.Cell("J" + row.ToString()).Style.Protection.SetLocked(false);
            ws.Cell("K" + row.ToString()).Style.Protection.SetLocked(false);

            ws.Cell("L" + row.ToString()).Style.Protection.SetLocked(false);
            ws.Cell("M" + row.ToString()).Style.Protection.SetLocked(false);
            ws.Cell("N" + row.ToString()).Style.Protection.SetLocked(false);
            ws.Cell("O" + row.ToString()).Style.Protection.SetLocked(false);
            ws.Cell("P" + row.ToString()).Style.Protection.SetLocked(false);
            ws.Cell("Q" + row.ToString()).Style.Protection.SetLocked(false);
            ws.Cell("R" + row.ToString()).Style.Protection.SetLocked(false);
            ws.Cell("S" + row.ToString()).Style.Protection.SetLocked(false);
            ws.Cell("T" + row.ToString()).Style.Protection.SetLocked(false);
            ws.Cell("U" + row.ToString()).Style.Protection.SetLocked(false);
            ws.Cell("V" + row.ToString()).Style.Protection.SetLocked(false);
            ws.Cell("W" + row.ToString()).Style.Protection.SetLocked(false);
            ws.Cell("X" + row.ToString()).Style.Protection.SetLocked(false);
            ws.Cell("Y" + row.ToString()).Style.Protection.SetLocked(false);
            ws.Cell("Z" + row.ToString()).Style.Protection.SetLocked(false);
            ws.Cell("AA" + row.ToString()).Style.Protection.SetLocked(false);
            ws.Cell("AB" + row.ToString()).Style.Protection.SetLocked(false);
            ws.Cell("AC" + row.ToString()).Style.Protection.SetLocked(false);
        }

        private static void SetColumnWidthAndWrapText(IXLWorksheet ws)
        {
            ws.Rows("1").Style.Fill.BackgroundColor = XLColor.Gray;
            ws.Columns("A").Width = 15;
            ws.Columns("B").Width = 40;
            ws.Columns("C").Width = 20;
            ws.Columns("D:AZ").Width = 40;
            ws.Rows().Style.Alignment.SetWrapText(true);
            //ws.Cell("D" + row.ToString()).Style.Alignment.SetWrapText(true);
            //ws.Cell("E" + row.ToString()).Style.Alignment.SetWrapText(true);
            //ws.Cell("F" + row.ToString()).Style.Alignment.SetWrapText(true);
            //ws.Cell("G" + row.ToString()).Style.Alignment.SetWrapText(true);
        }

        private static void ChangeCellsbackgroundColor(IXLWorksheet ws, int row)
        {
            ws.Cell("D" + row.ToString()).Style.Fill.BackgroundColor = XLColor.Gray;
            ws.Cell("E" + row.ToString()).Style.Fill.BackgroundColor = XLColor.Gray;
            ws.Cell("F" + row.ToString()).Style.Fill.BackgroundColor = XLColor.Gray;
            ws.Cell("G" + row.ToString()).Style.Fill.BackgroundColor = XLColor.Gray;

        }

        private static void UnlockCellsForEditing(IXLWorksheet ws, int row)
        {
            ws.Cell("D" + row.ToString()).Style.Protection.SetLocked(false);
            ws.Cell("E" + row.ToString()).Style.Protection.SetLocked(false);
            ws.Cell("F" + row.ToString()).Style.Protection.SetLocked(false);
            ws.Cell("G" + row.ToString()).Style.Protection.SetLocked(false);

            ws.Cell("H" + row.ToString()).Style.Protection.SetLocked(false);
            ws.Cell("I" + row.ToString()).Style.Protection.SetLocked(false);
            ws.Cell("J" + row.ToString()).Style.Protection.SetLocked(false);
            ws.Cell("K" + row.ToString()).Style.Protection.SetLocked(false);
            ws.Cell("L" + row.ToString()).Style.Protection.SetLocked(false);
            ws.Cell("M" + row.ToString()).Style.Protection.SetLocked(false);
            ws.Cell("N" + row.ToString()).Style.Protection.SetLocked(false);
            ws.Cell("O" + row.ToString()).Style.Protection.SetLocked(false);
            ws.Cell("P" + row.ToString()).Style.Protection.SetLocked(false);
            ws.Cell("Q" + row.ToString()).Style.Protection.SetLocked(false);
            ws.Cell("R" + row.ToString()).Style.Protection.SetLocked(false);
            ws.Cell("S" + row.ToString()).Style.Protection.SetLocked(false);
            ws.Cell("T" + row.ToString()).Style.Protection.SetLocked(false);
            ws.Cell("U" + row.ToString()).Style.Protection.SetLocked(false);
            ws.Cell("V" + row.ToString()).Style.Protection.SetLocked(false);
            ws.Cell("W" + row.ToString()).Style.Protection.SetLocked(false);

            ws.Cell("X" + row.ToString()).Style.Protection.SetLocked(false);
            ws.Cell("Y" + row.ToString()).Style.Protection.SetLocked(false);
            ws.Cell("Z" + row.ToString()).Style.Protection.SetLocked(false);
            ws.Cell("AA" + row.ToString()).Style.Protection.SetLocked(false);
            ws.Cell("AB" + row.ToString()).Style.Protection.SetLocked(false);
            ws.Cell("AC" + row.ToString()).Style.Protection.SetLocked(false);
        }

        public void DownloadOutPut(string filePath)
        {
            Response.ContentType = ContentType;
            Response.AppendHeader("Content-Disposition", "attachment; filename=" + Path.GetFileName(filePath));
            Response.WriteFile(filePath);
            Response.End();
        }

    }
}