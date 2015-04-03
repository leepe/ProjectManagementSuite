using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using ProjectManagementSuite.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;

namespace ProjectManagementSuite.CSharpLogic
{
    public class GenerateWorkbook
    {
        //--------------------------------------------------------------
        // create project workbook to standard format
        //--------------------------------------------------------------
        public static void generateProjectWorkBook(newProject nph, string f0)
        {

            String[] titles = { "Item", "Description", "Whse", "Item-Whse-State Status", "BUShort", "End" };
            String[] verticalTitles = { "Project Number", "Project Name", "Description", "Manual Fcst", "Movex CO Number" };
            // location of template file with VBA module
            string ftempl = HttpContext.Current.Server.MapPath("~/CSharpLogic") + "/ManualFcstLogic2.xls";
            //--------------------------------------------------------
            // create workbook and define sheets
            //--------------------------------------------------------
            using (FileStream ftemplate = new FileStream(ftempl, FileMode.Open, FileAccess.ReadWrite))
            {
                IWorkbook workbook = new HSSFWorkbook(ftemplate);
                //---------------------------------------------------------
                // get dictionaries
                //---------------------------------------------------------
                Dictionary<string, ICellStyle> sd = createStyles(workbook);
                //---------------------------------------------------------
                // produce a safe project sheet name
                //---------------------------------------------------------
                string tabName = Regex.Replace(nph.projectName.Substring(0, Math.Min(25, nph.projectName.Length)).Trim(), @"[^a-zA-Z0-9 -]", "");
                //ISheet sheet = workbook.CreateSheet("Proj~" + tabName);
                ISheet sheet = workbook.GetSheetAt(0);
                workbook.SetSheetName(0, "Proj~" + tabName);
                //---------------------------------------------------------
                // turn on sheet print set up specific things
                //---------------------------------------------------------
                sheet.DisplayGridlines = true;
                sheet.IsPrintGridlines = true;
                sheet.HorizontallyCenter = true;
                sheet.PrintSetup.Landscape = true;
                sheet.PrintSetup.PaperSize = (short)PaperSize.A4;
                sheet.Autobreaks = true;
                sheet.PrintSetup.FitHeight = (short)1;
                sheet.PrintSetup.FitWidth = (short)1;
                //---------------------------------------------------------
                // set up header row
                //---------------------------------------------------------
                IRow headerRow = sheet.CreateRow(8); // this is zero based row numbering
                headerRow.HeightInPoints = 19.00f;
                for (int i = 0; i < titles.Length; i++)
                {
                    ICell cell = headerRow.CreateCell(i + 1);
                    cell.SetCellValue(titles[i]);
                    cell.CellStyle = sd["cell_normal"];
                }
                //---------------------------------------------------------
                // columns for 24 months starting from today
                //---------------------------------------------------------
                DateTime d0 = DateTime.Now;
                DateTime startOfMonth = new DateTime(d0.Year, d0.Month, 1);
                for (int i = 0; i < 23; i++)
                {
                    ICell cell = headerRow.CreateCell(titles.Length + i);
                    cell.SetCellValue(startOfMonth.AddMonths(i));
                    cell.CellStyle = sd["cell_normal_date"];
                }
                //----------------------------------------------------------
                //freeze the first row
                //----------------------------------------------------------
                sheet.CreateFreezePane(6, 9);
                // output header rows
                IRow row = sheet.CreateRow(0);
                ICell cell0 = row.CreateCell(0);
                cell0.SetCellValue("Project Data");
                cell0.CellStyle = sd["cell_normal"];
                ICell cell1 = row.CreateCell(1);
                cell1.SetCellValue("Project ID");
                cell1.CellStyle = sd["cell_normal"];
                ICell cell2 = row.CreateCell(2);
                cell2.SetCellValue("***New***");
                cell2.CellStyle = sd["cell_green"];
                // add in reference to state on mastercontrolsheet
                ICell cell4 = row.CreateCell(4);
				cell4.SetCellFormula("MasterControlSheet!B1");
				cell4.CellStyle = sd["hState"];
                // roll down and output column values
                for (int i = 0; i < verticalTitles.Length; i++)
                {
                    IRow vtitleCol = sheet.CreateRow(i + 1);
                    ICell cell = vtitleCol.CreateCell(1);
                    ICell cell_1 = vtitleCol.CreateCell(2);
                    if (i == 0) cell_1.SetCellValue(nph.projectNumber.Trim());
                    else if (i == 1) cell_1.SetCellValue(Regex.Replace(nph.projectName.Trim(), @"[^a-zA-Z0-9 -]", ""));
                    else if (i == 2) cell_1.SetCellValue(Regex.Replace(nph.projectType.Trim(), @"[^a-zA-Z0-9 -]", ""));
                    else if (i == 3 && nph.mvxorders.Count > 0) cell_1.SetCellValue("No");
                    else if (i == 3 && nph.mvxorders.Count == 0) cell_1.SetCellValue("Yes");
                    else if (i == 4)
                    {
                        if (nph.mvxorders.Count > 0)
                        {
                            StringBuilder sb = new StringBuilder();
                            for (int j = 0; j < nph.mvxorders.Count; j++)
                            {
                                sb.Append(nph.mvxorders[j].order);
                                //sb.Append(Environment.NewLine);
                                sb.Append("\n");
                            }
                            cell_1.SetCellValue(sb.ToString());
                        }
                    }
                    cell.SetCellValue(verticalTitles[i]);
                    cell.CellStyle = sd["cell_normal"];
                    if (i != 3) cell_1.CellStyle = sd["cell_green"];
                    cell_1.CellStyle.Alignment = HorizontalAlignment.Left;
                    cell_1.CellStyle = sd["cell_green"];
                    cell_1.CellStyle.Alignment = HorizontalAlignment.Left;
                }
                // roll down the rest of the sheet and format
                for (int i = 9; i < 100; i++)
                {
                    IRow formatRow = sheet.CreateRow(i);
                    for (int j = 1; j < 29; j++)
                    {
                        ICell cell = formatRow.CreateCell(j);
                        if (j == 1 || j == 3 || (j >= 6 && j <= 28))
                        {
                            cell.CellStyle = sd["cell_green"];
                            cell.CellStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("General");
                        }
                    }
                }
                //set column widths, the width is measured in units of 1/256th of a character width
                sheet.SetColumnWidth(0, 256 * 16);
                sheet.SetColumnWidth(1, 256 * 25);
                sheet.SetColumnWidth(2, 256 * 60);
                sheet.SetColumnWidth(3, 256 * 9);
                sheet.SetColumnWidth(4, 256 * 30);
                sheet.SetColumnWidth(5, 256 * 10);
                sheet.SetZoom(5, 6);
                //------------------------------------------------------------------------
                // now add a second sheet called MasterControlSheet with person's state
                //------------------------------------------------------------------------
                //ISheet sheet2 = workbook.CreateSheet("MasterControlSheet");
                ISheet sheet2 = workbook.GetSheetAt(1);
                IRow topRow = sheet2.CreateRow(0);
                ICell cell_f = topRow.CreateCell(0);
                cell_f.SetCellValue("State");
                ICell cell_g = topRow.CreateCell(1);
                cell_g.SetCellValue(nph.projectState);
                //------------------------------------------------------------------------
                FileStream sw = File.Create(f0);
                workbook.Write(sw);
                sw.Close();
            }
        }

        //--------------------------------------------------------------
        // create dictionary of styles to be used in xlsx creation
        //--------------------------------------------------------------
        public static Dictionary<string, ICellStyle> createStyles(IWorkbook wb)
        {
            // create dictionary
            Dictionary<string, ICellStyle> styleDictionary = new Dictionary<string, ICellStyle>();
            // define styles to apply to output data
            ICellStyle icstyle1 = wb.CreateCellStyle();
            // fonts
            IFont dfont = wb.CreateFont();
            dfont.Boldweight = (short)FontBoldWeight.Bold;
            dfont.FontName = "Calibri";
            dfont.IsItalic = true;
            dfont.Color = HSSFColor.RoyalBlue.Index;
            dfont.FontHeightInPoints = 12;
            // red warning fonts
            IFont wfont = wb.CreateFont();
            wfont.Boldweight = (short)FontBoldWeight.Bold;
            wfont.IsItalic = true;
            wfont.FontName = "Calibri";
            wfont.Color = HSSFColor.Maroon.Index;
            wfont.FontHeightInPoints = 12;			
            //
            //---- Cell_Normal ------
            //
            icstyle1.Alignment = HorizontalAlignment.Left;
            icstyle1.WrapText = true;
            icstyle1.DataFormat = HSSFDataFormat.GetBuiltinFormat("@");
            styleDictionary.Add("cell_normal", icstyle1);
            //
            //---- Cell_Normal_Date ------
            //
            icstyle1.Alignment = HorizontalAlignment.Left;
            icstyle1.WrapText = true;
            icstyle1.DataFormat = HSSFDataFormat.GetBuiltinFormat("mmm-yy");
            styleDictionary.Add("cell_normal_date", icstyle1);
            //
            //---- Cell_Green ------------
            icstyle1 = createBorderedStyle(wb);
            icstyle1.Alignment = HorizontalAlignment.Left;
            icstyle1.FillForegroundColor = HSSFColor.LightGreen.Index;
            icstyle1.FillPattern = FillPattern.SolidForeground;
            icstyle1.DataFormat = HSSFDataFormat.GetBuiltinFormat("@");
            icstyle1.WrapText = true;
            styleDictionary.Add("cell_green", icstyle1);
            //
            //---- Header Style -----
            //
            IFont headerFont = wb.CreateFont();
            headerFont.Boldweight = (short)FontBoldWeight.Bold;
            headerFont.FontName = "Tahoma";
            icstyle1.Alignment = HorizontalAlignment.Center;
            icstyle1.DataFormat = HSSFDataFormat.GetBuiltinFormat("@");
            icstyle1.SetFont(headerFont);
            styleDictionary.Add("header", icstyle1);
            //
            //----- Header_Date Format -------
            //
            icstyle1.DataFormat = HSSFDataFormat.GetBuiltinFormat("mmm-yy");
            styleDictionary.Add("header_date", icstyle1);
            //
            //
            //
            //---- Cell_OK_Date ------
            //
            ICellStyle okstyle = wb.CreateCellStyle();
            okstyle.Alignment = HorizontalAlignment.Left;
            okstyle.WrapText = true;
            okstyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("mmm-yy");
            okstyle.SetFont(dfont);
            okstyle.FillForegroundColor = HSSFColor.LightYellow.Index;
            okstyle.FillPattern = FillPattern.SolidForeground;            // add to dictionary
            styleDictionary.Add("cell_ok_date", okstyle);
            //
            //---- Cell_Late_Date ------
            //
            //
            ICellStyle lateStyle = wb.CreateCellStyle();
            lateStyle.Alignment = HorizontalAlignment.Left;
            lateStyle.WrapText = true;
            lateStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("mmm-yy");
            lateStyle.SetFont(wfont);
            lateStyle.FillForegroundColor = HSSFColor.Grey25Percent.Index;
            lateStyle.FillPattern = FillPattern.SolidForeground;
            styleDictionary.Add("cell_late_date", lateStyle);
            //
            //
            //------ Late_Qty ------------
            //
            //
            ICellStyle lateFmt = wb.CreateCellStyle();
            lateFmt = createBorderedStyle(wb);
            lateFmt.SetFont(wfont);
            lateFmt.Alignment = HorizontalAlignment.Left;
            lateFmt.FillForegroundColor = HSSFColor.LightGreen.Index;
            lateFmt.FillPattern = FillPattern.SolidForeground;
            lateFmt.DataFormat = HSSFDataFormat.GetBuiltinFormat("@");
            lateFmt.WrapText = true;
            //
            styleDictionary.Add("Late_Qty", lateFmt);
            //
            //--------OnTime_Qty -----------
            //
            //
            ICellStyle ontimeFmt = wb.CreateCellStyle();
            ontimeFmt = createBorderedStyle(wb);
            ontimeFmt.SetFont(dfont);
            ontimeFmt.Alignment = HorizontalAlignment.Left;
            ontimeFmt.FillForegroundColor = HSSFColor.LightGreen.Index;
            ontimeFmt.FillPattern = FillPattern.SolidForeground;
            ontimeFmt.DataFormat = HSSFDataFormat.GetBuiltinFormat("@");
            ontimeFmt.WrapText = true;
            //
            styleDictionary.Add("OK_Qty", ontimeFmt);
			//
			//-------Highlight_State --------------
			//
			//
            ICellStyle highState = wb.CreateCellStyle();
            highState.FillForegroundColor = HSSFColor.LightGreen.Index; 
			highState.FillPattern = FillPattern.SolidForeground;
            highState.Alignment = HorizontalAlignment.Center; 
			IFont hsfnt = wb.CreateFont();
            //
            hsfnt.Boldweight = (short)FontBoldWeight.Bold; 
			hsfnt.FontName = "Calibri"; 
			hsfnt.IsItalic = true;
            hsfnt.Color = HSSFColor.Green.Index; 
			hsfnt.FontHeightInPoints = 12;
            highState.SetFont(hsfnt);
            //
			styleDictionary.Add("hState",highState);
			//
            return styleDictionary;
        }
        //-----------------------------------------------------------------
        // create bordered style for cells
        //-----------------------------------------------------------------
        private static ICellStyle createBorderedStyle(IWorkbook wb)
        {
            ICellStyle style = wb.CreateCellStyle();
            style.BorderRight = BorderStyle.Thin;
            style.BorderLeft = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;
            style.BorderBottom = BorderStyle.Thin;
            style.RightBorderColor = HSSFColor.Grey25Percent.Index;
            style.LeftBorderColor = HSSFColor.Grey25Percent.Index;
            style.TopBorderColor = HSSFColor.Grey25Percent.Index;
            style.BottomBorderColor = HSSFColor.Grey25Percent.Index;
            return style;
        }

    }
}