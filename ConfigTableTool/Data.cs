using System;
using System.Collections.Generic;
using System.Text;

using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;
using System.Drawing;

namespace GDScriptConfigTableTool.ConfigTableTool
{
    public class Data
    {
        String workbookName;
        XSSFWorkbook workbook;

        const int NAME_ROW = 0;
        const int ID_ROW = 1;
        const int TYPE_ROW = 2;
        const int LOCK_ROW = 3;

        int defaultWidth = 230 * 24;
        HorizontalAlignment defaultAlignment = HorizontalAlignment.Center;
        VerticalAlignment defaultVerticalAlignment = VerticalAlignment.Center;
        int nameDefaultFontSize = 11;
        float nameDefaultHeight = 24;
        Color nameDefaultFColor = Color.Black;
        Color nameDefaultBColor = Color.FromArgb(unchecked((int)0xFFA9D08E));
        XSSFCellStyle theNameStyle = null;

        int idDefaultFontSize = 12;
        float idDefaultHeight = 20;
        Color idDefaultFColor = Color.Black;
        Color idDefaultBColor = Color.FromArgb(unchecked((int)0xFFA9D08E));
        XSSFCellStyle theIdStyle = null;

        int typeDefaultFontSize = 10;
        float typeDefaultHeight = 14;
        Color typeDefaultFColor = Color.White;
        Color typeDefaultBColor = Color.FromArgb(unchecked((int)0xFF525252));
        XSSFCellStyle theTypeStyle = null;

        Color defaultBColor = Color.FromArgb(unchecked((int)0xFFE7E6E6));
        Color defaultFColor = Color.Black;
        XSSFCellStyle theDefaultStyle = null;


        public Data()
        {
        }

        public String GetWorkbookName()
        {
            return workbookName;
        }

        public void CreateWorkbook(String name)
        {
            workbookName = name;
            workbook = new XSSFWorkbook();
        }

        public void CreateWorkbookFromFile(String path)
        {
            workbookName = Path.GetFileNameWithoutExtension(path);
            var file = File.OpenRead(path);
            workbook = new XSSFWorkbook(file);
            file.Close();
        }

        public void CreateSheet(String name)
        {
            workbook.CreateSheet(name);
        }

        public byte[] ColorToBytes(Color color)
        {
            return new byte[]
            {
                color.R,
                color.G,
                color.B,
            };
        }

        XSSFCellStyle GenDefaultColumnStyle()
        {
            if (theDefaultStyle != null) return theDefaultStyle;
            var defaultColumnStyle = (XSSFCellStyle)workbook.CreateCellStyle();
            defaultColumnStyle.VerticalAlignment = VerticalAlignment.Center;
            defaultColumnStyle.BorderTop = BorderStyle.Thin;
            defaultColumnStyle.BorderBottom = BorderStyle.Thin;
            defaultColumnStyle.BorderLeft = BorderStyle.Thin;
            defaultColumnStyle.BorderRight = BorderStyle.Thin;
            defaultColumnStyle.TopBorderColor = 0;
            defaultColumnStyle.BottomBorderColor = 0;
            defaultColumnStyle.LeftBorderColor = 0;
            defaultColumnStyle.RightBorderColor = 0;
            var defaultFont = (XSSFFont)workbook.CreateFont();
            defaultFont.FontHeightInPoints = 11;
            defaultFont.SetColor(new XSSFColor(defaultFColor));
            defaultColumnStyle.SetFont(defaultFont);
            defaultColumnStyle.SetFillForegroundColor(new XSSFColor(defaultBColor));
            defaultColumnStyle.FillPattern = FillPattern.SolidForeground;
            return theDefaultStyle = defaultColumnStyle;
        }

        XSSFCellStyle GenNameStyle()
        {
            if (theNameStyle != null) return theNameStyle;
            var nameStyle = (XSSFCellStyle)workbook.CreateCellStyle();
            nameStyle.Alignment = defaultAlignment;
            nameStyle.VerticalAlignment = defaultVerticalAlignment;
            nameStyle.BorderTop = BorderStyle.Thin;
            nameStyle.BorderBottom = BorderStyle.Thin;
            nameStyle.BorderLeft = BorderStyle.Thin;
            nameStyle.BorderRight = BorderStyle.Thin;
            nameStyle.TopBorderColor = new XSSFColor(Color.Black).Index;
            nameStyle.BottomBorderColor = new XSSFColor(Color.Black).Index;
            nameStyle.LeftBorderColor = new XSSFColor(Color.Black).Index;
            nameStyle.RightBorderColor = new XSSFColor(Color.Black).Index;
            var nameFont = (XSSFFont)workbook.CreateFont();
            nameFont.FontHeightInPoints = nameDefaultFontSize;
            nameFont.SetColor(new XSSFColor(nameDefaultFColor));
            nameStyle.SetFont(nameFont);
            nameStyle.SetFillForegroundColor(new XSSFColor(nameDefaultBColor));
            nameStyle.FillPattern = FillPattern.SolidForeground;
            return theNameStyle = nameStyle;
        }

        XSSFCellStyle GenIdStyle()
        {
            if (theIdStyle != null) return theIdStyle;
            var idStyle = (XSSFCellStyle)workbook.CreateCellStyle();
            idStyle.Alignment = defaultAlignment;
            idStyle.VerticalAlignment = defaultVerticalAlignment;
            idStyle.BorderTop = BorderStyle.Thin;
            idStyle.BorderBottom = BorderStyle.Thin;
            idStyle.BorderLeft = BorderStyle.Thin;
            idStyle.BorderRight = BorderStyle.Thin;
            idStyle.TopBorderColor = new XSSFColor(Color.Black).Index;
            idStyle.BottomBorderColor = new XSSFColor(Color.Black).Index;
            idStyle.LeftBorderColor = new XSSFColor(Color.Black).Index;
            idStyle.RightBorderColor = new XSSFColor(Color.Black).Index;
            var idFont = (XSSFFont)workbook.CreateFont();
            idFont.FontHeightInPoints = idDefaultFontSize;
            idFont.SetColor(new XSSFColor(idDefaultFColor));
            idStyle.SetFont(idFont);
            idStyle.SetFillForegroundColor(new XSSFColor(idDefaultBColor));
            idStyle.FillPattern = FillPattern.SolidForeground;
            return theIdStyle = idStyle;
        }

        XSSFCellStyle GenTypeStyle()
        {
            if (theTypeStyle != null) return theTypeStyle;
            var typeStyle = (XSSFCellStyle)workbook.CreateCellStyle();
            typeStyle.Alignment = defaultAlignment;
            typeStyle.VerticalAlignment = defaultVerticalAlignment;
            typeStyle.BorderTop = BorderStyle.Thin;
            typeStyle.BorderBottom = BorderStyle.Thin;
            typeStyle.BorderLeft = BorderStyle.Thin;
            typeStyle.BorderRight = BorderStyle.Thin;
            typeStyle.TopBorderColor = new XSSFColor(Color.Black).Index;
            typeStyle.BottomBorderColor = new XSSFColor(Color.Black).Index;
            typeStyle.LeftBorderColor = new XSSFColor(Color.Black).Index;
            typeStyle.RightBorderColor = new XSSFColor(Color.Black).Index;
            var typeFont = (XSSFFont)workbook.CreateFont();
            typeFont.FontHeightInPoints = typeDefaultFontSize;
            typeFont.SetColor(new XSSFColor(typeDefaultFColor));
            typeStyle.SetFont(typeFont);
            typeStyle.SetFillForegroundColor(new XSSFColor(typeDefaultBColor));
            typeStyle.FillPattern = FillPattern.SolidForeground;
            return theTypeStyle = typeStyle;
        }

        void CreateHeadAt(XSSFSheet sheet, int column, DefinitionType def)
        {
            var nameRow = sheet.GetRow(NAME_ROW);
            if (nameRow == null)
            {
                nameRow = sheet.CreateRow(NAME_ROW);
                nameRow.HeightInPoints = nameDefaultHeight;
            }

            var idRow = sheet.GetRow(ID_ROW);
            if (idRow == null)
            {
                idRow = sheet.CreateRow(ID_ROW);
                idRow.HeightInPoints = idDefaultHeight;
            }

            var typeRow = sheet.GetRow(TYPE_ROW);
            if (typeRow == null)
            {
                typeRow = sheet.CreateRow(TYPE_ROW);
                typeRow.HeightInPoints = typeDefaultHeight;
            }

            ICell cell;
            cell = nameRow.CreateCell(column);
            cell.SetCellValue(def.name);
            cell.CellStyle = GenNameStyle();

            cell = idRow.CreateCell(column);
            cell.SetCellValue(def.id);
            cell.CellStyle = GenIdStyle();

            cell = typeRow.CreateCell(column);
            cell.SetCellValue(def.type.ToString());
            cell.CellStyle = GenTypeStyle();
        }
        public void CreateHead(String sheetName, HeadDefinition headDefinition)
        {
            var sheet = workbook.GetSheet(sheetName);
            var nameRow = sheet.CreateRow(NAME_ROW);
            nameRow.HeightInPoints = nameDefaultHeight;

            var idRow = sheet.CreateRow(ID_ROW);
            idRow.HeightInPoints = idDefaultHeight;

            var typeRow = sheet.CreateRow(TYPE_ROW);
            typeRow.HeightInPoints = typeDefaultHeight;

            var defaultColumnStyle = GenDefaultColumnStyle();

            for (int j=0; j<headDefinition.GetDefinitionCount(); ++j)
            {
                sheet.SetColumnWidth(j, defaultWidth);
                sheet.SetDefaultColumnStyle(j, defaultColumnStyle);
            }
            sheet.CreateFreezePane(0, LOCK_ROW);


            var nameStyle = GenNameStyle();

            var idStyle = GenIdStyle();

            var typeStyle = GenTypeStyle();

            int i = 0;
            foreach (DefinitionType def in headDefinition)
            {
                ICell cell;
                cell = nameRow.CreateCell(i);
                cell.SetCellValue(def.name);
                cell.CellStyle = nameStyle;

                cell = idRow.CreateCell(i);
                cell.SetCellValue(def.id);
                cell.CellStyle = idStyle;

                cell = typeRow.CreateCell(i);
                cell.SetCellValue(def.type.ToString());
                cell.CellStyle = typeStyle;

                i += 1;
            }
        }

        public void ReconstructHead(List<HeadDefinition> headDefinitionList)
        {
            foreach (HeadDefinition hd in headDefinitionList)
            {
                ReconstructHead(hd);
            }
        }
        public void ReconstructHead(HeadDefinition headDefinition)
        {
            String sheetName = headDefinition.SheetName;
            var sheet = (XSSFSheet)workbook.GetSheet(sheetName);
            if (sheet == null)
            {
                CreateSheet(sheetName);
                CreateHead(sheetName, headDefinition);
                return;
            }

            if (sheet.FirstRowNum != 0 && sheet.LastRowNum > sheet.FirstRowNum)
            {
                throw new RowNotStartAtZero($"in original sheet: ${headDefinition.SheetName}, workbook: ${headDefinition.WorkbookName}");
            }
            int rowCount = sheet.LastRowNum;
            if (rowCount == 0)
            {
                CreateHead(sheetName, headDefinition);
                return;
            }

            var tempSheetName = $"__temp__{sheetName}";
            workbook.SetSheetName(workbook.GetSheetIndex(sheet), tempSheetName);
            var tempSheet = sheet;
            sheet = (XSSFSheet)workbook.CreateSheet(sheetName);

            CreateHead(sheetName, headDefinition);

            var longestCellNum = 0;
            IRow firstRow = null;
            for (int i=tempSheet.FirstRowNum; i<=tempSheet.LastRowNum; ++i)
            {
                var row = tempSheet.GetRow(i);
                if (row == null) continue;
                if (row.LastCellNum > longestCellNum)
                {
                    longestCellNum = row.LastCellNum;
                    firstRow = row;
                }
            }
            

            Dictionary<String, int> idColumnMap = new Dictionary<String, int>();
            foreach (DefinitionType def in headDefinition)
            {
                for (int i=0; i<firstRow.LastCellNum; ++i)
                {
                    var row = (XSSFRow)tempSheet.GetRow(ID_ROW);
                    var cell = (XSSFCell)row.GetCell(i);
                    if (cell == null) continue;
                    var id = cell.StringCellValue;
                    if (id == def.id)
                    {
                        idColumnMap[id] = i;
                        break;
                    }
                }
            }

            Dictionary<int, bool> columnUsedMap = new Dictionary<int, bool>();
            int currentColumn = 0;
            foreach (DefinitionType def in headDefinition)
            {
                // if exist, then copy
                if (idColumnMap.ContainsKey(def.id))
                {
                    var column = idColumnMap[def.id];
                    columnUsedMap[column] = true;
                    CopyColumn(tempSheet, column, sheet, currentColumn);
                }
                else // otherwise, create a new column
                {
                    CreateHeadAt(sheet, currentColumn, def);
                }

                currentColumn += 1;
            }
            // copy the remainings columns
            for (int i = 0; i < firstRow.LastCellNum; ++i)
            {
                if (!columnUsedMap.ContainsKey(i) || columnUsedMap[i] != true)
                {
                    CopyColumn(tempSheet, i, sheet, currentColumn);
                    currentColumn += 1;
                }
            }
            // setup default style
            for (int j = 0; j < headDefinition.GetDefinitionCount(); ++j)
            {
                sheet.SetColumnWidth(j, defaultWidth);
                sheet.SetDefaultColumnStyle(j, GenDefaultColumnStyle());
            }
            sheet.CreateFreezePane(0, LOCK_ROW);

            // delete old sheet
            workbook.RemoveSheetAt(workbook.GetSheetIndex(tempSheet));
        }

        private void CopyColumn(XSSFSheet srcSheet, int srcColumn, XSSFSheet dstSheet, int dstColumn)
        {
            for (int i=srcSheet.FirstRowNum; i<=srcSheet.LastRowNum; ++i)
            {
                var srcRow = (XSSFRow)srcSheet.GetRow(i);
                if (srcRow == null) continue;
                var dstRow = (XSSFRow)dstSheet.GetRow(i);
                if (dstRow == null) dstRow = (XSSFRow)dstSheet.CreateRow(i);

                var srcCell = (XSSFCell)srcRow.GetCell(srcColumn);
                if (srcCell == null) continue;
                var dstCell = (XSSFCell)dstRow.GetCell(dstColumn);
                if (dstCell == null) dstCell = (XSSFCell)dstRow.CreateCell(dstColumn);

                var policy = (new CellCopyPolicy.Builder())
                    .CellFormula(true)
                    .CellStyle(true)
                    .CellValue(true)
                    .CopyHyperlink(true)
                    .RowHeight(true)
                    .Build();
                dstCell.CopyCellFrom(srcCell, policy);
            }
        }

        public void SaveTo(String path, bool overrided=false)
        {
            if (File.Exists(path))
            {
                if (overrided)
                {
                    File.Delete(path);
                } else
                {
                    throw new FileExistException($"{path} is already there!");
                }
            }
            var f = File.OpenWrite(path);
            workbook.Write(f);
        }

        public class FileExistException : Exception { public FileExistException(String msg) : base(msg) { } }
        public class RowNotStartAtZero : Exception { public RowNotStartAtZero(String msg) : base(msg) { } }
        public class EmptyIdentify : Exception { public EmptyIdentify(String msg) : base(msg) { } }
    }
}
