using System;
using System.Collections.Generic;
using System.Text;

using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;
using System.Drawing;

namespace GDScriptConfigTableTool.ConfigTableTool
{
    class Data
    {
        String workbookName;
        XSSFWorkbook workbook;

        const int NAME_ROW = 0;
        const int ID_ROW = 1;
        const int TYPE_ROW = 2;

        int defaultWidth = 230 * 24;
        HorizontalAlignment defaultAlignment = HorizontalAlignment.Center;
        VerticalAlignment defaultVerticalAlignment = VerticalAlignment.Center;
        int nameDefaultFontSize = 11;
        float nameDefaultHeight = 24;
        Color nameDefaultFColor = Color.Black;
        Color nameDefaultBColor = Color.FromArgb(unchecked((int)0xFFA9D08E));

        int idDefaultFontSize = 12;
        float idDefaultHeight = 20;
        Color idDefaultFColor = Color.Black;
        Color idDefaultBColor = Color.FromArgb(unchecked((int)0xFFA9D08E));

        int typeDefaultFontSize = 10;
        float typeDefaultHeight = 14;
        Color typeDefaultFColor = Color.White;
        Color typeDefaultBColor = Color.FromArgb(unchecked((int)0xFF525252));

        Color defaultBColor = Color.FromArgb(unchecked((int)0xFFE7E6E6));
        Color defaultFColor = Color.Black;



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

        public void CreateHead(String sheetName, HeadDefinition headDefinition)
        {
            var sheet = workbook.GetSheet(sheetName);
            var nameRow = sheet.CreateRow(NAME_ROW);
            nameRow.HeightInPoints = nameDefaultHeight;

            var idRow = sheet.CreateRow(ID_ROW);
            idRow.HeightInPoints = idDefaultHeight;

            var typeRow = sheet.CreateRow(TYPE_ROW);
            typeRow.HeightInPoints = typeDefaultHeight;

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

            for (int j=0; j<headDefinition.GetDefinitionCount(); ++j)
            {
                sheet.SetColumnWidth(j, defaultWidth);
                sheet.SetDefaultColumnStyle(j, defaultColumnStyle);
            }
            

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

        public void SaveTo(String dir, bool overrided=false)
        {
            var path = Path.Combine(dir, $"{workbookName}.xlsx");
            if (!overrided && File.Exists(path))
            {
                throw new FileExistException($"{path} is already there!");
            }
            var f = File.OpenWrite(path);
            workbook.Write(f);
        }

        public class FileExistException : Exception { public FileExistException(String msg) : base(msg) { } }
    }
}
