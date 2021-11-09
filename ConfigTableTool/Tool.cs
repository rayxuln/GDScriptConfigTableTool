using System;
using System.Collections.Generic;
using System.Text;

using NPOI.SS.UserModel;
using Newtonsoft.Json;
using System.IO;
using NPOI.XSSF.UserModel;

namespace GDScriptConfigTableTool.ConfigTableTool
{
    class Tool
    {
        public Tool()
        {
        }
        
        public void ExportHeadOnlyExcelFile(String dir, String headDefinitionPath)
        {
            List<String> s = new List<String>();
            s.Add(headDefinitionPath);
            ExportHeadOnlyExcelFile(dir, s);
        }

        public void ExportHeadOnlyExcelFile(String dir, List<String> headDefinitionPathList)
        {
            if (!Directory.Exists(dir))
            {
                throw new DirectoryNotFoundException(dir);
            }

            Data data = null;
            String workbookName = "";

            foreach (String hdPath in headDefinitionPathList)
            {
                var fileName = Path.GetFileNameWithoutExtension(hdPath);
                var ss = fileName.Split('_');
                if (ss.Length == 0)
                {
                    throw new InvalidHeadDefinitionFileName($"path: {hdPath}");
                }

                var hd = new HeadDefinition(hdPath);

                workbookName = ss[0];
                var sheetName = ss[0];

                if (ss.Length > 1)
                {
                    sheetName = ss[1];
                }

                if (data == null)
                {
                    data = new Data();
                    data.CreateWorkbook(workbookName);
                } else if (data.GetWorkbookName() != workbookName)
                {
                    throw new HeadDefinitionClassIsNotSame($"{workbookName} != {data.GetWorkbookName()} in path: {hdPath}");
                }

                data.CreateSheet(sheetName);
                data.CreateHead(sheetName, hd);
            }
            
            if (workbookName.Length == 0)
            {
                throw new EmptyHeadDefinitionClass("");
            }

            data.SaveTo(dir, false);
        }

        
        public class InvalidHeadDefinitionFileName : Exception { public InvalidHeadDefinitionFileName(String msg) : base(msg) { } }
        public class HeadDefinitionClassIsNotSame : Exception { public HeadDefinitionClassIsNotSame(String msg) : base(msg) { } }
        public class EmptyHeadDefinitionClass : Exception { public EmptyHeadDefinitionClass(String msg) : base(msg) { } }
    }
}
