using System;
using System.Collections.Generic;
using System.Text;

using NPOI.SS.UserModel;
using Newtonsoft.Json;
using System.IO;
using NPOI.XSSF.UserModel;

namespace GDScriptConfigTableTool.ConfigTableTool
{
    public class Tool
    {
        const String EXCEL_EXT = "xlsx";
        const String HEAD_DEFINITION_EXT = "json";

        public Tool()
        {
        }

        public String GenHeadDefinitionFileName(String name)
        {
            return $"{name}.{HEAD_DEFINITION_EXT}";
        }

        public String GenExcelFileName(String name)
        {
            return $"{name}.{EXCEL_EXT}";
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

            List<HeadDefinition> headDefinitionList = new List<HeadDefinition>();
            foreach (String hdPath in headDefinitionPathList)
            {
                var hd = new HeadDefinition(hdPath);
                headDefinitionList.Add(hd);
            }

            ExportHeadOnlyExcelFile(dir, headDefinitionList);
        }

        public void ExportHeadOnlyExcelFile(String outputDir, List<HeadDefinition> headDefinitionList)
        {
            Data data = null;
            String workbookName = "";

            foreach (HeadDefinition hd in headDefinitionList)
            {
                workbookName = hd.WorkbookName;
                var sheetName = hd.SheetName;

                if (data == null)
                {
                    data = new Data();
                    data.CreateWorkbook(workbookName);
                }
                else if (data.GetWorkbookName() != workbookName)
                {
                    throw new HeadDefinitionClassIsNotSame($"{workbookName} != {data.GetWorkbookName()} in path: {hd.SourcePath}");
                }

                data.CreateSheet(sheetName);
                data.CreateHead(sheetName, hd);
            }

            if (workbookName.Length == 0)
            {
                throw new EmptyHeadDefinitionClass("");
            }

            String path = Path.Combine(outputDir, $"{workbookName}.{EXCEL_EXT}");
            data.SaveTo(path, false);
        }

        public void ExportExcelFile(String outputDir, String hdInputDir)
        {
            List<String> headDefinitionPathList = new List<String>();
            foreach (String path in Directory.EnumerateFiles(hdInputDir))
            {
                if (path.EndsWith($".{HEAD_DEFINITION_EXT}"))
                {
                    headDefinitionPathList.Add(path);
                }
            }
            ExportExcelFile(outputDir, headDefinitionPathList);
        }

        public void ExportExcelFile(String outputDir, List<String> headDefinitionPathList)
        {
            if (headDefinitionPathList.Count == 0)
            {
                return;
            }
            List<HeadDefinition> headDefinitionList = new List<HeadDefinition>();
            foreach (String path in headDefinitionPathList)
            {
                headDefinitionList.Add(new HeadDefinition(path));
            }
            ExportExcelFile(outputDir, headDefinitionList);
        }

        public void ExportExcelFile(String outputDir, List<HeadDefinition> headDefinitionList)
        {
            Dictionary<String, List<HeadDefinition>> hdMap = new Dictionary<string, List<HeadDefinition>>();
            foreach (HeadDefinition hd in headDefinitionList)
            {
                if (!hdMap.ContainsKey(hd.WorkbookName))
                {
                    hdMap[hd.WorkbookName] = new List<HeadDefinition>();
                }
                hdMap[hd.WorkbookName].Add(hd);
            }

            foreach (String workbookName in hdMap.Keys)
            {
                String path = Path.Combine(outputDir, $"{workbookName}.{EXCEL_EXT}");
                if (File.Exists(path))
                {
                    // Reconstruct
                    Data data = new Data();
                    data.CreateWorkbookFromFile(path);
                    data.ReconstructHead(hdMap[workbookName]);
                    data.SaveTo(path, true);
                } else
                {
                    // Export empty with head
                    ExportHeadOnlyExcelFile(outputDir, hdMap[workbookName]);
                }

            }
        }

        public void ExportGDScript(String outputDir, String templateCode, String excelsDir, String defsDir)
        {
            Dictionary<String, Data> workbookNameDataMap = new Dictionary<string, Data>();
            Dictionary<String, List<HeadDefinition>> workbookNameHeadDifinitionListMap = new Dictionary<string, List<HeadDefinition>>();

            foreach (var fileName in Directory.EnumerateFiles(excelsDir))
            {
                if (fileName.EndsWith($".{EXCEL_EXT}"))
                {
                    var data = new Data();
                    data.CreateWorkbookFromFile(Path.Combine(excelsDir, fileName));
                    workbookNameDataMap[data.GetWorkbookName()] = data;
                }
            }

            foreach (var fileName in Directory.EnumerateFiles(defsDir))
            {
                if (fileName.EndsWith($".{HEAD_DEFINITION_EXT}"))
                {
                    var headDefinition = new HeadDefinition(Path.Combine(defsDir, fileName));
                    if (!workbookNameHeadDifinitionListMap.ContainsKey(headDefinition.WorkbookName))
                    {
                        workbookNameHeadDifinitionListMap[headDefinition.WorkbookName] = new List<HeadDefinition>();
                    }
                    workbookNameHeadDifinitionListMap[headDefinition.WorkbookName].Add(headDefinition);
                }
            }

            foreach (var workbookName in workbookNameDataMap.Keys)
            {
                if (workbookNameHeadDifinitionListMap.ContainsKey(workbookName))
                {
                    var data = workbookNameDataMap[workbookName];
                    var headDefinitionList = workbookNameHeadDifinitionListMap[workbookName];
                    ExportGDScript(outputDir, templateCode, data, headDefinitionList);
                }
            }
        }

        public void ExportGDScript(String outputDir, String templateCode, Data data, List<HeadDefinition> headDefinitionList)
        {
            foreach (HeadDefinition headDefinition in headDefinitionList)
            {
                ExportGDScript(outputDir, templateCode, data, headDefinition);
            }
        }

        public void ExportGDScript(String outputDir, String templateCode, Data data, HeadDefinition headDefinition)
        {
            Script script = new Script(templateCode);
            script.Create(data, headDefinition);
            script.SaveTo(Path.Combine(outputDir, script.FileName));
        }
        
        public class InvalidHeadDefinitionFileName : Exception { public InvalidHeadDefinitionFileName(String msg) : base(msg) { } }
        public class HeadDefinitionClassIsNotSame : Exception { public HeadDefinitionClassIsNotSame(String msg) : base(msg) { } }
        public class EmptyHeadDefinitionClass : Exception { public EmptyHeadDefinitionClass(String msg) : base(msg) { } }
    }
}
