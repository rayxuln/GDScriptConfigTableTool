using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace GDScriptConfigTableTool.ConfigTableTool
{
    public class Script
    {
        String codeTemplate;

        String className;
        public String FileName
        {
            get
            {
                return $"{className}.gd";
            }
        }
        String sourceCode;
        public Script(String template)
        {
            codeTemplate = template;
        }

        void ClearCode()
        {
            sourceCode = "";
        }

        void GenCode(Dictionary<String, String> paramMap)
        {
            StringBuilder builder = new StringBuilder(codeTemplate);
            foreach (var k in paramMap.Keys)
            {
                builder.Replace($"{{{k}}}", paramMap[k]);
            }
            sourceCode = builder.ToString();
        }

        String GenIndent(int num)
        {
            StringBuilder builder = new StringBuilder();
            for (int i=0; i<num; ++i)
            {
                builder.Append('\t');
            }
            return builder.ToString();
        }

        String GenFieldDecList(HeadDefinition headDefinition)
        {
            const int INDENT = 1;
            const String TEMP = "{0}var {1}";
            StringBuilder builder = new StringBuilder();
            String indent = GenIndent(INDENT);
            foreach (DefinitionType def in headDefinition)
            {
                builder.AppendLine(String.Format(TEMP, indent, def.id));
            }
            return builder.ToString();
        }

        String GenDataList(Data data, HeadDefinition headDefinition)
        {
            const int INDENT = 3;
            const String TEMP = "{0}DataType.new({{{1}}}),";
            const String K_V_TEMP = "\"{0}\": \"{1}\",";
            StringBuilder builder = new StringBuilder();
            StringBuilder kvBuilder = new StringBuilder();
            String indent = GenIndent(INDENT);
            for (int i=0; i<data.GetRowCount(headDefinition.SheetName); ++i)
            {
                if (!data.IsRowExist(headDefinition.SheetName, i)) continue;
                kvBuilder.Clear();
                foreach (DefinitionType def in headDefinition)
                {
                    String value = data.GetValueAt(headDefinition.SheetName, i, def.id);
                    if (value == null) continue;
                    kvBuilder.Append(String.Format(K_V_TEMP, def.id, value));
                }
                if (kvBuilder.Length == 0) continue;
                builder.AppendLine(String.Format(TEMP, indent, kvBuilder.ToString()));
            }
            return builder.ToString();
        }

        String GenDataHeadDef(HeadDefinition headDefinition)
        {
            const int INDENT = 2;
            const String TEMP = "{0}\"{1}\",";
            StringBuilder builder = new StringBuilder();
            String indent = GenIndent(INDENT);
            foreach (DefinitionType def in headDefinition)
            {
                builder.AppendLine(String.Format(TEMP, indent, def.id));
            }
            return builder.ToString();
        }

        String GenByFieldFuncList(HeadDefinition headDefinition)
        {
            const int INDENT = 1;
            const String TEMP = "func by_{1}(v) -> DataType:\n{0}return by(\"{1}\", v)\n";
            StringBuilder builder = new StringBuilder();
            String indent = GenIndent(INDENT);
            foreach (DefinitionType def in headDefinition)
            {
                builder.AppendLine(String.Format(TEMP, indent, def.id));
            }
            return builder.ToString();
        }

        public void Create(Data data, HeadDefinition headDefinition)
        {
            className = $"{headDefinition.WorkbookName}_{headDefinition.SheetName}";
            ClearCode();

            Dictionary<String, String> genMap = new Dictionary<string, string>();
            genMap["FIELD_DEC_LIST"] = GenFieldDecList(headDefinition);
            genMap["DATA_LIST"] = GenDataList(data, headDefinition);
            genMap["DATA_HEAD_DEF"] = GenDataHeadDef(headDefinition);
            genMap["BY_FIELD_FUNC_LIST"] = GenByFieldFuncList(headDefinition);

            GenCode(genMap);
        }

        public void SaveTo(String path)
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
            File.WriteAllText(path, sourceCode);
        }
        
    }
}
