using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Text;

using Newtonsoft.Json;
using Newtonsoft.Json.Converters;

namespace GDScriptConfigTableTool.ConfigTableTool
{
    public class DefinitionType
    {
        public string name = "";

        public string id;

        public enum Type
        {
            Real = 0,
            String,
            Boolean,
            Array,
            Dictionary,
        }

        [JsonConverter(typeof(StringEnumConverter))]
        public Type type = Type.String;

        public bool ignore = false;

        public String comment = null;
    }

    public class HeadDefinition : IEnumerable
    {
        List<DefinitionType> definitionList;

        string[] name;
        public String WorkbookName
        {
            get
            {
                return name[0];
            }
        }
        public String SheetName
        {
            get
            {
                if (name.Length > 1)
                {
                    var temp = new List<String>();
                    for (int i = 1; i < name.Length; ++i)
                    {
                        temp.Add(name[i]);
                    }
                    return String.Join('_', temp);
                }
                return name[0];
            }
        }
        String sourcePath;

        public String SourcePath
        {
            get
            {
                return sourcePath;
            }
        }

        public HeadDefinition(String path)
        {
            var fileName = Path.GetFileNameWithoutExtension(path);
            name = fileName.Split('_');
            if (name.Length == 0)
            {
                throw new InvalidFileName($"{fileName} in path: {path}");
            }

            sourcePath = path;

            var text = File.ReadAllText(path);
            definitionList = JsonConvert.DeserializeObject<List<DefinitionType>>(text);

            var set = new HashSet<String>();
            foreach (DefinitionType def in definitionList)
            {
                // Check dupicate identity
                if (set.Contains(def.id))
                {
                    throw new DuplicateIndentityException($"Duplicate identity: {def.id} in definition: {path}");
                }

                // Check empty identity
                if (def.id == null || def.id.Length == 0)
                {
                    throw new EmptyIndentity($"In definition: {path}");
                }
            }
        }

        public bool IsUsingOneName()
        {
            return name.Length == 1;
        }

        public IEnumerator GetEnumerator()
        {
            return definitionList.GetEnumerator();
        }

        public int GetDefinitionCount()
        {
            return definitionList.Count;
        }

        public class InvalidFileName : Exception { public InvalidFileName(String msg) : base(msg) { } }
        public class DuplicateIndentityException : Exception { public DuplicateIndentityException(String msg) : base(msg) { } }
        public class EmptyIndentity : Exception { public EmptyIndentity(String msg) : base(msg) { } }
    }
}
