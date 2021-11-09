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
    class DefinitionType
    {
        public string name = "";

        public string id;

        public enum Type
        {
            Real = 0,
            String,
            Boolean,
        }

        [JsonConverter(typeof(StringEnumConverter))]
        public Type type = Type.String;
    }

    class HeadDefinition : IEnumerable
    {
        List<DefinitionType> definitionList;
        public HeadDefinition(String path)
        {
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

        public IEnumerator GetEnumerator()
        {
            return definitionList.GetEnumerator();
        }

        public int GetDefinitionCount()
        {
            return definitionList.Count;
        }

        public class DuplicateIndentityException : Exception { public DuplicateIndentityException(String msg) : base(msg) { } }
        public class EmptyIndentity : Exception { public EmptyIndentity(String msg) : base(msg) { } }
    }
}
