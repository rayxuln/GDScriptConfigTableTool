using System;

namespace GDScriptConfigTableTool
{
    class Program
    {
        static void Main(string[] args)
        {
            (new Program()).Run(args);
        }

        String baseDir = "J:/gg_project/GDScriptConfigTableTool/assets";
        void Run(string[] args)
        {
            var hd = System.IO.Path.Combine(baseDir, "test_s1.json");
            var tool = new ConfigTableTool.Tool();
            tool.ExportHeadOnlyExcelFile(baseDir, hd);
        }
    }
}
