using System;
using System.Collections.Generic;
using System.IO;

namespace GDScriptConfigTableTool
{
    class Program
    {
        const string helpInfo = @"
#       GDScript Config Table Tool        #
#=========================================#
#                                         #
#           Author: Raiix                 #
#                                         #
#=========================================#

gds_cofig_table_tool <command> [options]

command:
    export_all_excel -od <output_dir> -dd <defs_dir>
    export_all_gdscript -od <output_dir> -ed <excels_dir> -dd <defs_dir> -tp <template_path>
    export_excel -od <output_dir> -dd <defs_dir> [-t <table_name> | -w <workbook_name>]
    export_gdscript -od <output_dir> -ed <excels_dir> -dd <defs_dir> -tp <template_path> [-t <table_name> | -w <workbook_name>]

options:
    -h : help


";

        Dictionary<String, String> optionMap;

        String GetOptionValue(String op)
        {
            if (optionMap.ContainsKey(op))
            {
                return optionMap[op];
            }
            Console.WriteLine($"Lack of param: {op}");
            return null;
        }

        static void Main(string[] args)
        {
            try
            {
                (new Program()).Run(args);
            } catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine(e.StackTrace);
            }
        }
        void Run(string[] args)
        {
            String cmd = "";
            optionMap = new Dictionary<string, string>();
            for (int i=0; i<args.Length; ++i)
            {
                String arg = args[i];
                switch (arg)
                {
                    case "-?":
                    case "-h":
                        Console.WriteLine(helpInfo);
                        return;
                }
                if (arg.StartsWith('-'))
                {
                    var op = arg.TrimStart('-');
                    if (i + 1 >= args.Length)
                    {
                        Console.WriteLine($"param {op} is empty!");
                        return;
                    }
                    var v = args[i + 1];
                    optionMap[op] = v;
                    i += 1;
                } else
                {
                    cmd = arg;
                }
            }

            switch (cmd)
            {
                case "export_all_excel":
                    {
                        var od = GetOptionValue("od");
                        var dd = GetOptionValue("dd");
                        if (od == null || dd == null) return;

                        var tool = new GDScriptConfigTableTool.ConfigTableTool.Tool();
                        tool.ExportExcelFile(od, dd);
                    }
                    break;
                case "export_all_gdscript":
                    {
                        var od = GetOptionValue("od");
                        var dd = GetOptionValue("dd");
                        var ed = GetOptionValue("ed");
                        var tp = GetOptionValue("tp");
                        if (od == null || dd == null || ed == null || tp == null) return;

                        var tool = new GDScriptConfigTableTool.ConfigTableTool.Tool();
                        tool.ExportGDScript(od, File.ReadAllText(tp), ed, dd);
                    }
                    break;
                case "export_excel":
                    {
                        var od = GetOptionValue("od");
                        var dd = GetOptionValue("dd");
                        if (od == null || dd == null) return;
                        var t = optionMap["t"];
                        var w = optionMap["w"];
                        if (t == null && w == null) Console.WriteLine("Lack of param: t or w!");

                        var tool = new GDScriptConfigTableTool.ConfigTableTool.Tool();
                        if (t != null)
                        {
                            var hd = new GDScriptConfigTableTool.ConfigTableTool.HeadDefinition(Path.Combine(dd, tool.GenHeadDefinitionFileName(t)));
                            tool.ExportExcelFile(od, new List<GDScriptConfigTableTool.ConfigTableTool.HeadDefinition> { hd });
                        } else
                        {
                            var hdList = new List<GDScriptConfigTableTool.ConfigTableTool.HeadDefinition>();
                            foreach (var file in Directory.EnumerateFiles(dd))
                            {
                                if (file.StartsWith(w))
                                {
                                    hdList.Add(new GDScriptConfigTableTool.ConfigTableTool.HeadDefinition(Path.Combine(dd, file)));
                                }
                            }
                            tool.ExportExcelFile(od, hdList);
                        }
                    }
                    break;
                case "export_gdscript":
                    {
                        var od = GetOptionValue("od");
                        var dd = GetOptionValue("dd");
                        var ed = GetOptionValue("ed");
                        var tp = GetOptionValue("tp");
                        if (od == null || dd == null || ed == null || tp == null) return;
                        var t = optionMap["t"];
                        var w = optionMap["w"];
                        if (t == null && w == null) Console.WriteLine("Lack of param: t or w!");

                        var tool = new GDScriptConfigTableTool.ConfigTableTool.Tool();
                        tool.ExportGDScript(od, File.ReadAllText(tp), ed, dd);

                        if (t != null)
                        {
                            var hd = new GDScriptConfigTableTool.ConfigTableTool.HeadDefinition(Path.Combine(dd, tool.GenHeadDefinitionFileName(t)));
                            var data = new GDScriptConfigTableTool.ConfigTableTool.Data();
                            data.CreateWorkbookFromFile(Path.Combine(ed, tool.GenExcelFileName(hd.WorkbookName)));
                            tool.ExportGDScript(od, File.ReadAllText(tp), data, hd);
                        } else
                        {
                            var data = new GDScriptConfigTableTool.ConfigTableTool.Data();
                            data.CreateWorkbookFromFile(Path.Combine(ed, tool.GenExcelFileName(w)));
                            var hdList = new List<GDScriptConfigTableTool.ConfigTableTool.HeadDefinition>();
                            foreach (var file in Directory.EnumerateFiles(dd))
                            {
                                if (file.StartsWith(w))
                                {
                                    hdList.Add(new GDScriptConfigTableTool.ConfigTableTool.HeadDefinition(Path.Combine(dd, file)));
                                }
                            }
                            tool.ExportGDScript(od, File.ReadAllText(tp), data, hdList);
                        }
                        
                    }
                    break;
                case "":
                    Console.WriteLine(helpInfo);
                    break;
                default:
                    Console.WriteLine($"Unsupported command: {cmd}");
                    break;
            }
        }
    }
}
