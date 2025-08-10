using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;

using Excel = Microsoft.Office.Interop.Excel;

namespace CrosslChecker
{
    internal class Program
    {
        static Encoding CfileEncoding = Encoding.Default;
        static List<string> KeyWords = new List<string>() {
            "u1", "u2","u4","s1","s2","s4",
            "auto", "break", "case", "char", "const", "continue", 
            "define", "default", "do", "double", "else", "enum", "extern", 
            "float", "for", "goto", "if", "inline", "int", "long", 
            "register", "restrict", "return", "short", "signed", 
            "sizeof", "static", "struct", "switch", "typedef", 
            "typeof", "union", "unsigned", "void", "volatile" };

        static JobObject JobObject = new JobObject();

        [DllImport("user32")]
        private extern static int GetWindowThreadProcessId(int hwnd, out int lpdwprocessid);

        struct Result
        {
            public List<string> Def;
            public List<string> Ref;
            public Result(List<string> def, List<string> @ref)
            {
                Def = def;
                Ref = @ref;
            }

        }

        static void Main(string[] args)
        {
            //var xlApp = CreateExcelApplication();

            Console.WriteLine();
            Console.ForegroundColor = ConsoleColor.Red;

            if (args.Length == 1)
            {
                string l_mat = null;
                string c_mat = null;
                string header = null;

                var files = Directory.GetFiles(args[0], "*", SearchOption.TopDirectoryOnly).Where(x => File.Exists(x)).ToList();

                foreach (var name in files)
                {
                    if (name.EndsWith(".c"))
                    {
                        if (string.IsNullOrEmpty(l_mat))
                        {
                            l_mat = name;
                        }

                        if (name.EndsWith("l_mat.c"))
                        {
                            l_mat = name;
                        }
                        if (name.EndsWith("c_mat.c"))
                        {
                            c_mat = name;
                        }
                    }
                    if (name.EndsWith(".h"))
                    {
                        header = name;
                    }
                }

                CheckCfiles(l_mat, c_mat, header);
            }
            else
            {
                Console.WriteLine("引数がおかしい");
                Console.ResetColor();
            }

            Console.WriteLine();
            Console.ReadKey();
        }

        static private void CheckCfiles(string path_l_mat, string path_c_mat, string path_header)
        {
            List<string> add_ref;
            var inc_words = new List<string>();
            var mac_words = new List<string>();
            var var_words = new List<string>();
            var ref1_words = new List<string>();
            var ref2_words = new List<string>();

            var def_words = new Dictionary<string, int>(1000);
            var ref_words = new Dictionary<string, int>(1000);
            var lines = File.ReadAllLines(path_header, CfileEncoding);
            var sec_list = SearchSections(lines);

            var inc_sta = SearchSection(lines, "インクルード");
            var inc_end = NextSection(sec_list, inc_sta);
            if (inc_sta >= 0 && inc_sta < inc_end)
            {
                inc_words = GetCommentWords(lines, inc_sta, inc_end);
            }

            var mac_sta = SearchSection(lines, "マクロ");
            var mac_end = NextSection(sec_list, mac_sta);
            if (mac_sta >= 0 && mac_sta < mac_end)
            {
                var r = GetDefRefMacros(lines, mac_sta, mac_end);
                mac_words = r.Def;
                ref1_words = r.Ref;
            }

            var var_sta = SearchSection(lines, "変数定義");
            var var_end = NextSection(sec_list, var_sta);
            if (var_sta >= 0 && var_sta < var_end)
            {
                var r = GetDefRefWords(lines, var_sta, var_end);
                var_words = r.Def;
                ref2_words = r.Ref;
            }


            foreach (var name in ref1_words.Concat(ref2_words))
            {
                if (!ref_words.ContainsKey(name))
                {
                    ref_words.Add(name, 1);
                }
            }

            foreach (var name in mac_words.Concat(var_words))
            {
                if (!def_words.ContainsKey(name))
                {
                    def_words.Add(name, 1);
                }
                else
                {
                    def_words[name] = def_words[name] + 1;
                }
            }

            /*  */
            add_ref = CheckCfile(path_l_mat, def_words.Keys.ToList(), new List<string>());

            if (!string.IsNullOrEmpty(path_c_mat))
            {
                /*  */
                add_ref = CheckCfile(path_c_mat, def_words.Keys.ToList(), add_ref);
            }

            foreach (var name in inc_words)
            {
                if (!def_words.ContainsKey(name))
                {
                    def_words.Add(name, 1);
                }
                else
                {
                    def_words[name] = def_words[name] + 1;
                }
            }

            /*  */
            CheckDefRef(def_words.Keys.Concat(add_ref).ToList(), ref_words.Keys.Concat(add_ref).ToList(), path_header);  
        }

        static private List<string> CheckCfile(string filename, List<string> head_words, List<string> add_ref)
        {
            var inc_words = new List<string>();
            var mac_words = new List<string>();
            var var_words = new List<string>();
            var fnc_words = new List<string>();
            var ref1_words = new List<string>();
            var ref2_words = new List<string>();
            var ref3_words = new List<string>();

            var def_words = new Dictionary<string, int>(1000);
            var ref_words = new Dictionary<string, int>(1000);
            var lines = File.ReadAllLines(filename, CfileEncoding);
            var sec_list = SearchSections(lines);

            var inc_sta = SearchSection(lines, "インクルード");
            var inc_end = NextSection(sec_list, inc_sta);
            if (inc_sta >= 0 && inc_sta < inc_end)
            {
                inc_words = GetCommentWords(lines, inc_sta, inc_end);
            }

            var mac_sta = SearchSection(lines, "マクロ");
            var mac_end = NextSection(sec_list, mac_sta);
            if (mac_sta >= 0 && mac_sta < mac_end)
            {
                var r = GetDefRefMacros(lines, mac_sta, mac_end);
                mac_words = r.Def;
                ref1_words = r.Ref;
            }

            var var_sta = SearchSection(lines, "変数定義");
            var var_end = NextSection(sec_list, var_sta);
            if (var_sta >= 0 && var_sta < var_end)
            {
                var r = GetDefRefWords(lines, var_sta, var_end);
                var_words = r.Def;
                ref2_words = r.Ref;
            }

            var fnc_sta = SearchSection(lines, "関数");
            var fnc_end = NextSection(sec_list, fnc_sta);
            if (fnc_sta >= 0 && fnc_sta < fnc_end)
            {
                var r = GetFuncDefRefWords(lines, fnc_sta, fnc_end);
                var_words = r.Def;
                ref3_words = r.Ref;
            }

            foreach (var name in head_words.Concat(inc_words).Concat(mac_words).Concat(var_words).Concat(fnc_words))
            {
                if (!def_words.ContainsKey(name))
                {
                    def_words.Add(name, 1);
                }
                else
                {
                    def_words[name] = def_words[name] + 1;
                }
            }

            var deflist = def_words.Keys.Concat(add_ref).ToList();

            foreach (var name in add_ref)
            {
                ref_words.Add(name, 1);
            }

            foreach (var name in ref1_words.Concat(ref2_words).Concat(ref3_words))
            {
                if (!ref_words.ContainsKey(name))
                {
                    ref_words.Add(name, 1);
                    add_ref.Add(name);
                }
            }


            CheckDefRef(deflist, ref_words.Keys.ToList(), filename);

            add_ref = add_ref.Except(inc_words).Distinct().ToList();

            return add_ref;
        }

        static private void CheckDefRef(List<string> def_list, List<string> ref_list, string filename)
        {
            var many_def = def_list.Except(ref_list.Concat(KeyWords)).ToList();
            var many_ref = ref_list.Except(def_list.Concat(KeyWords)).ToList();

            if (many_def.Count > 0)
            {
                Console.ForegroundColor = ConsoleColor.White;
                Console.BackgroundColor = ConsoleColor.Red;
                Console.WriteLine("■定義が多い " + filename);
                Console.ResetColor();
                foreach (var name in many_def)
                {
                    Console.WriteLine(" " + name);
                }
            }

            if (many_ref.Count > 0)
            {
                Console.ForegroundColor = ConsoleColor.White;
                Console.BackgroundColor = ConsoleColor.Red;
                Console.WriteLine("■定義が足りてない " + filename);
                Console.ResetColor();
                foreach (var name in many_ref)
                {
                    Console.WriteLine(" " + name);
                }
            }

            Console.ResetColor();
        }

        static private List<string> GetWords(string line)
        {
            var deflist = new List<string>();
            var defreg = new Regex(@"(?<WORD>[a-zA-Z_][\w]*)([\W]+(?<WORD>[a-zA-Z_][\w]*))*");

            foreach (Match match in defreg.Matches(line))
            {
                var captures = match.Groups["WORD"].Captures;

                foreach (Capture captur in captures)
                {
                    deflist.Add(captur.Value);
                }
            }

            return deflist;
        }

        static private Result GetFuncDefRefWords(string[] lines, int sta, int end)
        {
            int st = 0;
            int nest = 0;
            var autovarlist = new List<string>();
            var allline = string.Join(" ", lines, sta, end - sta).Replace('\t',' ');
            var deflist = new List<string>();
            var reflist = new List<string>();
            var coment = new Regex(@"/\*.*?\*/");
            var defreg = new Regex(@"(?<WORD>[a-zA-Z_][\w]*)([\W]+(?<WORD>[a-zA-Z_][\w]*))*");

            // コメント削除
            var sb = new StringBuilder();
            var mcmt = coment.Match(allline);
            while (mcmt.Success)
            {
                sb.Clear();
                sb.Append(allline.Substring(0, mcmt.Index));
                sb.Append(' ');
                sb.Append(allline.Substring(mcmt.Index + mcmt.Length));
                sb.Append(' ');
                allline = sb.ToString();
                mcmt = coment.Match(allline);
            }

            // 関数ごとに分割
            sb.Clear();
            foreach (var c in allline)
            {
                if (st == 0)
                {
                    if (c != '(')
                    {
                        sb.Append(c);
                    }
                    else
                    {
                        deflist.Add(sb.ToString().Split(new char[] { ' ' }, System.StringSplitOptions.RemoveEmptyEntries).Last());
                        sb.Clear();
                        st = 1;
                        nest = 1;
                        autovarlist.Clear();
                    }
                }
                else if (st == 1)
                {
                    sb.Append(c);

                    if (c == '(')
                    {
                        nest++;
                    }
                    else if (c == ')')
                    {
                        nest--;
                    }

                    if (nest == 0)
                    {
                        autovarlist.AddRange(GetWords(sb.ToString()));
                        sb.Clear();
                        st = 2;
                    }
                }
                else if (st == 2)
                {
                    if (c == '{')
                    {
                        st = 3;
                        nest = 1;
                    }
                }
                else if (st == 3)
                {
                    sb.Append(c);

                    if (c == '{')
                    {
                        nest++;
                    }
                    else if (c == '}')
                    {
                        nest--;
                    }

                    if (nest == 0)
                    {
                        var r = GetFuncInnerDefRefWords(sb.ToString());
                        autovarlist.AddRange(r.Def);
                        r.Ref.RemoveAll(autovarlist.Contains);
                        reflist.AddRange(r.Ref);
                        sb.Clear();
                        st = 0;
                    }
                }
                else if (st == 4)
                {

                }
            }

            return new Result(deflist, reflist);
        }

        static private Result GetFuncInnerDefRefWords(string allline)
        {
            int last = 0;
            var sb = new StringBuilder();
            var deflist = new List<string>();
            var defreg = new Regex(@"(?<WORD>[a-zA-Z_][\w]*)([\s]+(?<DEFWORD>[a-zA-Z_][\w]*)+)[\W]*?(=.*?)?;");

            var matches = defreg.Matches(allline);
            if (matches.Count == 0)
            {
                return new Result(deflist, GetWords(allline));
            }
            else
            {
                for (int i = 0; i < matches.Count; i++)
                {
                    Match match = matches[i];
                    var captures = match.Groups["DEFWORD"].Captures;

                    foreach (Capture captur in captures)
                    {
                        deflist.Add(captur.Value);
                    }

                    if (i < matches.Count - 1)
                    {
                        sb.Append(allline.Substring(last, match.Index - last));
                        sb.Append(' ');
                        last = match.Index + match.Length;
                    }
                    else
                    {
                        sb.Append(allline.Substring(last, match.Index - last));
                        sb.Append(' ');
                        sb.Append(allline.Substring(match.Index + match.Length));
                        sb.Append(' ');
                    }
                }

                return new Result(deflist, GetWords(sb.ToString()));
            }  
        }

        static private Result GetDefRefWords(string[] lines, int sta, int end)
        {
            var sb = new StringBuilder();
            var allline = string.Join(" ", lines, sta, end - sta);
            var deflist = new List<string>();
            var reflist = new List<string>();
            var coment = new Regex(@"/\*.*?\*/");
            var defreg = new Regex(@"(?<WORD>[a-zA-Z_][\w]*)([\W]+(?<WORD>[a-zA-Z_][\w]*))*");
            var refreg = new Regex(@"=[\W]*(?<WORD>[a-zA-Z_][\w]*)([\W]+(?<WORD>[a-zA-Z_][\w]*))*?[\W]*;");

            // コメント削除
            var mcmt = coment.Match(allline);
            while (mcmt.Success)
            {
                sb.Clear();
                sb.Append(allline.Substring(0, mcmt.Index));
                sb.Append(' ');
                sb.Append(allline.Substring(mcmt.Index + mcmt.Length));
                sb.Append(' ');
                allline = sb.ToString();
                mcmt = coment.Match(allline);
            }

            /* 参照部 */
            foreach (Match match in refreg.Matches(allline))
            {
                var captures = match.Groups["WORD"].Captures;

                foreach (Capture captur in captures)
                {
                    reflist.Add(captur.Value);
                }
            }

            var refmcmt = refreg.Match(allline);
            while (refmcmt.Success)
            {
                sb.Clear();
                sb.Append(allline.Substring(0, refmcmt.Index));
                sb.Append(' ');
                sb.Append(allline.Substring(refmcmt.Index + refmcmt.Length));
                sb.Append(' ');
                allline = sb.ToString();
                refmcmt = refreg.Match(allline);
            }


            /* 定義部 */
            foreach (Match match in defreg.Matches(allline))
            {
                var captures = match.Groups["WORD"].Captures;

                foreach (Capture captur in captures)
                {
                    deflist.Add(captur.Value);
                }
            }

            return new Result(deflist, reflist);
        }

        static private Result GetDefRefMacros(string[] lines, int sta, int end)
        {
            var allline = string.Join(" ", lines, sta, end - sta);
            var deflist = new List<string>();
            var reflist = new List<string>();
            var coment = new Regex(@"/\*.*?\*/");
            var defreg = new Regex(@"#\s*define\s+(?<DEFWORD>[a-zA-Z_][\w]*)([\W]+(?<REFWORD>[a-zA-Z_][\w]*))*?[\W]+?(?=#\s*define\s+)");

            // コメント削除
            var sb = new StringBuilder();
            var mcmt = coment.Match(allline);
            while (mcmt.Success)
            {
                sb.Clear();
                sb.Append(allline.Substring(0, mcmt.Index));
                sb.Append(' ');
                sb.Append(allline.Substring(mcmt.Index + mcmt.Length));
                sb.Append(' ');
                allline = sb.ToString();
                mcmt = coment.Match(allline);
            }

            var matches = defreg.Matches(allline + " #define ");

            /* 参照部 */
            foreach (Match match in matches)
            {
                var captures = match.Groups["REFWORD"].Captures;

                foreach (Capture captur in captures)
                {
                    reflist.Add(captur.Value);
                }
            }


            /* 定義部 */
            foreach (Match match in matches)
            {
                var captures = match.Groups["DEFWORD"].Captures;

                foreach (Capture captur in captures)
                {
                    deflist.Add(captur.Value);
                }
            }

            return new Result(deflist, reflist);
        }

        static private List<string> GetCommentWords(string[] lines, int sta, int end)
        {
            var list = new List<string>();
            var reg = new Regex(@"/\*[\W]*((?<WORD>[a-zA-Z_][\w]*)([\W]+(?<WORD>[a-zA-Z_][\w]*))*)[\W]*\*/");

            for (var i = sta; i < end; i++)
            {
                var matches = reg.Matches(lines[i]);

                foreach (Match match in matches)
                {
                    if (match.Success)
                    {
                        var captures = match.Groups["WORD"].Captures;

                        foreach (Capture captur in captures)
                        {
                            list.Add(captur.Value);
                        }
                    }
                }
            }

            return list;
        }

        static private int NextSection(List<int> sec_list, int num)
        {
            for (var i = 0; i < sec_list.Count - 2; i++)
            {
                if (sec_list[i] == num)
                {
                    return sec_list[i + 1] - 3;
                }
            }

            return sec_list.Last();
        }

        static private List<int> SearchSections(string[] lines)
        {
            var list = new List<int>();
            var enclosure = new Regex(@"/\*[\*-=\s]+\*/");
            var reg = new Regex(@"/\*[\*-=\s]+[^\*-=\s]+[\*-=\s]+\*/");

            for (var i = 1; i < lines.Length - 2; i++)
            {
                if (reg.IsMatch(lines[i]))
                {
                    if (enclosure.IsMatch(lines[i - 1]) && enclosure.IsMatch(lines[i + 1]))
                    {
                        list.Add(i + 2);
                    }
                }
            }

            list.Add(lines.Length - 1);

            return list;
        }

        static private int SearchSection(string[] lines, string keyword)
        {
            var enclosure = new Regex(@"/\*[\*-=\s]+\*/");
            var reg = new Regex(@"/\*[\*-=\s]+" + keyword + @"[\*-=\s]+\*/");

            for (var i = 1; i < lines.Length - 2; i++)
            {
                if (reg.IsMatch(lines[i]))
                {
                    if (enclosure.IsMatch(lines[i - 1]) && enclosure.IsMatch(lines[i + 1]))
                    {
                        return i + 2;
                    }
                }
            }

            return -1;
        }

        static private Excel.Application CreateExcelApplication()
        {
            int pid;
            var xlApp = new Excel.Application();
            Excel.Workbooks xlBooks = xlApp.Workbooks;
            Excel.Workbook xlBook = xlBooks.Add();
            Excel.Windows xlWindows = xlApp.Windows;
            Excel.Window xlWindow = xlWindows[1];
            xlApp.Visible = true;
            GetWindowThreadProcessId(xlWindow.Hwnd, out pid);
            JobObject.AddProcess(Process.GetProcessById(pid));
            xlApp.Visible = false;
            xlBook.Close(false);
            Marshal.ReleaseComObject(xlWindow);
            Marshal.ReleaseComObject(xlWindows);
            Marshal.ReleaseComObject(xlBook);
            Marshal.ReleaseComObject(xlBooks);
            return xlApp;
        }

    }
}
