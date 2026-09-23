using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Xml;

namespace XlsxDoctor
{
    /// <summary>
    /// 输出包完整性校验：XML 可解析 / 关系目标存在 / Content_Types 覆盖完整 / 无悬空 Override。
    /// 任何一条不过，调用方必须丢弃输出文件。
    /// </summary>
    public static class PackageValidator
    {
        static readonly Regex RelsRx = new Regex("<Relationship\\b([^>]*)/>", RegexOptions.Compiled);
        static readonly Regex DefaultRx = new Regex("<Default\\b([^>]*)/>", RegexOptions.Compiled);
        static readonly Regex OverrideRx = new Regex("<Override\\b([^>]*)/>", RegexOptions.Compiled);
        static readonly Regex XmlPartRx = new Regex("\\.(xml|rels|vml)$", RegexOptions.Compiled);
        static readonly Regex AbsoluteUriRx = new Regex("^[A-Za-z][A-Za-z0-9+.\\-]*:", RegexOptions.Compiled);

        public static List<string> Validate(string path)
        {
            List<string> errs = new List<string>();

            using (FileStream fs = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (ZipArchive zip = new ZipArchive(fs, ZipArchiveMode.Read))
            {
                List<ZipArchiveEntry> files = new List<ZipArchiveEntry>();
                HashSet<string> nameSet = new HashSet<string>(StringComparer.Ordinal);
                foreach (ZipArchiveEntry e in zip.Entries)
                {
                    if (e.Length == 0 && e.FullName.EndsWith("/", StringComparison.Ordinal)) continue;
                    files.Add(e);
                    nameSet.Add(e.FullName);
                }

                // 1) 每个 XML 部件都必须能解析
                for (int i = 0; i < files.Count; i++)
                {
                    ZipArchiveEntry f = files[i];
                    if (!XmlPartRx.IsMatch(f.FullName)) continue;
                    try { CheckWellFormed(Util.ReadAllText(f)); }
                    catch (Exception) { errs.Add("XML 无法解析: " + f.FullName); }
                }

                // 2) 每条内部关系都必须指向存在的部件
                for (int i = 0; i < files.Count; i++)
                {
                    ZipArchiveEntry f = files[i];
                    if (!f.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)) continue;
                    string dir = Util.GetRelDir(f.FullName);
                    string text = Util.ReadAllText(f);
                    foreach (Match m in RelsRx.Matches(text))
                    {
                        string a = m.Groups[1].Value;
                        if (Util.GetAttr(a, "TargetMode") == "External") continue;
                        string t = Util.GetAttr(a, "Target");
                        if (t.Length == 0) continue;

                        // Target 是 URI 引用，可能带片段。形状超链接指向本工作簿单元格时写的就是
                        // 纯片段（例如 "#汇总说明!A1"）——按 OPC 规范它指向源部件自身，没有独立部件可查。
                        int hash = t.IndexOf('#');
                        string partPath = hash >= 0 ? t.Substring(0, hash) : t;
                        if (partPath.Length == 0) continue;

                        // 带 scheme 的绝对 URI（http://、file:// 等）不是包内部件
                        if (AbsoluteUriRx.IsMatch(partPath)) continue;

                        string res = partPath.StartsWith("/", StringComparison.Ordinal)
                            ? partPath.TrimStart('/')
                            : (dir.Length == 0 ? partPath : dir + "/" + partPath);
                        res = Util.NormalizePartPath(res);
                        if (!nameSet.Contains(res))
                            errs.Add("关系目标不存在: " + f.FullName + " -> " + t);
                    }
                }

                // 3) Content_Types 必须覆盖全部部件，且不能有悬空 Override
                ZipArchiveEntry ctEntry = null;
                for (int i = 0; i < files.Count; i++)
                    if (files[i].FullName == "[Content_Types].xml") { ctEntry = files[i]; break; }

                if (ctEntry == null)
                {
                    errs.Add("缺少 [Content_Types].xml");
                }
                else
                {
                    string ct = Util.ReadAllText(ctEntry);

                    HashSet<string> defaults = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    foreach (Match m in DefaultRx.Matches(ct))
                        defaults.Add(Util.GetAttr(m.Groups[1].Value, "Extension"));

                    HashSet<string> overrides = new HashSet<string>(StringComparer.Ordinal);
                    foreach (Match m in OverrideRx.Matches(ct))
                        overrides.Add(Util.GetAttr(m.Groups[1].Value, "PartName"));

                    for (int i = 0; i < files.Count; i++)
                    {
                        string fn = files[i].FullName;
                        string ext = Path.GetExtension(fn).TrimStart('.').ToLowerInvariant();
                        if (!overrides.Contains("/" + fn) && !defaults.Contains(ext))
                            errs.Add("缺少 ContentType 声明: " + fn);
                    }

                    foreach (string o in overrides)
                        if (!nameSet.Contains(o.TrimStart('/')))
                            errs.Add("Content_Types 指向不存在的部件: " + o);
                }
            }

            return errs;
        }

        static void CheckWellFormed(string xml)
        {
            XmlReaderSettings s = new XmlReaderSettings();
            s.DtdProcessing = DtdProcessing.Prohibit;
            s.XmlResolver = null;
            s.CloseInput = true;
            using (StringReader sr = new StringReader(xml))
            using (XmlReader xr = XmlReader.Create(sr, s))
            {
                while (xr.Read()) { }
            }
        }
    }
}
