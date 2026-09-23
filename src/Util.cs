using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;

namespace XlsxDoctor
{
    /// <summary>一个文本部件，连同它的原始编码与 BOM 状态，保证写回时编码不变。</summary>
    public sealed class TextPart
    {
        public string Name;
        public string Text;
        public Encoding Encoding;
        public bool HasBom;

        public byte[] Encode()
        {
            byte[] body = Encoding.GetBytes(Text);
            if (!HasBom) return body;
            byte[] pre = Encoding.GetPreamble();
            if (pre == null || pre.Length == 0) return body;
            byte[] all = new byte[pre.Length + body.Length];
            Buffer.BlockCopy(pre, 0, all, 0, pre.Length);
            Buffer.BlockCopy(body, 0, all, pre.Length, body.Length);
            return all;
        }
    }

    internal static class Util
    {
        static readonly Regex EncDeclRx = new Regex("encoding\\s*=\\s*\"([^\"]+)\"", RegexOptions.Compiled);

        /// <summary>取 XML 元素属性值，属性顺序无关。</summary>
        public static string GetAttr(string attrText, string name)
        {
            Match m = Regex.Match(attrText, "\\b" + Regex.Escape(name) + "=\"([^\"]*)\"");
            return m.Success ? m.Groups[1].Value : "";
        }

        public static byte[] ReadAllBytes(ZipArchiveEntry e)
        {
            using (Stream s = e.Open())
            using (MemoryStream ms = new MemoryStream())
            {
                byte[] buf = new byte[65536];
                int n;
                while ((n = s.Read(buf, 0, buf.Length)) > 0) ms.Write(buf, 0, n);
                return ms.ToArray();
            }
        }

        public static string ReadAllText(ZipArchiveEntry e)
        {
            return DecodePart(e.FullName, ReadAllBytes(e)).Text;
        }

        /// <summary>按 BOM / XML 声明探测编码，并返回解码后的文本。</summary>
        public static TextPart DecodePart(string name, byte[] bytes)
        {
            Encoding enc;
            bool hasBom = false;
            int bomLen = 0;

            if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
            {
                enc = new UTF8Encoding(true); hasBom = true; bomLen = 3;
            }
            else if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE)
            {
                enc = Encoding.Unicode; hasBom = true; bomLen = 2;
            }
            else if (bytes.Length >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF)
            {
                enc = Encoding.BigEndianUnicode; hasBom = true; bomLen = 2;
            }
            else
            {
                enc = new UTF8Encoding(false);
                // 没有 BOM 时看 XML 声明；OOXML 规定是 UTF-8，但别假设
                int probe = Math.Min(bytes.Length, 200);
                string head = Encoding.ASCII.GetString(bytes, 0, probe);
                Match m = EncDeclRx.Match(head);
                if (m.Success)
                {
                    try
                    {
                        enc = Encoding.GetEncoding(m.Groups[1].Value);
                        if (enc is UTF8Encoding) enc = new UTF8Encoding(false);
                    }
                    catch (ArgumentException) { enc = new UTF8Encoding(false); }
                }
            }

            TextPart p = new TextPart();
            p.Name = name;
            p.Encoding = enc;
            p.HasBom = hasBom;
            p.Text = enc.GetString(bytes, bomLen, bytes.Length - bomLen);
            return p;
        }

        public static string GetRelDir(string relsPartName)
        {
            int i = relsPartName.LastIndexOf("/_rels/", StringComparison.Ordinal);
            if (i < 0) return "";
            return relsPartName.Substring(0, i);
        }

        /// <summary>把部件路径规范化（合并 . 与 ..），用于校验关系目标是否存在。</summary>
        public static string NormalizePartPath(string p)
        {
            p = p.Replace('\\', '/');
            string[] segs = p.Split('/');
            List<string> stack = new List<string>();
            for (int i = 0; i < segs.Length; i++)
            {
                string s = segs[i];
                if (s.Length == 0 || s == ".") continue;
                if (s == "..")
                {
                    if (stack.Count > 0) stack.RemoveAt(stack.Count - 1);
                    continue;
                }
                stack.Add(s);
            }
            return string.Join("/", stack.ToArray());
        }
    }
}
