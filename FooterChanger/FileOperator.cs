using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
namespace Airdl
{
    class FileOperator
    {

        /// <summary>
        /// 创建文件
        /// </summary>
        /// <param name="path"></param>
        public static void Create(string path)
        {
            FileStream filestream = new FileStream(path, FileMode.Create);
            StreamWriter writer = new StreamWriter(filestream);
            writer.Close();
            filestream.Close();
        }
        /// <summary>
        /// 写文件
        /// </summary>
        /// <param name="path">文件路径</param>
        /// <param name="content">写入的内容</param>
        /// <param name="encoding">字符编码</param>
        public static void Write(string path, string content, Encoding encoding = null, bool append = false)
        {
            if (encoding == null)
                encoding = Encoding.UTF8;
            FileStream filestream;
            if (append)
                filestream = new FileStream(path, FileMode.Append);
            else
                filestream = new FileStream(path, FileMode.Create);
            StreamWriter writer = new StreamWriter(filestream, encoding);
            writer.Write(content);
            writer.Flush();
            writer.Close();
            filestream.Close();
        }


        /// <summary>
        /// 写文件
        /// </summary>
        /// <param name="path"></param>
        /// <param name="contentFile"></param>
        /// <param name="encoding"></param>
        public static void Write(string path, HashSet<string> contentFile, Encoding encoding = null)
        {
            if (encoding == null)
                encoding = Encoding.UTF8;
            StringBuilder content = new StringBuilder();
            foreach (string str in contentFile)
            {
                content.Append(str).Append("\n");
            }
            Write(path, content.ToString(), encoding);
        }
        /// <summary>
        /// 写文件
        /// </summary>
        /// <param name="path"></param>
        /// <param name="content"></param>
        /// <param name="encoding"></param>
        public static void Write(string path, Dictionary<string, float> contentFile, Encoding encoding = null)
        {
            if (encoding == null)
                encoding = Encoding.UTF8;
            String content = "";
            foreach (KeyValuePair<string, float> pair in contentFile)
            {
                content += pair.Value + ",";
            }
            content.Remove(content.Length - 1, 1);
            content += "\n";
            Write(path, content.ToString(), encoding, true);
        }
        /// <summary>
        /// 读文件
        /// </summary>
        /// <param name="path">文件路径</param>
        /// <returns>文件文本内容</returns>
        public static string Read(string path)
        {
            StringBuilder ret = new StringBuilder();
            if (System.IO.File.Exists(path))
            {
                StreamReader reader = new StreamReader(path);
                while (!reader.EndOfStream)
                    ret.Append(reader.ReadLine()).Append('\n');
                reader.Close();
            }
            return ret.ToString();
        }

        /// <summary>
        /// 删除文件
        /// </summary>
        /// <param name="path">所删除文件的位置</param>
        public static void Delete(string path)
        {
            if (System.IO.File.Exists(path))
                System.IO.File.Delete(path);
        }

        /// <summary>
        /// 复制文件
        /// </summary>
        /// <param name="sourceFile"></param>
        /// <param name="destFile"></param>
        public static void Copy(string sourceFile, string destFile)
        {
            if (System.IO.File.Exists(sourceFile))
                System.IO.File.Copy(sourceFile, destFile, true);
        }


        /// <summary>
        /// 创建文件夹
        /// </summary>
        /// <param name="dir">所创建的文件夹名称</param>
        public static void MKDir(string dir)
        {
            if (!System.IO.Directory.Exists(dir))
                System.IO.Directory.CreateDirectory(dir);
        }

    }
}
