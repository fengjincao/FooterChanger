using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace FooterChanger
{
    class WordOperator
    {

        private bool unEdited = false;



        public static string AlertWordFooter(String filePath, String newFooter)
        {
            //read only 模式读取 word
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
            {
                // Get a main document part. 
                MainDocumentPart mainPart = doc.MainDocumentPart;
                // Get the document structure and add some text.

                Body body = mainPart.Document.Body;

                //body contains all the paragraph

                var textBoxInFirstPara = body.GetFirstChild<Paragraph>().Descendants<TextBoxContent>();
                foreach (var textBoxItem in textBoxInFirstPara)
                {
                    foreach (var txt in textBoxItem.Descendants<Text>())
                    {
                        if (txt.Text.Contains("Animbus"))
                        {
                            txt.Text = txt.Text;
                        }
                    }
                }
                return "heading";

                var oldHeader = mainPart.HeaderParts;
                foreach (var item in oldHeader)
                {
                    if (item.Header.InnerText.Contains("6.5"))
                    {
                        int changeNum = 0;
                        foreach (var text in item.RootElement.Descendants<Text>())
                        {
                            if (changeNum>0 || text.Text.Contains("Animbus"))
                            {
                                if (text.Text.Contains("Animbus"))
                                {
                                    changeNum=3;
                                }
                                else
                                {
                                    if (changeNum == 3)
                                    {
                                        text.Text = text.Text.Replace("6", "7");
                                    }
                                    else if (changeNum == 1)
                                    {
                                        text.Text = text.Text.Replace("5", "7");
                                    }
                                    
                                    changeNum--;
                                }
                               
                            }
                        } 
                    }
                }
                //So...  just get what you need  "
                String content = body.InnerText;
                
                return content;


                String content_xml = body.InnerXml;

                var paras = body.ChildElements;
            }
        }
        public static void CreateWordDoc(string filepath, string msg)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Create(filepath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
            {
                // Add a main document part. 
                MainDocumentPart mainPart = doc.AddMainDocumentPart();

                // Create the document structure and add some text.
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());
                Paragraph para = body.AppendChild(new Paragraph());
                Run run = para.AppendChild(new Run());

                // String msg contains the text, "Hello, Word!"
                run.AppendChild(new Text(msg));

            }
        }


        public static void AlterWordDoc(string filepath, string msg)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Create(filepath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
            {
                // Add a main document part. 
                MainDocumentPart mainPart = doc.AddMainDocumentPart();

                // Create the document structure and add some text.
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());
                Paragraph para = body.AppendChild(new Paragraph());
                Run run = para.AppendChild(new Run());

                // String msg contains the text, "Hello, Word!"
                run.AppendChild(new Text(msg));

            }
        }

        public static void DeleteWordDoc(string filepath, string msg)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Create(filepath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
            {
                // Add a main document part. 
                MainDocumentPart mainPart = doc.AddMainDocumentPart();

                // Create the document structure and add some text.
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());
                Paragraph para = body.AppendChild(new Paragraph());
                Run run = para.AppendChild(new Run());

                // String msg contains the text, "Hello, Word!"
                run.AppendChild(new Text(msg));

            }
        }
        public static DataSet ReadWordDocForTable(String filepath, Boolean isLonelyTable)
        {
            DataSet tableSet = new DataSet("Data");
            String strType = "System.String";
            DataTable dataDT = new DataTable("Data_dt");
            dataDT.Columns.Add("DName", System.Type.GetType(strType));
            dataDT.Columns.Add("NName", System.Type.GetType(strType));
            dataDT.Columns.Add("Type", System.Type.GetType(strType));
            dataDT.Columns.Add("Meaning", System.Type.GetType(strType));
            dataDT.Columns.Add("InitValue", System.Type.GetType(strType));
            dataDT.Columns.Add("Note", System.Type.GetType(strType));
            dataDT.Columns.Add("RangeType", System.Type.GetType(strType));
            dataDT.Columns.Add("RangeMin", System.Type.GetType(strType));
            dataDT.Columns.Add("RangeMax", System.Type.GetType(strType));


            //read only 模式读取 word
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filepath, false))
            {
                // Get a main document part. 
                MainDocumentPart mainPart = doc.MainDocumentPart;

                // Get the document structure and add some text.
                Body body = mainPart.Document.Body;

                //body contains all the paragraph

                foreach (Table tbl in body.Elements<Table>())
                {
                    //取得TableRow陣列
                    var rows = tbl.Elements<TableRow>().ToArray();
                    Boolean isFirstFlag = true;
                    DataTable dt = new DataTable();
                    if (!isLonelyTable)
                    {
                        String name = tbl.ElementsBefore().Last(c => c is Paragraph && c.InnerText.Contains("表")).InnerText;
                        dt.TableName = name.Replace("表", "").Split(' ')[0]; ;
                    }
                    for (int i = 0; i < rows.Length; i++)
                    {
                        #region
                        if (isLonelyTable)
                        {
                            //取得TableRow的TableCell陣列
                            var cells = rows[i].Elements<TableCell>().ToArray();
                            if (isFirstFlag)
                            {
                                if (cells[0].Elements<Paragraph>().First().InnerText.Contains("变量名"))
                                {
                                    isFirstFlag = false;
                                    continue;
                                }
                                else
                                {
                                    throw new Exception("请选择正确格式的数据字典文件！");
                                }
                            }

                            DataRow currRow = dataDT.NewRow();
                            //顯示每列的內容

                            for (int j = 0; j < cells.Length; j++)
                            {
                                currRow[j] = cells[j].Elements<Paragraph>().First().InnerText;
                            }
                            dataDT.Rows.Add(currRow);
                        }
                        #endregion
                        #region
                        else
                        {
                            var cells = rows[i].Elements<TableCell>().ToArray();
                            if (isFirstFlag)
                            {
                                for (int j = 0; j < cells.Length; j++)
                                {
                                    dt.Columns.Add(cells[j].Elements<Paragraph>().First().InnerText, System.Type.GetType(strType));
                                }
                                isFirstFlag = false;
                            }
                            else
                            {
                                DataRow currRow = dt.NewRow();
                                //顯示每列的內容

                                for (int j = 0; j < cells.Length; j++)
                                {
                                    currRow[j] = cells[j].Elements<Paragraph>().First().InnerText;
                                }
                                dt.Rows.Add(currRow);
                            }
                        }
                        #endregion
                    }
                    if (isLonelyTable)
                    {
                        tableSet.Tables.Add(dataDT);
                    }
                    else
                    {
                        tableSet.Tables.Add(dt);
                    }
                }

            }
            return tableSet;
        }

        public static String ReadWordDocForInnerText(string filepath)
        {
            //Name Guard Init Procedure Transitions Others
            String strType = "System.String";
            DataTable modeDT = new DataTable("mode_dt");
            modeDT.Columns.Add("Name", System.Type.GetType(strType));
            modeDT.Columns.Add("Guard", System.Type.GetType(strType));
            modeDT.Columns.Add("Init", System.Type.GetType(strType));
            modeDT.Columns.Add("Procedure", System.Type.GetType(strType));
            modeDT.Columns.Add("Transitions", System.Type.GetType(strType));
            modeDT.Columns.Add("Others", System.Type.GetType(strType));


            //read only 模式读取 word
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filepath, false))
            {
                // Get a main document part. 
                MainDocumentPart mainPart = doc.MainDocumentPart;
                // Get the document structure and add some text.

                Body body = mainPart.Document.Body;

                //body contains all the paragraph

                //So...  just get what you need  "
                String content = body.InnerText;

                return content;


                String content_xml = body.InnerXml;

                var paras = body.ChildElements;



                foreach (var para in paras)
                {

                    if (para is Paragraph)
                    {
                        String paraContent = para.InnerText;
                        string pattern = "任务名称:|编号:|功能:|前置条件:|输入:|输出:|局部变量:|公式:";

                        if (Regex.IsMatch(paraContent, pattern))
                        {
                            if (paraContent.Contains("公式"))
                            {

                            }
                        }

                        string[] module = Regex.Split(content, "(任务名称：)|(编号：)|(功能：)|(前置条件：)|(输入：)|(输出：)|(公式：)", RegexOptions.IgnoreCase);


                    }
                    else if (para is Table)
                    {


                    }
                }

            }
        }



        public static List<String> ReadWordForParaNum(string filePath)
        {
            List<String> ParaNums = new List<string>();
            //read only 模式读取 word
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, false))
            {
                // Get a main document part. 
                MainDocumentPart mainPart = doc.MainDocumentPart;

                // Get the document structure and add some text.
                Body body = mainPart.Document.Body;

                //body contains all the paragraph

                foreach (Paragraph tbl in body.Elements<Paragraph>())
                {
                    if (tbl.Elements<ParagraphProperties>().Count() > 1)
                    {
                        var temp = tbl.Elements<ParagraphProperties>();

                    }
                }

            }


            return ParaNums;
        }
    }
}
