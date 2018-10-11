using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;
using System.Data.SqlClient;
using System.Configuration;

namespace PDM
{
    public partial class PDM : Form
    {

        string cs = ConfigurationManager.ConnectionStrings["DBCS"].ConnectionString;

        public PDM()
        {
            InitializeComponent();
        }

        private void PDM_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.SelectedPath == "")
            {
                MessageBox.Show("Please enter a valid path to save the generated file.", "Important Message");
            }
            else {
                performOperation();
            }
            
        }

        private void performOperation()
        {
            List<TableInformation> tableData = new List<TableInformation>();
            var date = DateTime.Now;

            #region Connect To DB            
            string statement = Constatnts.statement1;
            using (SqlConnection con = new SqlConnection(cs))
            {
                SqlCommand executeStatement = new SqlCommand(statement, con);
                con.Open();
                SqlDataReader reader = executeStatement.ExecuteReader();
                while (reader.Read())
                {
                    TableInformation t = new TableInformation();
                    t.SynonymName = reader["Synonym_NAME"].ToString();
                    t.TableName = reader["Table_NAME"].ToString();
                    tableData.Add(t);
                }
            }

            List<TableInformation> sortedTableData = tableData.OrderBy(order => order.SynonymName).ToList();
            #endregion


            #region Starting Word Application
            object oMissing = System.Reflection.Missing.Value;
            object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */

            //Start Word and create a new document.
            Word._Application oWord;
            Word._Document oDoc;
            oWord = new Word.Application();
            oWord.Visible = true;
            oDoc = oWord.Documents.Add(ref oMissing, ref oMissing,
            ref oMissing, ref oMissing);
            #endregion


            //Insert a paragraph at the beginning of the document.
            Word.Paragraph oPara1;
            oPara1 = oDoc.Content.Paragraphs.Add(ref oMissing);
            oPara1.Range.Text = "All tables are under dbo schema and all columns are not nullable.";
            oPara1.Range.Font.Bold = 1;
            oPara1.Format.SpaceAfter = 24;    //24 pt spacing after paragraph.
            oPara1.Range.InsertParagraphAfter();

            #region Insert Table Information
            //Insert a table, fill it with data, and make the first row bold

            Word.Table oTable;
            Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oTable = oDoc.Tables.Add(wrdRng, sortedTableData.Count + 1, 3, ref oMissing, ref oMissing);
            oTable.Range.ParagraphFormat.SpaceAfter = 6;
            int r;
            oTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            oTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            oTable.Cell(1, 1).Range.Text = "NO";
            oTable.Cell(1, 2).Range.Text = "TABLE NAME";
            oTable.Cell(1, 3).Range.Text = "SYNONYM NAME";
            oTable.Cell(1, 1).Range.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray40;
            oTable.Cell(1, 2).Range.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray40;
            oTable.Cell(1, 3).Range.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray40;
            oTable.Range.Font.Bold = 0;
            int z = 0;
            for (r = 2; r <= sortedTableData.Count + 1; r++)
            {
                oTable.Cell(r, 1).Range.Text = (z + 1).ToString();
                oTable.Cell(r, 2).Range.Text = sortedTableData[z].TableName;
                oTable.Cell(r, 3).Range.Text = sortedTableData[z].SynonymName;
                z++;
            }

            oTable.Rows[1].Range.Font.Bold = 1;
            oTable.Columns[1].PreferredWidth = oWord.InchesToPoints(0.4f);
            oTable.Columns[2].PreferredWidth = oWord.InchesToPoints(2.5f);
            oTable.Columns[3].PreferredWidth = oWord.InchesToPoints(2.3f);
            #endregion


            //Add some text after the table.
            Word.Paragraph oPara4;
            object oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oPara4 = oDoc.Content.Paragraphs.Add(ref oRng);
            oPara4.Range.InsertParagraphBefore();
            oPara4.Range.Text = "Detailed Table Information";
            oPara4.Format.SpaceAfter = 6;
            oPara4.Range.InsertParagraphAfter();

            #region Setting Header
            foreach (Word.Section section in oDoc.Sections)
            {
                Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldPage);
                headerRange.Text = "Govt Portal \t \t Protech Solutions \nPDM";
            }
            #endregion

            #region Setting Footer
            foreach (Word.Section wordSection in oDoc.Sections)
            {
                Word.Range footerRange = wordSection.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                footerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                footerRange.Text = date.ToString("MMMM dd, yyyy");
            }
            #endregion

            //Get Table Meta Data

            foreach (var t in sortedTableData)
            {
                statement = Constatnts.statement2;
                List<string> primaryKeysList = new List<string>();

                //Get Primary Key List
                #region Get Primary Key List for Each Table
                using (SqlConnection con1 = new SqlConnection(cs))
                {
                    string s = Constatnts.statement3;
                    SqlCommand cmd = new SqlCommand(s, con1);
                    con1.Open();
                    SqlDataReader reader;
                    cmd.Parameters.AddWithValue("@As_Table_NAME", t.TableName);
                    reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        string colName = reader["COLUMN_NAME"].ToString();
                        primaryKeysList.Add(colName);
                    }
                }
                #endregion

                using (SqlConnection con = new SqlConnection(cs))
                {
                    SqlCommand cmd = new SqlCommand(statement, con);
                    con.Open();
                    SqlDataReader reader;
                    cmd.Parameters.AddWithValue("@As_Table_NAME", t.TableName);
                    reader = cmd.ExecuteReader();
                    List<TableMetaData> listOfTMD = new List<TableMetaData>();
                    while (reader.Read())
                    {
                        TableMetaData tmd = new TableMetaData();
                        tmd.TableName = t.TableName;
                        tmd.ColumnName = reader["COLUMN_NAME"].ToString();
                        tmd.Id = reader["ORDINAL_POSITION"].ToString();
                        tmd.PK = reader["IS_NULLABLE"].ToString();
                        tmd.DataType = reader["DATA_TYPE"].ToString();
                        tmd.Length = reader["LENGTH"].ToString();
                        tmd.Precision = reader["PRECISION"].ToString();
                        tmd.Scale = reader["SCALE"].ToString();
                        tmd.Identity = reader["IsIdentity"].ToString();
                        listOfTMD.Add(tmd);
                    }
                    //Add para to show table Name and Synonym Name
                    Word.Paragraph oPara3;
                    oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                    oPara3 = oDoc.Content.Paragraphs.Add(ref oRng);
                    oPara3.Range.ParagraphFormat.SpaceAfter = 6;
                    oPara3.Range.Text = $@"Table Name: {t.TableName}        Synonym Name: {t.SynonymName}";

                    wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                    oTable = oDoc.Tables.Add(wrdRng, listOfTMD.Count + 1, 8, ref oMissing, ref oMissing);
                    oTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    oTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    oTable.Range.ParagraphFormat.SpaceAfter = 6;
                    oTable.Range.Font.Bold = 1;
                    oTable.Cell(1, 1).Range.Text = "ColumnName";
                    oTable.Cell(1, 2).Range.Text = "Id";
                    oTable.Cell(1, 3).Range.Text = "PK";
                    oTable.Cell(1, 4).Range.Text = "DataType";
                    oTable.Cell(1, 5).Range.Text = "Length";
                    oTable.Cell(1, 6).Range.Text = "Precision";
                    oTable.Cell(1, 7).Range.Text = "Scale";
                    oTable.Cell(1, 8).Range.Text = "IsIdentity";
                    oTable.Cell(1, 1).Range.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray40;
                    oTable.Cell(1, 2).Range.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray40;
                    oTable.Cell(1, 3).Range.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray40;
                    oTable.Cell(1, 4).Range.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray40;
                    oTable.Cell(1, 5).Range.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray40;
                    oTable.Cell(1, 6).Range.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray40;
                    oTable.Cell(1, 7).Range.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray40;
                    oTable.Cell(1, 8).Range.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray40;
                    oTable.Rows[1].Range.Font.Bold = 1;
                    oTable.Columns[1].PreferredWidth = oWord.InchesToPoints(2.4f);
                    oTable.Columns[2].PreferredWidth = oWord.InchesToPoints(0.35f);
                    oTable.Columns[3].PreferredWidth = oWord.InchesToPoints(0.42f);
                    oTable.Columns[4].PreferredWidth = oWord.InchesToPoints(0.9f);
                    oTable.Columns[5].PreferredWidth = oWord.InchesToPoints(0.7f);
                    oTable.Columns[6].PreferredWidth = oWord.InchesToPoints(0.75f);
                    oTable.Columns[7].PreferredWidth = oWord.InchesToPoints(0.7f);
                    oTable.Columns[8].PreferredWidth = oWord.InchesToPoints(0.75f);
                    wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                    wrdRng.InsertParagraphAfter();
                    wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                    wrdRng.InsertParagraphAfter();
                    oTable.Range.Font.Bold = 0;
                    int q = 0;
                    for (r = 2; r <= listOfTMD.Count + 1; r++)
                    {
                        oTable.Cell(r, 1).Range.Text = listOfTMD[q].ColumnName;
                        oTable.Cell(r, 2).Range.Text = listOfTMD[q].Id;
                        if (primaryKeysList.Contains(listOfTMD[q].ColumnName))
                        {
                            oTable.Cell(r, 3).Range.Text = "1";
                        }
                        else
                        {
                            oTable.Cell(r, 3).Range.Text = "0";
                        }
                        oTable.Cell(r, 4).Range.Text = listOfTMD[q].DataType;
                        oTable.Cell(r, 5).Range.Text = listOfTMD[q].Length;
                        oTable.Cell(r, 6).Range.Text = listOfTMD[q].Precision;
                        oTable.Cell(r, 7).Range.Text = listOfTMD[q].Scale;
                        oTable.Cell(r, 8).Range.Text = listOfTMD[q].Identity;
                        q++;
                    }
                    oTable.Rows[1].Range.Font.Bold = 1;
                }
            }
            string fileName = textBox_OutputPath.Text == "" ? @"C:" : textBox_OutputPath.Text;
            fileName = textBox_OutputPath.Text + @"\PDM";
            oDoc.SaveAs2(fileName);
            this.Close();
        }

        private void folderBrowserDialog1_HelpRequest(object sender, EventArgs e)
        {

        }

        private void Browse_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.ShowDialog();
            string selectedPath = folderBrowserDialog1.SelectedPath;
            if (selectedPath == "") {
                MessageBox.Show("Please enter a valid path to save the generated file.");
            }
            textBox_OutputPath.Text = folderBrowserDialog1.SelectedPath;
        }

        
    }
}
