using System;
using System.Data;
using System.Windows.Forms;
using iTextSharp.text.pdf;
using iTextSharp.text;
using System.IO;

//Author : Shamseer(Dotnet Mob)
//Find Us On
//Export DataTable to Pdf using iTextSharp dll
//Blog : http://dotnetmob.com/
//Youtube Channel : https://www.youtube.com/c/DotnetMob
namespace ExportToPdf
{
    public partial class Form1 : Form
    {
        #region Functions
        public Form1()
        {
            InitializeComponent();
        }

        DataTable MakeDataTable()
        {
               //Create friend table object
               DataTable friend = new DataTable();

               //Define columns
               friend.Columns.Add("Sl. No");
               friend.Columns.Add("Name");
               friend.Columns.Add("Country");
               friend.Columns.Add("Region");

               //Populate with friends :)
               friend.Rows.Add("1", "Smith", "United State", "California");
               friend.Rows.Add("2", "Jack", "United Kingdom", "London");
               friend.Rows.Add("3", "Kajal Singh", "India", "Delhi");
               friend.Rows.Add("4", "Emmanuel", "South Africa", "Pretoria");
               friend.Rows.Add("5", "Lucas", "Germany", "Berlin");
               friend.Rows.Add("6", "Khalid", "UAE", "Dubai");
               friend.Rows.Add("7", "William", "Australia", "Canberra");

               return friend;
           } 
           #endregion



        #region Events
        void ExportDataTableToPdf(DataTable dtblTable, String strPdfPath, string strHeader)
        {
            System.IO.FileStream fs = new FileStream(strPdfPath, FileMode.Create, FileAccess.Write, FileShare.None);
            Document document = new Document();
            document.SetPageSize(iTextSharp.text.PageSize.A4);
            PdfWriter writer = PdfWriter.GetInstance(document, fs);
            document.Open();

            //Report Header
            BaseFont bfntHead = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            Font fntHead = new Font(bfntHead, 16, 1, Color.GRAY);
            Paragraph prgHeading = new Paragraph();
            prgHeading.Alignment = Element.ALIGN_CENTER;
            prgHeading.Add(new Chunk(strHeader.ToUpper(), fntHead));
            document.Add(prgHeading);

            //Author
            Paragraph prgAuthor = new Paragraph();
            BaseFont btnAuthor = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            Font fntAuthor = new Font(btnAuthor, 8, 2, Color.GRAY);
            prgAuthor.Alignment = Element.ALIGN_RIGHT;
            prgAuthor.Add(new Chunk("Author : Dotnet Mob", fntAuthor));
            prgAuthor.Add(new Chunk("\nRun Date : " + DateTime.Now.ToShortDateString(), fntAuthor));
            document.Add(prgAuthor);

            //Add a line seperation
            Paragraph p = new Paragraph(new Chunk(new iTextSharp.text.pdf.draw.LineSeparator(0.0F, 100.0F, Color.BLACK, Element.ALIGN_LEFT, 1)));
            document.Add(p);

            //Add line break
            document.Add(new Chunk("\n", fntHead));

            //Write the table
            PdfPTable table = new PdfPTable(dataGridView1.Columns.Count);
            //Table header
            BaseFont btnColumnHeader = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            Font fntColumnHeader = new Font(btnColumnHeader, 10, 1, Color.WHITE);
            for (int i = 0; i < dtblTable.Columns.Count; i++)
            {
                PdfPCell cell = new PdfPCell();
                cell.BackgroundColor = Color.GRAY;
                cell.AddElement(new Chunk(dtblTable.Columns[i].ColumnName.ToUpper(), fntColumnHeader));
                table.AddCell(cell);
            }
            //table Data
            for (int i = 0; i < dtblTable.Rows.Count; i++)
            {
                for (int j = 0; j < dtblTable.Columns.Count; j++)
                {
                    table.AddCell(dtblTable.Rows[i][j].ToString());
                }
            }

            document.Add(table);
            document.Close();
            writer.Close();
            fs.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                DataTable dtbl = MakeDataTable();
                ExportDataTableToPdf(dtbl, @"D:\5semproject\resources\ExportToPdf\ExportToPdf\test.pdf", "Friend List");
                if (cbxOpen.Checked)
                {
                    System.Diagnostics.Process.Start(@"D:\5semproject\resources\ExportToPdf\ExportToPdf\test.pdf");
                    this.WindowState = System.Windows.Forms.FormWindowState.Minimized;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error Message");
            }
        }
        #endregion

        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'databaseDataSet.Table' table. You can move, or remove it, as needed.
            this.tableTableAdapter.Fill(this.databaseDataSet.Table);

        }
    }
}
