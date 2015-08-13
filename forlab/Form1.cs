using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;

namespace forlab
{
    public partial class Form1 : Form
    {
        float[] myArray = new float[20];
        float[] deviationArray = new float[20];
        float[] squareDeviationArray = new float[20];
        float[] contrArray = new float[6];
        float allsum, otkl, avg, disp;
        public Form1()
        {
            InitializeComponent();
        }

        private void addButton_Click(object sender, EventArgs e)
        {
            Calculation Start = new Calculation();
            int iter = 1;
            float Result;
            
/*            var textArea = new List<float>();
            
              textArea.Add(System.Convert.ToSingle(textBox1.Text));
              textArea.Add(System.Convert.ToSingle(textBox2.Text));
              textArea.Add(System.Convert.ToSingle(textBox3.Text));
              textArea.Add(System.Convert.ToSingle(textBox4.Text));
              textArea.Add(System.Convert.ToSingle(textBox5.Text));
              textArea.Add(System.Convert.ToSingle(textBox6.Text));
              textArea.Add(System.Convert.ToSingle(textBox7.Text));
              textArea.Add(System.Convert.ToSingle(textBox8.Text));
              textArea.Add(System.Convert.ToSingle(textBox9.Text));
              textArea.Add(System.Convert.ToSingle(textBox10.Text));
              textArea.Add(System.Convert.ToSingle(textBox11.Text));
              textArea.Add(System.Convert.ToSingle(textBox12.Text));
              textArea.Add(System.Convert.ToSingle(textBox13.Text));
              textArea.Add(System.Convert.ToSingle(textBox14.Text));
              textArea.Add(System.Convert.ToSingle(textBox15.Text));
              textArea.Add(System.Convert.ToSingle(textBox16.Text));
              textArea.Add(System.Convert.ToSingle(textBox17.Text));
              textArea.Add(System.Convert.ToSingle(textBox18.Text));
              textArea.Add(System.Convert.ToSingle(textBox19.Text));
              textArea.Add(System.Convert.ToSingle(textBox20.Text));
  */            
            
 //           float[] squareDeviationArray = new float[20];
               
            //List<TextBox> tB = new List<TextBox>();
           /* foreach (Control ctl in this.Controls)
            {
                if (ctl is TextBox) tB.Add((TextBox)ctl);
            }*/

           // IEnumerable <>
            /* for (int i = 0; i < 20; i++)
           {
                int j = 1;
                myArray[i] = System.Convert.ToSingle(textBox1. [j].Text);
                j++;
            }*/
            
            myArray[0] = System.Convert.ToSingle(textBox1.Text);
            myArray[1] = System.Convert.ToSingle(textBox2.Text);
            myArray[2] = System.Convert.ToSingle(textBox3.Text);
            myArray[3] = System.Convert.ToSingle(textBox4.Text);
            myArray[4] = System.Convert.ToSingle(textBox5.Text);
            myArray[5] = System.Convert.ToSingle(textBox6.Text);
            myArray[6] = System.Convert.ToSingle(textBox7.Text);
            myArray[7] = System.Convert.ToSingle(textBox8.Text);
            myArray[8] = System.Convert.ToSingle(textBox9.Text);
            myArray[9] = System.Convert.ToSingle(textBox10.Text);
            myArray[10] = System.Convert.ToSingle(textBox11.Text);
            myArray[11] = System.Convert.ToSingle(textBox12.Text);
            myArray[12] = System.Convert.ToSingle(textBox13.Text);
            myArray[13] = System.Convert.ToSingle(textBox14.Text);
            myArray[14] = System.Convert.ToSingle(textBox15.Text);
            myArray[15] = System.Convert.ToSingle(textBox16.Text);
            myArray[16] = System.Convert.ToSingle(textBox17.Text);
            myArray[17] = System.Convert.ToSingle(textBox18.Text);
            myArray[18] = System.Convert.ToSingle(textBox19.Text);
            myArray[19] = System.Convert.ToSingle(textBox20.Text);
            
            Result = Start.AverageCalculation(myArray);
            avg = Result;
            allsum = myArray.Sum();
            label3.Text = System.Convert.ToString(Result);
            Result = Start.Dispersion(myArray,ref deviationArray,ref squareDeviationArray);
            disp = Result;
            for (int i = 0; i < 20; i++)
            {
                dataGridView1[0, i].Value = iter++;
                dataGridView1[1, i].Value = myArray[i];
                dataGridView1[2, i].Value = deviationArray[i];
                dataGridView1[3, i].Value = squareDeviationArray[i]; 
            }
            otkl = Result;
            label5.Text = System.Convert.ToString(Result);
            Result = Start.FactorOfVariation();
            label7.Text = System.Convert.ToString(Result);
            System.Array.Sort(myArray);
            Result = Start.TmaxCalculation(myArray);
            if (Result < 2.62)

                label9.Text = System.Convert.ToString(Result);

            else

                label9.Text = "Максимальный критерий превышает границу в 2.62";
            
            Result = Start.TminCalculation(myArray);
            if (Result < 2.62)
                label11.Text = System.Convert.ToString(Result);
            else
                label11.Text = "Минимальный критерий превышает границу в 2.62";
            Start.ContrMap(ref contrArray);
            label19.Text = System.Convert.ToString(contrArray[0]);
            label20.Text = System.Convert.ToString(contrArray[1]);
            label21.Text = System.Convert.ToString(contrArray[2]);
            label22.Text = System.Convert.ToString(contrArray[3]);
            label23.Text = System.Convert.ToString(contrArray[4]);
            label24.Text = System.Convert.ToString(contrArray[5]);
            
        }

        public void AddRows(DataGridView dgw)
        {
            //добавляет m строк в элемент управления dgw       
            //Заполнение DGView строками
            for (int i = 0; i < 19; i++)
            {
                dgw.Rows.Add();
                dgw.Rows[i].HeaderCell.Value
                    = "row" + i.ToString();
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            AddRows(dataGridView1);
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            object oMissing = System.Reflection.Missing.Value;
            object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */

            //Start Word and create a new document.
            Word._Application oWord;
            Word._Document oDoc;
            oWord = new Word.Application();
            oWord.Visible = true;
            oDoc = oWord.Documents.Add(ref oMissing, ref oMissing,
                ref oMissing, ref oMissing);

            //Insert a paragraph at the beginning of the document.
            Word.Paragraph oPara1;
            oPara1 = oDoc.Content.Paragraphs.Add(ref oMissing);
            oPara1.Range.Text = "Внутрішньолабораторній контроль якості";
            oPara1.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
            oPara1.Range.Font.Bold = 1;
            oPara1.Format.SpaceAfter = 12;    //24 pt spacing after paragraph.
            oPara1.Range.InsertParagraphAfter();

            //Insert a paragraph at the end of the document.
            Word.Paragraph oPara2;
            object oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oPara2 = oDoc.Content.Paragraphs.Add(ref oRng);
            oPara2.Range.Font.Bold = 0;
            oPara2.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
            oPara2.Range.Text = "Лабораторія - ЦКДЛ КЗ\"Дніпропетровський клінічний онкологічний центр\" Дніпропетровської обласної ради";
            oPara2.Range.ParagraphFormat.SpaceAfter = 1;
            oPara2.Format.SpaceAfter = 3;
            oPara2.Range.InsertParagraphAfter();
            
            //Insert another paragraph.
            Word.Paragraph oPara3;
            oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oPara3 = oDoc.Content.Paragraphs.Add(ref oRng);
            
            oPara3.Range.Text = "Показник, що визначається:________________________";
            oPara3.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;
            oPara3.Range.ParagraphFormat.SpaceAfter = 1;
            oPara3.Range.InsertParagraphAfter();
            oPara3.Range.Font.Bold = 0;
            oPara3.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;
            oPara3.Range.Text = "Методика:________________________";
            oPara3.Range.ParagraphFormat.SpaceAfter = 1;
            oPara3.Range.InsertParagraphAfter();
            oPara3.Range.Font.Bold = 0;
            oPara3.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;
            oPara3.Range.Text = "Одиниці вимірювання________ Прилад _________________________________________";
            oPara3.Range.ParagraphFormat.SpaceAfter = 1;
            oPara3.Range.InsertParagraphAfter();
            //oPara3.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
            oPara2.Range.Font.Bold = 0;
            oPara1.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
            oPara2.Range.Text = "Побудова контрольноъ карти індивідуальних значень";
            oPara3.Range.ParagraphFormat.SpaceAfter = 1;
            //oPara3.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;
            oPara3.Range.InsertParagraphAfter();
            oPara3.Range.Font.Bold = 0;
            oPara3.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;
            oPara3.Range.Text = "Контрольний матеріал________ Серія_________________________________________";
            oPara3.Range.ParagraphFormat.SpaceAfter = 1;
            oPara3.Range.InsertParagraphAfter();
            oPara3.Range.Font.Bold = 0;
            oPara3.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;
            oPara3.Range.Text = "Термін використання __________________";
            oPara3.Range.ParagraphFormat.SpaceAfter = 1;
            //oPara3.Range.Font.Bold = 0;
            oPara3.Range.Font.Bold = 0;
            oPara3.Format.SpaceAfter = 3;
            oPara3.Format.SpaceAfter = 12;
            oPara3.Range.InsertParagraphAfter();

            //Insert a 3 x 5 table, fill it with data, and make the first row
            //bold and italic.
            Word.Table oTable;
            Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oTable = oDoc.Tables.Add(wrdRng, 22, 5, ref oMissing, ref oMissing);
            oTable.Range.ParagraphFormat.SpaceAfter = 2;
            
            int r, c, i; c = 1; i = 0;
            string strText;
            oTable.Range.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            oTable.Range.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
         //   oTable.Borders.InsideLineWidth = Word.WdLineWidth.wdLineWidth050pt;
           // oTable.Range.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleDashDot;
            oTable.Cell(1, 1).Range.Text = "Дата";
            oTable.Cell(1, 2).Range.Text = "Номер";
            oTable.Cell(1, 3).Range.Text = "Полученные значения Хі";
            oTable.Cell(1, 4).Range.Text = "Отклонение от среднего d";
            oTable.Cell(1, 5).Range.Text = "Квадрат отклонения";
            for (r = 2; r <= 21; r++, i++, c++)
            {
                strText = System.Convert.ToString(c);
                oTable.Cell(r, 2).Range.Text = strText;
                
                strText = System.Convert.ToString(myArray[i]);
                oTable.Cell(r, 3).Range.Text = strText;
                
                strText = System.Convert.ToString(deviationArray[i]);
                oTable.Cell(r, 4).Range.Text = strText;
                
                strText = System.Convert.ToString(squareDeviationArray[i]);
                oTable.Cell(r, 5).Range.Text = strText;
             }
            oTable.Cell(22, 1).Range.Text = "";
            oTable.Cell(22, 2).Range.Text = "n=20";
            oTable.Cell(22, 3).Range.Text = allsum.ToString();
            oTable.Cell(22, 4).Range.Text = "";
            oTable.Cell(22, 5).Range.Text = otkl.ToString();
            //oTable.Range.ParagraphFormat.SpaceAfter = 6;
            Word.Paragraph oPara4;
            oPara4 = oDoc.Content.Paragraphs.Add(ref oMissing);
            oPara4.Range.ParagraphFormat.SpaceAfter = 1;
            oPara4.Range.Text = "x = "+ avg + "                             S = " + disp;
            oPara4.Range.ParagraphFormat.SpaceAfter = 1;
            oPara4.Range.InsertParagraphAfter();
            oPara4.Range.Text = "Перевірка на можливість вилучення__________________________________________";
            oPara4.Range.ParagraphFormat.SpaceAfter = 1;
            oPara4.Range.InsertParagraphAfter();
            oPara4.Range.Text = "Розрахунок коефіцієнта варіації: CV = " + disp*avg*100 + "%";
            oPara4.Range.ParagraphFormat.SpaceAfter = 1;
            oPara4.Range.InsertParagraphAfter();
            oPara4.Range.Text = "Розрахунок значень для побудови контрольної карти:";
            oPara4.Range.ParagraphFormat.SpaceAfter = 1;
            oPara4.Range.InsertParagraphAfter();
            oPara4.Range.Text = "x + 1S = " + System.Convert.ToString(contrArray[0]) + ";                                                                                          x - 1S = " + System.Convert.ToString(contrArray[3]);
            oPara4.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphDistribute;
            oPara4.Range.ParagraphFormat.SpaceAfter = 1;
            oPara4.Range.InsertParagraphAfter();
            oPara4.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphDistribute;
            oPara4.Range.Text = "x + 2S = " + System.Convert.ToString(contrArray[1]) + ";                                                                                          x - 2S = " + System.Convert.ToString(contrArray[4]);
            oPara4.Range.ParagraphFormat.SpaceAfter = 1;
            oPara4.Range.InsertParagraphAfter();
            oPara4.Range.Text = "x + 3S = " + System.Convert.ToString(contrArray[2]) + ";                                                                                          x - 3S = " + System.Convert.ToString(contrArray[5]);
            oPara4.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphDistribute;
            oPara4.Range.ParagraphFormat.SpaceAfter = 1;
            oPara4.Range.InsertParagraphAfter();
            oPara4.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphDistribute;
            oPara4.Range.ParagraphFormat.SpaceBefore = 10; 
            oPara4.Range.Text = "Дата _________________                                                    Підпис _________________";




/*            
            oTable.Rows[1].Range.Font.Bold = 0;
            oTable.Rows[1].Range.Font.Italic = 1;

            //Add some text after the table.
            Word.Paragraph oPara4;
            oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oPara4 = oDoc.Content.Paragraphs.Add(ref oRng);
            oPara4.Range.InsertParagraphBefore();
            oPara4.Range.Text = "And here's another table:";
            oPara4.Format.SpaceAfter = 24;
            oPara4.Range.InsertParagraphAfter();

            //Insert a 5 x 2 table, fill it with data, and change the column widths.
            wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oTable = oDoc.Tables.Add(wrdRng, 5, 2, ref oMissing, ref oMissing);
            oTable.Range.ParagraphFormat.SpaceAfter = 6;
            for (r = 1; r <= 5; r++)
                for (c = 1; c <= 2; c++)
                {
                    strText = "r" + r + "c" + c;
                    oTable.Cell(r, c).Range.Text = strText;
                }
            oTable.Columns[1].Width = oWord.InchesToPoints(2); //Change width of columns 1 & 2
            oTable.Columns[2].Width = oWord.InchesToPoints(3);

            //Keep inserting text. When you get to 7 inches from top of the
            //document, insert a hard page break.
            object oPos;
            double dPos = oWord.InchesToPoints(7);
            oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range.InsertParagraphAfter();
            do
            {
                wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                wrdRng.ParagraphFormat.SpaceAfter = 6;
                wrdRng.InsertAfter("A line of text");
                wrdRng.InsertParagraphAfter();
                oPos = wrdRng.get_Information
                               (Word.WdInformation.wdVerticalPositionRelativeToPage);
            }
            while (dPos >= Convert.ToDouble(oPos));
            object oCollapseEnd = Word.WdCollapseDirection.wdCollapseEnd;
            object oPageBreak = Word.WdBreakType.wdPageBreak;
            wrdRng.Collapse(ref oCollapseEnd);
            wrdRng.InsertBreak(ref oPageBreak);
            wrdRng.Collapse(ref oCollapseEnd);
            wrdRng.InsertAfter("We're now on page 2. Here's my chart:");
            wrdRng.InsertParagraphAfter();

            //Insert a chart.
            Word.InlineShape oShape;
            object oClassType = "MSGraph.Chart.8";
            wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oShape = wrdRng.InlineShapes.AddOLEObject(ref oClassType, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing);

            //Demonstrate use of late bound oChart and oChartApp objects to
            //manipulate the chart object with MSGraph.
            object oChart;
            object oChartApp;
            oChart = oShape.OLEFormat.Object;
            oChartApp = oChart.GetType().InvokeMember("Application",
                BindingFlags.GetProperty, null, oChart, null);

            //Change the chart type to Line.
            object[] Parameters = new Object[1];
            Parameters[0] = 4; //xlLine = 4
            oChart.GetType().InvokeMember("ChartType", BindingFlags.SetProperty,
                null, oChart, Parameters);

            //Update the chart image and quit MSGraph.
            oChartApp.GetType().InvokeMember("Update",
                BindingFlags.InvokeMethod, null, oChartApp, null);
            oChartApp.GetType().InvokeMember("Quit",
                BindingFlags.InvokeMethod, null, oChartApp, null);
            //... If desired, you can proceed from here using the Microsoft Graph 
            //Object model on the oChart and oChartApp objects to make additional
            //changes to the chart.

            //Set the width of the chart.
            oShape.Width = oWord.InchesToPoints(6.25f);
            oShape.Height = oWord.InchesToPoints(3.57f);

            //Add text after the chart.
            wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            wrdRng.InsertParagraphAfter();
            wrdRng.InsertAfter("THE END.");
*/
            //Close this form.
            this.Close();
        }

        private void Close_Click(object sender, EventArgs e)
        {
          //  this.Close();
        }
    }
}
