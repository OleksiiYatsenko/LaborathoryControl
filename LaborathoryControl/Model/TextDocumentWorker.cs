using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LaborathoryControl.Enum;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Win32;
using System.Runtime.InteropServices;

namespace LaborathoryControl.Model
{
    public class TextDocumentWorker
    {
        private TextRedactors _textRedactor;
        private List<Data> data;
        private Calculation calc;


        public TextDocumentWorker(IEnumerable<Data> values, Calculation calc)
        {
            data = new List<Data>(values);
            this.calc = calc;
            _textRedactor = WhatsInstaled();
        }

        private TextRedactors WhatsInstaled()
        {
            using (RegistryKey microsoft = Registry.LocalMachine.OpenSubKey("Software").OpenSubKey("Microsoft"))
            {
                if (microsoft != null)
                {
                    RegistryKey word = microsoft.OpenSubKey("Word");

                    if (word != null)
                    {
                        return TextRedactors.Word;
                    }
                }
            }
            string baseKey;
            if (Marshal.SizeOf(typeof(IntPtr)) == 8) 
                baseKey = @"SOFTWARE\Wow6432Node\OpenOffice.org\";
            else
                baseKey = @"SOFTWARE\OpenOffice.org\";
            string key = baseKey + @"Layers\URE\1";

            RegistryKey OpenOffice = Registry.CurrentUser.OpenSubKey(key);
                if (OpenOffice == null)
                    OpenOffice = Registry.LocalMachine.OpenSubKey(key);
                string urePath = OpenOffice.GetValue("UREINSTALLLOCATION") as string;
                if (!string.IsNullOrEmpty(urePath))
                {
                    OpenOffice.Close();
                    return TextRedactors.OpenOffice;
                }
            return TextRedactors.None;
        }

        public void MakeDocument()
        {
            switch(_textRedactor)
            {
                case TextRedactors.Word:
                    {
                        WorkWithMsWord();
                        return;
                    }
                case TextRedactors.OpenOffice:
                    {
                        return;
                    }
                default:
                    return;
            }
        }

        private void WorkWithMsWord()
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

            double allsum = 0;
            
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
            for (int r = 2, i = 0; r <= 21; r++, i++)
            {
                strText = System.Convert.ToString(data[i].Number);
                oTable.Cell(r, 2).Range.Text = strText;

                strText = System.Convert.ToString(data[i].Value);
                oTable.Cell(r, 3).Range.Text = strText;

                strText = System.Convert.ToString(data[i].Deviation);
                oTable.Cell(r, 4).Range.Text = strText;

                strText = System.Convert.ToString(data[i].SquaredDeviation);
                oTable.Cell(r, 5).Range.Text = strText;

                allsum += data[i].Value;
            }
            oTable.Cell(22, 1).Range.Text = "";
            oTable.Cell(22, 2).Range.Text = "n=20";
            oTable.Cell(22, 3).Range.Text = allsum.ToString();
            oTable.Cell(22, 4).Range.Text = "";
            oTable.Cell(22, 5).Range.Text = calc.Variance.ToString();
            //oTable.Range.ParagraphFormat.SpaceAfter = 6;
            Word.Paragraph oPara4;
            oPara4 = oDoc.Content.Paragraphs.Add(ref oMissing);
            oPara4.Range.ParagraphFormat.SpaceAfter = 1;
            oPara4.Range.Text = "x = " + calc.Average.ToString() + "                             S = " + calc.Variance.ToString();
            oPara4.Range.ParagraphFormat.SpaceAfter = 1;
            oPara4.Range.InsertParagraphAfter();
            oPara4.Range.Text = "Перевірка на можливість вилучення__________________________________________";
            oPara4.Range.ParagraphFormat.SpaceAfter = 1;
            oPara4.Range.InsertParagraphAfter();
            oPara4.Range.Text = "Розрахунок коефіцієнта варіації: CV = " + calc.Variance * calc.Average * 100 + "%";
            oPara4.Range.ParagraphFormat.SpaceAfter = 1;
            oPara4.Range.InsertParagraphAfter();
            oPara4.Range.Text = "Розрахунок значень для побудови контрольної карти:";
            oPara4.Range.ParagraphFormat.SpaceAfter = 1;
            oPara4.Range.InsertParagraphAfter();
            oPara4.Range.Text = "x + 1S = " + System.Convert.ToString(calc.ContrArr[0]) + ";                                                                                          x - 1S = " + System.Convert.ToString(calc.ContrArr[3]);
            oPara4.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphDistribute;
            oPara4.Range.ParagraphFormat.SpaceAfter = 1;
            oPara4.Range.InsertParagraphAfter();
            oPara4.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphDistribute;
            oPara4.Range.Text = "x + 2S = " + System.Convert.ToString(calc.ContrArr[1]) + ";                                                                                          x - 2S = " + System.Convert.ToString(calc.ContrArr[4]);
            oPara4.Range.ParagraphFormat.SpaceAfter = 1;
            oPara4.Range.InsertParagraphAfter();
            oPara4.Range.Text = "x + 3S = " + System.Convert.ToString(calc.ContrArr[2]) + ";                                                                                          x - 3S = " + System.Convert.ToString(calc.ContrArr[5]);
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
        }

        private void WorkWithOpenOffice()
        {

        }

    }
}
