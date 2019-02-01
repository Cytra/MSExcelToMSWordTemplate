using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Resources;
using System.Text;
using System.Threading.Tasks;
using Task = Microsoft.Office.Interop.Word.Task;


namespace ZivileWpfApp.ViewModels
{
    public static class MSWordFuncions
    {
        public static void GenerateWordFile(string no, string biNo, string moKo, string skoVaPa, string AsKo, string addres, string skoDyd, string skolSusiLaik, string SutarDa, string adresPagalGr, string Islaid, string folderLocation)
        {
            object oMissing = System.Reflection.Missing.Value;
            object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */

            //Start Word and create a new document.
            _Application oWord;
            _Document oDoc;
            oWord = new Application();
            oWord.Visible = false;
            oDoc = oWord.Documents.Add(ref oMissing, ref oMissing,
            ref oMissing, ref oMissing);

            //Create Horizontal Page
            oDoc.PageSetup.TogglePortrait();
            oDoc.PageSetup.TopMargin = 20;
            oDoc.PageSetup.BottomMargin = 20;
            oDoc.PageSetup.RightMargin = 30;
            oDoc.PageSetup.LeftMargin = 30;

            ////Insert a paragraph at the beginning of the document.
            //Paragraph oPara1;
            //oPara1 = oDoc.Content.Paragraphs.Add(ref oMissing);
            //oPara1.Range.Text = "Heading 1";
            //oPara1.Range.Font.Bold = 1;
            //oPara1.Format.SpaceAfter = 24;    //24 pt spacing after paragraph.
            //oPara1.Range.InsertParagraphAfter();

            ////Insert a paragraph at the end of the document.
            //Paragraph oPara2;
            //object oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            //oPara2 = oDoc.Content.Paragraphs.Add(ref oRng);
            //oPara2.Range.Text = "Heading 2";
            //oPara2.Format.SpaceAfter = 6;
            //oPara2.Range.InsertParagraphAfter();

            ////Insert another paragraph.
            Paragraph oPara1;
            object oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oPara1 = oDoc.Content.Paragraphs.Add(ref oRng);
            oPara1.Range.Text = "Išrašas iš Bylų perdavimo teisminiam išieškojimui akto Nr. 15";
            oPara1.Range.Font.Bold = 1;
            oPara1.Format.SpaceAfter = 0;
            oPara1.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            oPara1.Range.Font.Size = 12;
            oPara1.Range.Font.Name = "Times New Roman";
            oPara1.Range.LanguageID = WdLanguageID.wdLithuanian;
            oPara1.Range.InsertParagraphAfter();

            Paragraph oPara2;
            oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oPara2 = oDoc.Content.Paragraphs.Add(ref oRng);
            oPara2.Range.Text = "2019 m. sausio 31 d.";
            oPara2.Range.Font.Bold = 0;
            oPara2.Format.SpaceAfter = 5;
            oPara2.Range.InsertParagraphAfter();

            Paragraph oPara3;
            oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oPara3 = oDoc.Content.Paragraphs.Add(ref oRng);
            oPara3.Range.Text = "Bylų perdavimo teisminiam išieškojimui aktas Nr. 15";
            oPara3.Range.Font.Bold = 1;
            oPara3.Format.SpaceAfter = 0;
            oPara3.Range.InsertParagraphAfter();

            Paragraph oPara4;
            oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oPara4 = oDoc.Content.Paragraphs.Add(ref oRng);
            oPara4.Range.Text = "2018 m. gruodžio 5 d.";
            oPara4.Range.Font.Bold = 0;
            oPara4.Format.SpaceAfter = 12;
            oPara4.Range.InsertParagraphAfter();

            Paragraph oPara5;
            oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oPara5 = oDoc.Content.Paragraphs.Add(ref oRng);
            oPara5.Range.Text = "Šalys, pasirašydamos šį aktą, patvirtina, kad remiantis 2017 m. gruodžio 1 d. Teisinių paslaugų sutartimi Nr. 03-860 (30.01), Klientas perduoda, o Advokatų profesinė bendrija „Baublienė, Goliančik ir partneriai“ priima šias bylas teisminiam išieškojimui:";
            oPara5.Range.Font.Bold = 0;
            oPara5.Format.SpaceAfter = 10;
            oPara5.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            oPara5.Range.InsertParagraphAfter();

            Paragraph oPara6;
            oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oPara6 = oDoc.Content.Paragraphs.Add(ref oRng);
            oPara6.Range.Text = "Su pareiškimais dėl teismo įsakymo išdavimo:";
            oPara6.Range.Font.Bold = 1;
            oPara6.Format.SpaceAfter = 12;
            oPara6.Range.InsertParagraphAfter();

            //Insert a 3 x 5 table, fill it with data, and make the first row
            //bold and italic.
            Table oTable;
            Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;

            var columns = 11;
            var rows = 2;
            oTable = oDoc.Tables.Add(wrdRng, rows, columns, ref oMissing, ref oMissing);
            oTable.Range.ParagraphFormat.SpaceAfter = 6;
            int r, c;

            for (r = 1; r <= rows; r++)
                for (c = 1; c <= columns; c++)
                {
                    if (r == 1 && c == 1)
                    {
                        oTable.Cell(r, c).Range.Text = "Eilės Nr.";
                    }
                    if (r == 1 && c == 2)
                    {
                        oTable.Cell(r, c).Range.Text = "Vidinės bylos Nr.";
                    }
                    if (r == 1 && c == 3)
                    {
                        oTable.Cell(r, c).Range.Text = "Mokėtojo kodas";
                    }
                    if (r == 1 && c == 4)
                    {
                        oTable.Cell(r, c).Range.Text = "Skolininko Vardas Pavardė";
                    }
                    if (r == 1 && c == 5)
                    {
                        oTable.Cell(r, c).Range.Text = "Asmens kodas/Gimimo data";
                    }
                    if (r == 1 && c == 6)
                    {
                        oTable.Cell(r, c).Range.Text = "Adresas, kur susidarė skola";
                    }
                    if (r == 1 && c == 7)
                    {
                        oTable.Cell(r, c).Range.Text = "Skolos dydis už vandenį";
                    }
                    if (r == 1 && c == 8)
                    {
                        oTable.Cell(r, c).Range.Text = "Skolos susidarymo laikotarpis";
                    }
                    if (r == 1 && c == 9)
                    {
                        oTable.Cell(r, c).Range.Text = "Sutarties data ir Nr. (jei sutartis buvo sudaryta)";
                    }
                    if (r == 1 && c == 10)
                    {
                        oTable.Cell(r, c).Range.Text = "Adresas, pagal GRT išrašą";
                    }
                    if (r == 1 && c == 11)
                    {
                        oTable.Cell(r, c).Range.Text = "Išlaidos už teisines paslaugas, Eur + PVM";
                    }


                    if (r == 2 && c == 1)
                    {
                        oTable.Cell(r, c).Range.Text = no;
                    }
                    if (r == 2 && c == 2)
                    {
                        oTable.Cell(r, c).Range.Text = biNo;
                    }
                    if (r == 2 && c == 3)
                    {
                        oTable.Cell(r, c).Range.Text = moKo;
                    }
                    if (r == 2 && c == 4)
                    {
                        oTable.Cell(r, c).Range.Text = skoVaPa;
                    }
                    if (r == 2 && c == 5)
                    {
                        oTable.Cell(r, c).Range.Text = AsKo;
                    }
                    if (r == 2 && c == 6)
                    {
                        oTable.Cell(r, c).Range.Text = addres;
                    }
                    if (r == 2 && c == 7)
                    {
                        oTable.Cell(r, c).Range.Text = skoDyd;
                    }
                    if (r == 2 && c == 8)
                    {
                        oTable.Cell(r, c).Range.Text = skolSusiLaik;
                    }
                    if (r == 2 && c == 9)
                    {
                        oTable.Cell(r, c).Range.Text = SutarDa;
                    }
                    if (r == 2 && c == 10)
                    {
                        oTable.Cell(r, c).Range.Text = adresPagalGr;
                    }
                    if (r == 2 && c == 11)
                    {
                        oTable.Cell(r, c).Range.Text = Islaid;
                    }

                    oTable.Cell(r, c).Range.Borders[WdBorderType.wdBorderLeft].LineStyle = WdLineStyle.wdLineStyleSingle;
                    oTable.Cell(r, c).Range.Borders[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleSingle;
                    oTable.Cell(r, c).Range.Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleSingle;
                    oTable.Cell(r, c).Range.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
                    oTable.Cell(r, c).Range.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    //string strText = "r" + r + "c" + c;

                }
            oTable.Rows[1].Range.Font.Bold = 1;
            oTable.Rows[1].Alignment = WdRowAlignment.wdAlignRowCenter;
            oTable.Rows[1].Range.Font.Size = 9;
            oTable.Rows[2].Range.Font.Bold = 0;
            oTable.Rows[2].Range.Font.Size = 9;

            Paragraph oPara7;
            oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oPara7 = oDoc.Content.Paragraphs.Add(ref oRng);
            //oPara7.Range.Text = "Su pareiškimais dėl teismo įsakymo išdavimo:";
            //oPara7.Range.Font.Bold = 1;
            //oPara7.Format.SpaceAfter = 2;
            //oPara7.Range.InsertParagraphAfter();

            //Insert a chart.
            InlineShape oShape;
            wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            //wrdRng.InsertParagraphBefore();
            oShape = wrdRng.InlineShapes.AddPicture(@"C:\Users\bc1729\Desktop\C sharp personal projects\ZivileWpfApp\ZivileWpfApp\Resources\Rekvizitai.png");
            //oShape.ScaleHeight = 100;
            oShape.ScaleWidth = 115;



            //oPara6 = oDoc.Content.Paragraphs.Add(ref oRng);
            //oPara6.Range.Text = "Su pareiškimais dėl teismo įsakymo išdavimo:";
            //oPara6.Range.Font.Bold = 1;
            //oPara6.Format.SpaceAfter = 12;
            //oPara6.Range.InsertParagraphAfter();



            //Add some text after the table.
            //Paragraph oPara6;
            //oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            //oPara6 = oDoc.Content.Paragraphs.Add(ref oRng);
            //oPara6.Range.InsertParagraphBefore();
            //oPara6.Range.Text = "And here's another table:";
            //oPara6.Format.SpaceAfter = 24;
            //oPara6.Range.InsertParagraphAfter();

            //Insert a 5 x 2 table, fill it with data, and change the column widths.
            //wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            //oTable = oDoc.Tables.Add(wrdRng, 5, 2, ref oMissing, ref oMissing);
            //oTable.Range.ParagraphFormat.SpaceAfter = 6;
            //for (r = 1; r <= 5; r++)
            //    for (c = 1; c <= 2; c++)
            //    {
            //        strText = "r" + r + "c" + c;
            //        oTable.Cell(r, c).Range.Text = strText;
            //    }
            //oTable.Columns[1].Width = oWord.InchesToPoints(2); //Change width of columns 1 & 2
            //oTable.Columns[2].Width = oWord.InchesToPoints(3);

            ////Keep inserting text. When you get to 7 inches from top of the
            ////document, insert a hard page break.
            //object oPos;
            //double dPos = oWord.InchesToPoints(7);
            //oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range.InsertParagraphAfter();
            //do
            //{
            //    wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            //    wrdRng.ParagraphFormat.SpaceAfter = 6;
            //    wrdRng.InsertAfter("A line of text");
            //    wrdRng.InsertParagraphAfter();
            //    oPos = wrdRng.get_Information
            //                           (WdInformation.wdVerticalPositionRelativeToPage);
            //}
            //while (dPos >= Convert.ToDouble(oPos));
            //object oCollapseEnd = WdCollapseDirection.wdCollapseEnd;
            //object oPageBreak = WdBreakType.wdPageBreak;
            //wrdRng.Collapse(ref oCollapseEnd);
            //wrdRng.InsertBreak(ref oPageBreak);
            //wrdRng.Collapse(ref oCollapseEnd);
            //wrdRng.InsertAfter("We're now on page 2. Here's my chart:");
            //wrdRng.InsertParagraphAfter();

            //Insert a chart.
            //InlineShape oShape;
            //object oClassType = "MSGraph.Chart.8";
            //wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            //oShape = wrdRng.InlineShapes.AddOLEObject(ref oClassType, ref oMissing,
            //ref oMissing, ref oMissing, ref oMissing,
            //ref oMissing, ref oMissing, ref oMissing);

            //Demonstrate use of late bound oChart and oChartApp objects to
            //manipulate the chart object with MSGraph.
            //object oChart;
            //object oChartApp;
            //oChart = oShape.OLEFormat.Object;


            ////Change the chart type to Line.
            //object[] Parameters = new Object[1];
            //Parameters[0] = 4; //xlLine = 4


            ////Update the chart image and quit MSGraph.

            ////... If desired, you can proceed from here using the Microsoft Graph 
            ////Object model on the oChart and oChartApp objects to make additional
            ////changes to the chart.

            ////Set the width of the chart.
            //oShape.Width = oWord.InchesToPoints(6.25f);
            //oShape.Height = oWord.InchesToPoints(3.57f);

            //Add text after the chart.
            //wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            //wrdRng.InsertParagraphAfter();
            //wrdRng.InsertAfter("THE END.");

            //SAVE DOCUMENT
            //oDoc.SaveAs(@"H:\NewDocument.docx");

            var name = skoVaPa.Replace(" ", "_");
            name = skoVaPa.Replace("\"", "_");
            name = skoVaPa.Replace(@"\n", "_");
            oDoc.SaveAs2($"{ folderLocation}//Išrašas iš Bylų perdavimo teisminiam išieškojimui akto_{name.Replace("/", "-")}.docx");

            ;

            oDoc.Close();
            oDoc = null;

            oWord.Quit();
            oWord = null;


        }
    }
}
