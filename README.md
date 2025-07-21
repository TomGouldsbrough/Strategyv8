This code uses data taken from data logging software to create visual graphs of tyre degradation. It uses the xml data taken from the data logging software, then through user inputs it then knows the length of the stint and how many stints to compare against. This data is then extrapolated by creating an equation for the race distance. Future work will be to add in pit length time, degradation behind dirty air and lap times of opponents to create the best strategy. 
I show the entire code below for people to critique and give feedback on who do not want to download the coad.
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Windows.Forms.DataVisualization.Charting;
using Microsoft.VisualBasic;
using System.Drawing.Drawing2D;
using MathNet.Numerics.LinearAlgebra;
using MathNet.Numerics.LinearRegression;
using DocumentFormat.OpenXml.Office2019.Drawing.Model3D;
using System.Xml;
using DocumentFormat.OpenXml.Drawing.Charts;


namespace StrategyWithChartV6
{
    public partial class Form1 : Form
    {
        public Form1()
        {

            InitializeComponent();
            double Grad1 = 0; // Gradient for First equation
            double c1 = 0; // y intercept for first equation
            double Grad2 = 0; // Gradient for second equation
            double c2 = 0; // y intercept for second equation
            int RaceLaps = 0; // number of laps in the race
            int FirstLaps = 0; // number of laps completed for the first test
            int SecondLaps = 0; // number of laps completed for second test
            List<double> FinalFirstStrat = new List<double>(); // List of the points in the frst strategy
            List<double> SecondStratFinal = new List<double>(); // list of the points in the second strategy
            List<string> Strategy = new List<string>(); // list of the final strategy.
            List<double> FirstTest = new List<double>(); // List of all points in the first test
            List<double> SecondTest = new List<double>();// list of all point in the second test 
            List<double> FirstLapsLength = new List<double>(); // length of the laps of the first test
            List<double> SecondLapsLength = new List<double>(); // length of the laps of the second test.

            List<double> FinalStratQuadList = new List<double>();
            List<double> SecondStratQuadList = new List<double>(); // list for quadtratics

            List<double> QuadRaceLapsFirst = new List<double>();
            List<double> QuadRaceLapsSecond = new List<double>();
            double a1 = 0;
            double b1 = 0;
            double c1_quad = 0;

            double a2 = 0;
            double b2 = 0;
            double c2_quad = 0;

            double sumX = 0;
            double sumY = 0;
            double sumXY = 0;
            double sumX2 = 0;

            double sumXS = 0;
            double sumYS = 0;
            double sumXYS = 0;
            double sumX2S = 0;





            double doubleValue = 0;
            double doubleValue2 = 0;
            int Laps = 0;



            //xml extraction

            List<double> TotalGripArray = new List<double>();
            string FileName = "TotalGrip.xlsx";

            List<double> TotalLapsarray = new List<double>();
            string FileName2 = "TimeLaps.xlsx";

            if (FileName == null)
            {
                FileName = "TotalGrip.xlsx";
            }

            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(FileName, false))
            {
                WorkbookPart workbookPart = doc.WorkbookPart;
                Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().First();
                WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();


                if (!sheetData.Elements<Row>().Any())
                {
                    MessageBox.Show("Current File is Empty");
                }
                else
                {
                    foreach (Row row in sheetData.Elements<Row>())
                    {
                        foreach (Cell cell in row.Elements<Cell>())
                        {
                            if (cell.CellValue != null)
                            {
                                string cellvalue2 = cell.CellValue.InnerText;

                                if (cell.DataType == null || cell.DataType.Value == CellValues.Number)
                                {
                                    if (double.TryParse(cellvalue2, out doubleValue))
                                    {
                                        TotalGripArray.Add(doubleValue);
                                    }
                                }
                            }

                        }
                    }

                    string Lapsstr = TotLaps.Text;
                    Laps = int.Parse(Lapsstr);

                }

            }
            using (SpreadsheetDocument doc2 = SpreadsheetDocument.Open(FileName2, false))
            {
                WorkbookPart workbookPart2 = doc2.WorkbookPart;
                Sheet sheet2 = workbookPart2.Workbook.Descendants<Sheet>().First();
                WorksheetPart worksheetPart2 = (WorksheetPart)workbookPart2.GetPartById(sheet2.Id);
                SheetData sheetData = worksheetPart2.Worksheet.Elements<SheetData>().First();


                if (!sheetData.Elements<Row>().Any())
                {
                    MessageBox.Show("Current File is Empty");
                }
                else
                {
                    foreach (Row row2 in sheetData.Elements<Row>())
                    {
                        foreach (Cell cell2 in row2.Elements<Cell>())
                        {
                            if (cell2.CellValue != null)
                            {
                                string cellvalue2 = cell2.CellValue.InnerText;

                                if (cell2.DataType == null || cell2.DataType.Value == CellValues.Number)
                                {
                                    if (double.TryParse(cellvalue2, out doubleValue2))
                                    {
                                        TotalLapsarray.Add(doubleValue2);
                                    }
                                }
                            }

                        }
                    }


                }

            }


            //for(int i = 0; i != TotalLapsarray.Count;i++)
            //{
            //    MessageBox.Show(TotalLapsarray[i].ToString());
            //}


            //Console.WriteLine("Enter the amount of compounds to compare");
            //int CompInt = int.Parse(Console.ReadLine());
            //comparision

            int CompInt = 2;
            if (CompInt == 1)
            {

            }
            else if (CompInt == 2)
            {
                // take the laps in the full array into separate arrays
                string FirstLapStr = FirstLapLen.Text;
                FirstLaps = int.Parse(FirstLapStr);


                string SecondLapStr = SecondLapLen.Text;

                SecondLaps = int.Parse(SecondLapStr);

                for (int r = 0; r != FirstLaps; r++)
                {
                    double FirsttestVar = TotalGripArray[r];
                    FirstTest.Add(FirsttestVar);
                    // putting the total grip into the first arary
                }

                int k4 = 0;
                for (int e = FirstLaps; e != TotalGripArray.Count; e++)
                {

                    SecondTest.Add(TotalGripArray[e]);


                    k4 += 1;


                    // putting the total grip into the second arary

                }
                int countFT = FirstTest.Count();
                int countST = SecondTest.Count();



                //Console.WriteLine("First Test");
                //for (int m = 0; m != countFT; m++)
                //{
                //    Console.WriteLine(FirstTest[m]);
                //}
                //Console.WriteLine("Second Test");
                //for (int h = 0; h != countST; h++)
                //{
                //    Console.WriteLine(SecondTest[h]);
                //}
                List<int> FirstTotLapsint = Enumerable.Range(1, countFT).ToList();
                List<double> FirstTotLap = FirstTotLapsint.Select(n => (double)n).ToList();
                List<int> SecondTotLapsint = Enumerable.Range(1, countST).ToList();
                List<double> SecondTotLaps = SecondTotLapsint.Select(n => (double)n).ToList();

                int Count1 = FirstLaps;
                int Count2 = SecondLaps - 1;
                var coefficient = 0;

                if (FirstTest.Count() < 2 || SecondTest.Count() < 2)
                {
                    MessageBox.Show("not enough data to extrapolate the data");
                }
                else // extrapolating the data
                {
                    Grad1 = (FirstTest[Count1 - 1] - FirstTest[0]) / (Count1 - 1);
                    c1 = (FirstTest[0]) - (1 * Grad1);
                    Grad2 = (SecondTest[Count2 - 1] - SecondTest[0]) / (Count2 - 1);
                    c2 = (SecondTest[0]) - (1 * Grad2);




                }
                c1 = Math.Round(c1, 3);
                Grad1 = Math.Round(Grad1, 3);
                c2 = Math.Round(c2, 3);
                Grad2 = Math.Round(Grad2, 3);
                string LapLengthStr = LapLength.Text;
                RaceLaps = int.Parse(LapLengthStr);



                for (int p = 0; p != RaceLaps; p++)
                {
                    FinalFirstStrat.Add((Grad1 * (p + 1)) + c1);
                    SecondStratFinal.Add((Grad2 * (p + 1)) + c2);


                }

                //for (int g = 0; g != RaceLaps; g++) // adding into the strategy list
                //{
                //    if (FinalFirstStrat[g] > SecondStratFinal[g])
                //    {
                //        Strategy.Add("First Compound");
                //    }
                //    else if (SecondStratFinal[g] > FinalFirstStrat[g])
                //    {
                //        Strategy.Add("Second Compound");
                //    }
                //    else if (FinalFirstStrat == SecondStratFinal && Strategy.Count >= 1)
                //    {
                //        Strategy.Add(Strategy[g - 1]);
                //    }
                //    else if (FinalFirstStrat == SecondStratFinal && Strategy.Count < 1)
                //    {
                //        Strategy.Add("First Strat");
                //    }
                //}
                //for (int x = 0; x != RaceLaps; x++)
                //{
                //    MessageBox.Show(Strategy[x]);
                //}




            }
            // adding into the graphs
            List<double> RaceLapsArray = new List<double>();



            for (int i = 0; i != FinalFirstStrat.Count; i++)
            {
                //MessageBox.Show("Second Test");
                //MessageBox.Show((SecondTest[i]).ToString());
                RaceLapsArray.Add(i + 1);


            }
            int h = 0;
            int h2 = 0;
            for (int o = 0; o != FirstLaps; o++)
            {
                h += 1;
                FirstLapsLength.Add(h);


            }
            for (int o2 = 0; o2 != SecondLaps; o2++)
            {
                h2 += 1;
                SecondLapsLength.Add(h2);
            }
            var x1 = FirstLapsLength.ToArray();
            var y1 = FirstTest.ToArray();
            var x2 = SecondLapsLength.ToArray();
            var y2 = SecondTest.ToArray();

            // extrapolating for x^2



            for (int p = 0; p != RaceLaps; p++)
            { // maths for extrapolating the data for a number of total laps
                double lapNum = p + 1;



            }
            for (int k = 1; k != FinalStratQuadList.Count + 1; k++)
            {
                QuadRaceLapsFirst.Add(k);
            }
            for (int k2 = 1; k2 != SecondStratQuadList.Count + 1; k2++)
            {
                QuadRaceLapsSecond.Add(k2);
            }



            // More accurate y=mc+c - trendline
            int nL = 0;
            if (FirstLapsLength.Count > SecondLapsLength.Count)
            {
                nL = FirstLapsLength.Count;


            }
            else
            {
                nL = SecondLapsLength.Count;
            }

            for (int i = 0; i < nL; i++)
            {

                double x = FirstLapsLength[i];
                double y = FirstTest[i];

                sumX += x;
                sumY += y;
                sumXY += x * y;
                sumX2 += x * x;


            }
            for (int l = 0; l != SecondLapsLength.Count; l++)
            {
                double xS = SecondLapsLength[l];
                double yS = SecondTest[l];

                sumXS += xS;
                sumYS += yS;
                sumXYS += xS * yS;
                sumX2S += xS * xS;
            }
            double m = (nL * sumXY - sumX * sumY) / (nL * sumX2 - sumX * sumX);
            double c = (sumY - m * sumX) / nL;

            double mS = ((SecondLapsLength.Count * sumXYS) - (sumXS * sumYS)) / ((SecondLapsLength.Count * sumX2S) - (sumXS * sumXS));
            double cS = (sumYS - mS * sumXS) / SecondLapsLength.Count;
            List<double> trendlineY = new List<double>();
            List<double> trendlineX = new List<double>();
            List<double> trendlineYS = new List<double>();
            List<double> trendlineXS = new List<double>();

            for (int i = 0; i != RaceLapsArray.Count + 1; i++)
            {
                double x = i + 1; ;
                double y = m * x + c;

                double xs = i + 1;
                double ys = mS * xs + cS;

                trendlineX.Add(x);
                trendlineY.Add(y);

                trendlineXS.Add(xs);
                trendlineYS.Add(ys);


            }










            chart1.ChartAreas[0].AxisX.Title = "Lap Number";
            chart1.ChartAreas[0].AxisY.Title = "Total Grip";
            chart1.Titles.Clear();
            chart1.Titles.Add("Tyre Compound, Total Grip against Lap");
            Series FirstComp = chart1.Series.Add("FinalFirstStrat");
            FirstComp.ChartType = SeriesChartType.Line;
            FirstComp.Points.DataBindXY(RaceLapsArray, FinalFirstStrat);
            Series SecondComp = chart1.Series.Add("Second Comp");
            SecondComp.ChartType = SeriesChartType.Line;
            SecondComp.Points.DataBindXY(RaceLapsArray, SecondStratFinal);
            // second chart
            chart2.ChartAreas[0].AxisX.Title = "Lap Number";
            chart2.ChartAreas[0].AxisY.Title = "Total Grip";
            chart2.Titles.Clear();
            chart2.Titles.Add(" Tyre Compound against lap, full data");
            Series FirstStrat = chart2.Series.Add("First Strat");
            FirstStrat.ChartType = SeriesChartType.Line;
            FirstStrat.Points.DataBindXY(FirstLapsLength, FirstTest);
            Series SecondStrat = chart2.Series.Add("Second Strat");
            SecondStrat.ChartType = SeriesChartType.Line;
            SecondStrat.Points.DataBindXY(SecondLapsLength, SecondTest);

            //chart 3 quadtratic 
            chart3.ChartAreas[0].AxisX.Title = "Lap Number";
            chart3.ChartAreas[0].AxisY.Title = "Total Grip";
            chart3.Titles.Clear();
            chart3.Titles.Add("Tyre Compound against lap, Quadratic");
            Series FirstStratQuad = chart3.Series.Add("First Strat");
            FirstStratQuad.ChartType = SeriesChartType.Line;
            FirstStratQuad.Points.DataBindXY(trendlineX, trendlineY);
            Series SecondStratQuad = chart3.Series.Add("Second Strat");
            SecondStratQuad.ChartType = SeriesChartType.Line;
            SecondStratQuad.Points.DataBindXY(trendlineXS, trendlineYS);







        }

        private void chart1_Click(object sender, EventArgs e)
        {
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void LapLength_TextChanged(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
