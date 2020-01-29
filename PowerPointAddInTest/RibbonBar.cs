using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools.Ribbon;

namespace PowerPointAddInTest
{
    public partial class RibbonBar
    {
        private void RibbonBar_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void buttonStart_Click(object sender, RibbonControlEventArgs e)
        {
            var thisWindow = Globals.ThisAddIn.Application;
            var thisSlide = thisWindow.ActiveWindow.Presentation.Slides[1];
            Chart thisChart = GetChart(thisSlide);
            var tempSlide = Globals.ThisAddIn.Application.Presentations.Open(@"C:\test\template.pptx").Slides[1];
            Chart tempChart = GetChart(tempSlide);
            if (thisChart != null && tempChart != null)
            {
                try
                {
                    ChangeChartsStyle(thisChart, tempChart);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Exception", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("One of the slides haven't chart", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            tempSlide.Parent.Close();
            if (MessageBox.Show("Done", string.Empty, MessageBoxButtons.OK, MessageBoxIcon.Information) == DialogResult.OK)
            {
                GetScreenshoot();
            }
        }

        private void ChangeChartsStyle(Chart thisChart, Chart tempChart)
        {
            ChangeChartsColors(thisChart, tempChart);
            ChangeChartTitle(thisChart, tempChart);
        }

        private void ChangeChartTitle(Chart thisChart, Chart tempChart)
        {
            TextRange2 thisTextRange = thisChart.ChartTitle.Format.TextFrame2.TextRange;
            TextRange2 tempTextRange = tempChart.ChartTitle.Format.TextFrame2.TextRange;
            thisTextRange.Font.Fill.ForeColor.RGB = tempTextRange.Font.Fill.ForeColor.RGB;
            thisTextRange.Font.Italic = tempTextRange.Font.Italic;
            thisTextRange.Font.Size = tempTextRange.Font.Size;
            thisTextRange.Font.Spacing = tempTextRange.Font.Spacing;
            thisTextRange.Font.Bold = tempTextRange.Font.Bold;
        }

        private void ChangeChartsColors(Chart thisChart, Chart tempChart)
        {
            Series thisChartSeries = thisChart.SeriesCollection(1);
            Series tempChartSeries = tempChart.SeriesCollection(1);
            List<double> thisValues = GetValues(thisChartSeries);
            List<double> tempValues = GetValues(tempChartSeries);

            int thisPrevIdx = 0, tempPrevIdx = 0;
            for (int i = 1; i <= thisChartSeries.Points().Count; i++)
            {
                int thisIdx = GetNextMin(thisValues.ToArray(), thisPrevIdx);
                int tempIdx = GetNextMin(tempValues.ToArray(), tempPrevIdx);
                thisChartSeries.Points()[thisIdx].Format.Fill.ForeColor.RGB = tempChartSeries.Points()[tempIdx].Format.Fill.ForeColor.RGB;
                thisPrevIdx = GetNextMin(thisValues.ToArray(), thisPrevIdx);
                tempPrevIdx = GetNextMin(tempValues.ToArray(), tempPrevIdx);
            }
        }

        private void GetScreenshoot()
        {
            Bitmap bmp = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.CopyFromScreen(0, 0, 0, 0, Screen.PrimaryScreen.Bounds.Size);
                bmp.Save(@"C:\test\result.png");
            }
        }

        private List<double> GetValues(Series chartSeries)
        {
            List<double> tempValues = new List<double>();
            foreach (var item in chartSeries.Values as Array)
            {
                tempValues.Add((double)item);
            }

            return tempValues;
        }

        private int GetNextMin(double[] values, int prevIndex)
        {
            prevIndex--;
            int nextMinIdx = MaxValueIdx(values);
            for (int i = 0; i < values.Length; i++)
            {
                if (prevIndex != -1)
                {
                    if (values[i] > values[prevIndex] && values[i] < values[nextMinIdx])
                    {
                        nextMinIdx = i;
                    }
                }
                else
                {
                    if (values[i] < values[nextMinIdx])
                    {
                        nextMinIdx = i;
                    }
                }
            }
            return nextMinIdx + 1;
        }

        private int MaxValueIdx(double[] values)
        {
            int maxIdx = 0;
            for (int i = 0; i < values.Length; i++)
            {
                if (values[i] > values[maxIdx])
                {
                    maxIdx = i;
                }
            }
            return maxIdx;
        }

        private Chart GetChart(Slide thisSlide)
        {
            for (int i = 1; i <= thisSlide.Shapes.Count; i++)
            {
                if (thisSlide.Shapes[i].HasChart == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    return thisSlide.Shapes[i].Chart;
                }
            }

            return null;
        }
    }
}