using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace Nissi
{
    public partial class Form1 : Form
    {
        private double[] dataX;
        private double[] dataY;
        private string series_name_data_default = "Данные";
        private string series_name_data;

        private double[] dataXe;

        private double[] data_line;
        private double coef_line_a, coef_line_b;
        private string series_name_line = "Линейная";

        private double[] data_pow;
        private double coef_pow_a, coef_pow_b;
        private string series_name_pow = "Степенная";

        private double[] data_exp;
        private double coef_exp_a, coef_exp_b;
        private string series_name_exp = "Экспоненциальная";

        private double stepX;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int RangeCb, RangeCe, RangeRb, RangeRe;//Диапазоны оси Абцисс
            int RangeDCb, RangeDCe, RangeDRb, RangeDRe;//Диапазоны оси Ординат
            try
            {
                RangeCb = Convert.ToInt32(textBoxRangeCb.Text);
                RangeCe = Convert.ToInt32(textBoxRangeCe.Text);
                RangeRb = Convert.ToInt32(textBoxRangeRb.Text);
                RangeRe = Convert.ToInt32(textBoxRangeRe.Text);
                RangeDCb = Convert.ToInt32(textBoxRangeDCb.Text);
                RangeDCe = Convert.ToInt32(textBoxRangeDCe.Text);
                RangeDRb = Convert.ToInt32(textBoxRangeDRb.Text);
                RangeDRe = Convert.ToInt32(textBoxRangeDRe.Text);
            }
            catch
            {
                MessageBox.Show("Проблема с текстом в полях диапазона, должны быть числа", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            //Проверяем диапазн, истина == представление данных строкой иначе столбцом
            bool RangeByRows = RangeCb != RangeCe;
            bool RangeDByRows = RangeDCb != RangeDCe; //Оси Ординат
            int CountX, CountY;

            if (RangeByRows && RangeRb != RangeRe)
            {
                MessageBox.Show("Проблема с диапозоном ячеек Оси Абцисс!", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            //Вычисление числа элементов X
            if (RangeByRows)
                CountX = RangeCe - RangeCb + 1;
            else
                CountX = RangeRe - RangeRb + 1;
            
            //Вычисление числа элементов Y
            if (RangeDByRows)
                CountY = RangeDCe - RangeDCb + 1;
            else
                CountY = RangeDRe - RangeDRb + 1;

            if(CountX != CountY)
            {
                MessageBox.Show("Число элементов Оси Абцисс и оси Ординат не сошлось", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            dataX = new double[CountX];
            dataY = new double[CountY];

            OpenFileDialog openDialog = new OpenFileDialog();
            openDialog.Filter = "Файл Excel|*.XLSX;*.XLS";
            var result = openDialog.ShowDialog();
            if (result != DialogResult.OK)
            {
                MessageBox.Show("Файл не выбран!", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            //string fileName = System.IO.Path.GetFileName(openDialog.FileName);

            Microsoft.Office.Interop.Excel.Application ExcelApp;
            Microsoft.Office.Interop.Excel.Workbook ExcelWorkbook;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorksheet;

            try
            {
                ExcelApp = new Microsoft.Office.Interop.Excel.Application();
                ExcelWorkbook = ExcelApp.Workbooks.Open(openDialog.FileName);
                ExcelWorksheet = ExcelWorkbook.Sheets[ numericUpDownDataListNum.Value ];
            }
            catch(Exception exe)
            {
                /*
                 * Скорее всего произошла утечка ресурсов((
                 */
                MessageBox.Show(exe.ToString(), "Возникла ошибка при открытии файла Excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            /*Microsoft.Office.Interop.Excel.Range r =;
            int x = r.Column;*/
            try
            {
                //Загрузка оси Х
                if (RangeByRows)//если данные хранятся в файле по строкам
                {
                    int p = 0;
                    for (int i = RangeCb - 1; i < RangeCe; i++)
                    {
                        int j = RangeRb;
                        dataX[p++] = Convert.ToDouble(ExcelWorksheet.Cells[j, i + 1].Text.ToString());
                    }
                }
                else//если данные хранятся в файле по столбцам
                {
                    int p = 0;
                    for (int j = RangeRb - 1; j < RangeRe; j++)
                    {
                        int i = RangeCb;
                        dataX[p++] = Convert.ToDouble(ExcelWorksheet.Cells[j + 1, i].Text.ToString());
                    }
                }

                //Загрузка Y
                if (RangeDByRows)//если данные хранятся в файле в виде строки
                {
                    int p = 0;
                    for (int i = RangeDCb - 1; i < RangeDCe; i++)
                    {
                        int j = RangeDRb;
                        dataY[p++] = Convert.ToDouble(ExcelWorksheet.Cells[j, i + 1].Text.ToString());
                    }
                    if (checkBoxCaptionY.Checked && RangeDCb > 1)//Загрузка подписи данных
                        series_name_data = ExcelWorksheet.Cells[RangeDRb, RangeDCb - 1].Text.ToString();
                    else
                        series_name_data = series_name_data_default;
                }
                else
                {
                    int p = 0;
                    for (int j = RangeDRb - 1; j < RangeDRe; j++)
                    {
                        int i = RangeDCb;
                        dataY[p++] = Convert.ToDouble(ExcelWorksheet.Cells[j + 1, i].Text.ToString());
                    }
                    if (checkBoxCaptionY.Checked && RangeDRb > 1)//Загрузка подписи данных
                        series_name_data = ExcelWorksheet.Cells[RangeDRb - 1, RangeDCb].Text.ToString();
                    else
                        series_name_data = series_name_data_default;
                }
            }
            catch
            {
                MessageBox.Show("Ошибка в файле Excel либо неверный диаппазон");
            }
            ExcelWorkbook.Close(false, Type.Missing, Type.Missing); // закрыть файл не сохраняя
            ExcelApp.Quit(); // Закрыть экземпляр Excel
            GC.Collect();   //Инициировать сборщик мусора

            //Вычисление среднего шага оси Х
            stepX = (dataX.Last() - dataX.First()) / (dataX.Length - 1);
            
            //Вычисления, построение графиков
            calculate();
            set_chart_series();
            checkBoxSeries_CheckedChanged(null, null);// Проверка выводимых линий трендов

            update_labels_coef();

            update_predict_table();
            update_acorrel_table();
        }

        private void tabPage2_Enter(object sender, EventArgs e)
        {
            /*
            calculate();
            set_chart_series();
            */
        }

        private void numericUpDownPredictNum_ValueChanged(object sender, EventArgs e)
        {
            //Вычисления и построение графиков
            calculate();
            set_chart_series();
        }

        private void update_labels_coef()
        {
            labelCoefLine.Text  = "y = " + coef_line_a.ToString("N") + " + " + coef_line_b.ToString("N") + " * x";
            labelCoefPow.Text = "y = " + coef_pow_a.ToString("E4") + " * x ^ " + coef_pow_b.ToString("N");
            labelCoefExp.Text = "y = e ^ (" + coef_exp_a.ToString("N") + " + " + coef_exp_b.ToString("N") + " * x)";

            labelRxyLine.Text = (coef_line_b * Math.Sqrt(sigma2(dataX) ) / Math.Sqrt(sigma2(dataY))).ToString("N");
            labelAperLine.Text = (approx_error(dataY, data_line)).ToString("N") + " %";
            labelR2Line.Text = (1.0 - sigma2(data_line) / sigma2(dataY)).ToString("N");

            labelRxyPow.Text = ( Math.Sqrt(sigma2(dataX)) / Math.Sqrt(sigma2(dataY))).ToString("N");
            labelAperPow.Text = (approx_error(dataY, data_pow)).ToString("N") + " %";
            labelR2Pow.Text = (1.0 - sigma2(data_pow) / sigma2(dataY)).ToString("N");

            labelRxyExp.Text = (Math.Exp(coef_exp_b) * Math.Sqrt(sigma2(dataX)) / Math.Sqrt(sigma2(dataY))).ToString("N");
            labelAperExp.Text = (approx_error(dataY, data_exp)).ToString("N") + " %";
            labelR2Exp.Text = (1.0 - sigma2(data_exp) / sigma2(dataY)).ToString("N");

            labelCorrel.Text = correlation(dataX, dataY).ToString("N");
        }

        private void update_predict_table()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("сдвиг");
            dt.Columns.Add("Год");
            dt.Columns.Add("Линейная функция");
            dt.Columns.Add("Степенная функция");
            dt.Columns.Add("Экспоненциальная функция");

            int offset = -2;
            double year = dataX[dataX.Length + offset];
            for(int i = 0; i < 20; i++)
            {
                double new_year = year + stepX * i;
                DataRow r = dt.NewRow();
                r["сдвиг"] = offset + i;
                r["Год"] = new_year;
                r["Линейная функция"] = f_line(new_year).ToString("N");
                r["Степенная функция"] = f_pow(new_year).ToString("N");
                r["Экспоненциальная функция"] = f_exp(new_year).ToString("N");
                dt.Rows.Add(r);
            }
            dataGridViewForecast.DataSource = dt;
        }

        private void update_acorrel_table()
        {
            string name_lag = "Лаг";
            string name_acor = "Коэффициент автокорреляции";
            int n = dataY.Length / 2;
            double[] aY = new double[dataY.Length];
            
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add(name_lag);
            dt.Columns.Add(name_acor);
            for(int i = 0; i < n; i++ )
            {
                for (int j = 0; j < aY.Length; j++)
                {
                    aY[j] = dataY[ (i + j) % dataY.Length];
                }
                DataRow r = dt.NewRow();
                r[name_lag] = i;
                r[name_acor] = correlation(dataY, aY);
                dt.Rows.Add(r);
            }
            dataGridViewAutocorrelation.DataSource = dt;
        }


        private void checkBoxSeries_CheckedChanged(object sender, EventArgs e)
        {
            chart1.Series[series_name_line].Enabled = checkBoxSeriesLine.Checked;
            chart1.Series[series_name_pow].Enabled = checkBoxSeriesPow.Checked;
            chart1.Series[series_name_exp].Enabled = checkBoxSeriesExp.Checked;
        }

        private void set_chart_series()
        {
            chart1.Series.Clear();
            chart1.ChartAreas[0].AxisX.Minimum = dataX.First();
            chart1.ChartAreas[0].AxisX.Maximum = dataX.Last() + stepX * (double)numericUpDownPredictNum.Value;
            chart1.ChartAreas[0].AxisX.MajorGrid.Interval = stepX;

            chart1.Series.Add(series_name_data);
            chart1.Series[series_name_data].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Point;
            chart1.Series[series_name_data].ChartArea = "ChartArea1";
            chart1.Series[series_name_data].Points.DataBindXY(dataX, dataY);
            

            chart1.Series.Add(series_name_line);
            chart1.Series[series_name_line].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
            chart1.Series[series_name_line].ChartArea = "ChartArea1";
            chart1.Series[series_name_line].Points.DataBindXY(dataXe, data_line);

            chart1.Series.Add(series_name_pow);
            chart1.Series[series_name_pow].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
            chart1.Series[series_name_pow].BorderDashStyle = System.Windows.Forms.DataVisualization.Charting.ChartDashStyle.Dash;
            chart1.Series[series_name_pow].ChartArea = "ChartArea1";
            chart1.Series[series_name_pow].Points.DataBindXY(dataXe, data_pow);
           
            

            chart1.Series.Add(series_name_exp);
            chart1.Series[series_name_exp].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
            chart1.Series[series_name_exp].ChartArea = "ChartArea1";
            chart1.Series[series_name_exp].Points.DataBindXY(dataXe, data_exp);
        }

        private void make_dataXe(int predict_num)
        {
            dataXe = new double[dataX.Length + predict_num];
            for (int i = 0; i < dataXe.Length; i++)
            {
                if (i < dataX.Length)
                    dataXe[i] = dataX[i];
                else
                    dataXe[i] = dataXe[i - 1] + stepX;
            }
        }

        private void calculate()
        {
            make_dataXe((int)numericUpDownPredictNum.Value);

            calc_line();
            calc_powf();
            calc_exp();
        }

        private double mean(double[] x)
        {
            double res = 0;
            int n = x.Length;
            for (int i = 0; i < n; i++)
            {
                res += x[i];
            }
            return res / n;
        }
        
        private double sigma2(double[] x)
        {
            double res = 0;
            double meanx = mean(x);
            int n = x.Length;
            for(int i = 0; i < n; i++)
            {
                double tmp = (x[i] - meanx);
                res += tmp * tmp;
            }
            return res / n;
        }

        private double mean_product(double[] x, double[] y)
        {
            double res = 0;
            int n = x.Length;
            for(int i = 0; i < n; i++)
            {
                res += x[i] * y[i];
            }
            return res / n;
        }

        private double f_line(double x)
        {
            return coef_line_a + coef_line_b * x;
        }
        private void calc_line()
        {
            coef_line_b = (mean_product(dataX, dataY) - mean(dataX) * mean(dataY)) / sigma2(dataX);
            coef_line_a = mean(dataY) - coef_line_b * mean(dataX);
            data_line = new double[dataXe.Length];
            for(int i = 0; i < data_line.Length; i++)
            {
                data_line[i] = f_line( dataXe[i] );
            }
        }

        private double f_pow(double x)
        {
            return coef_pow_a * Math.Pow(x, coef_pow_b);
        }
        private void calc_powf()
        {
            int n = dataX.Length;
            double[] logX = new double[n];
            double[] logY = new double[n];
            for(int i = 0; i < n; i++)
            {
                logX[i] = Math.Log(dataX[i]);
                logY[i] = Math.Log(dataY[i]);
            }
            double b_log = (mean_product(logX, logY) - mean(logX) * mean(logY)) / sigma2(logX);
            double a_log = mean(logY) - b_log * mean(logX);
            coef_pow_a = Math.Exp(a_log);
            coef_pow_b = b_log;
            data_pow = new double[dataXe.Length];
            for(int i = 0; i < data_pow.Length; i++)
            {
                data_pow[i] = f_pow(dataXe[i]);
            }
        }

        private double f_exp(double x)
        {
            return Math.Exp(coef_exp_a + coef_exp_b * x);
        }
        private void calc_exp()
        {
            int n = dataX.Length;
            double[] logY = new double[n];
            for (int i = 0; i < n; i++)
            {
                logY[i] = Math.Log(dataY[i]);
            }
            double b_log = (mean_product(dataX, logY) - mean(dataX) * mean(logY)) / sigma2(dataX);
            coef_exp_a = mean(logY) - b_log * mean(dataX);
            coef_exp_b = b_log;
            data_exp = new double[dataXe.Length];
            for(int i = 0; i < data_exp.Length; i++)
            {
                data_exp[i] = f_exp(dataXe[i]);
            }
        }

        private double approx_error(double[] Y, double[] f)
        {
            double res = 0;
            int n = Y.Length;
            for(int i = 0; i < n; i++)
            {
                res += Math.Abs( (Y[i] - f[i]) / Y[i]);
            }
            return res / n * 100;
        }

        private double correlation(double[] X, double[] Y)
        {
            double res = 0;
            int n = X.Length;
            for(int i = 0; i < n; i++)
            {
                res += (X[i] - mean(X)) * (Y[i] - mean(Y));
            }
            return res / (Math.Sqrt(sigma2(X)) * Math.Sqrt(sigma2(Y))* n) ;
        }

    }
}
