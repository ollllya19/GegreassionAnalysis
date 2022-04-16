g System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace CorrelationAnalysis
{
    public class CorrelationAnalysis
    {
        protected int column;
        protected int row;
        protected double[,] matr;
        protected double[] averValues;
        protected double[] dispEstMatr;
        protected double[,] standartMatr;
        protected double[,] covarMatr;
        protected double[,] corrMatr;
        protected int[,] hyptmatr;
        protected const double tTable = 1.994954;

        public CorrelationAnalysis(int row, int column)
        {
            this.row = row;
            this.column = column;
            matr = new double[row, column];
            averValues = new double[column];
            dispEstMatr = new double[column];
            standartMatr = new double[row, column];
            covarMatr = new double[column, column];
            corrMatr = new double[column, column];
            hyptmatr = new int[column, column];
        }

        public double[,] Matr { get => matr; }

        public double[,] StandartMatr { get => standartMatr; }

        public double[,] CovarMatr { get => covarMatr; }

        public double[,] CorrMatr { get => corrMatr; }

        public double[] DispMatr { get => dispEstMatr; }

        public int[,] HyptMatr { get => hyptmatr; }

        public void Execute()
        {
            ExelWork exelObj = new ExelWork(row, column);
            matr = exelObj.ExportFile();

            FindAverage(matr);
            FindVarianceEstimate(matr);
            StandartizedMatrix(matr);
            CorrelationMatrix();
            CovariationMatrix(matr);
            HypoteticMatr();
        }

        protected void FindAverage(double[,] matr)
        {
            for (int i = 0; i < column; i++)
            {
                double sum = 0;
                for (int j = 0; j < row; j++)
                {
                    sum += matr[j, i];
                }
                averValues[i] = sum / row;
            }
        }

        protected void FindVarianceEstimate(double[,] matr)
        {
            for (int i = 0; i < column; i++)
            {
                double sum = 0;
                for (int j = 0; j < row; j++)
                {
                    sum += Math.Pow(matr[j, i] - averValues[i], 2);
                }
                dispEstMatr[i] = sum / row;
            }
        }

        protected void StandartizedMatrix(double[,] matr)
        {
            for (int i = 0; i < row; i++)
            {
                for (int j = 0; j < column; j++)
                {
                    standartMatr[i, j] = (matr[i, j] - averValues[j]) / Math.Sqrt(dispEstMatr[j]);
                }
            }
        }

        protected void CovariationMatrix(double[,] matr)
        {
            for (int i = 0; i < column; i++)
            {
                for (int j = 0; j < column; j++)
                {
                    double sum = 0;
                    for (int k = 0; k < row; k++)
                    {
                        sum += (matr[k, i] - averValues[i]) * (matr[k, j] - averValues[j]);
                    }
                    covarMatr[i, j] = sum / row;
                }
            }
        }

        protected void CorrelationMatrix()
        {
            for (int i = 0; i < column; i++)
            {
                for (int j = 0; j < column; j++)
                {
                    double sum = 0;
                    for (int k = 0; k < row; k++)
                    {
                        sum += standartMatr[k, i] * standartMatr[k, j];
                    }
                    corrMatr[i, j] = sum / row;
                }
            }
        }

        protected void HypoteticMatr()
        {
            for (int i = 0; i < column; i++)
            {
                for (int j = 0; j < column; j++)
                {
                    if (Math.Abs(corrMatr[i, j] * Math.Sqrt(row - 2)
                        / Math.Sqrt(1 - Math.Pow(corrMatr[i, j], 2))) < tTable)
                        hyptmatr[i, j] = 0;
                    else
                        hyptmatr[i, j] = 1;
                }
            }

        }
    }


    //Matrix X is of 69x9 elements
    //Y is vector of 69 elements

    class RegressionAnalysis
    {
        protected CorrelationAnalysis corrAn;
        protected int n, m;
        protected double[,] matrX;
        protected double[,] matrTx;
        protected double[] y;


        public RegressionAnalysis(CorrelationAnalysis corrAn, int n, int m)
        {
            this.corrAn = corrAn;
            this.n = n;
            this.m = m;
            matrX = new double[n, m];
            matrTx = new double[m, n];

            for (int i = 0; i < n; i++)
            {
                for (int j = 0, k = 0; j < m; j++, k++)
                {
                    if (j == m - 1)
                        matrX[i, j] = 1;
                    else
                    {
                        if (j == 3 || j == 8)
                            k++;
                        matrX[i, j] = corrAn.Matr[i, k];
                    }

                }
            }

            y = new double[n];
            for (int i = 0; i < n; i++)
                y[i] = corrAn.Matr[i, 8];
        }

        public double[,] MatrX { get => matrX; }

        public double[,] MatrTX { get => matrTx; }

    }

    //class working with matrixes
    class Matrix
    {
        public static double[,] TransposeMatr(double[,] matr, int n, int m)
        {
            double[,] rez = new double[n, m];
            for (int i = 0; i < n; i++)
            {
                for (int j = 0; j < m; j++)
                {
                    rez[j, i] = matr[i, j];
                }
            }
            return rez;
        }

        public static double[,] MultipleMatr(double[,] matr1, double[,] matr2, int n1, int m1, int n2, int m2)
        {
            double[,] rez = new double[n1, m2];
            if (m1 == n2)
            {
                double sum = 0;

                for (int i = 0; i < n1; i++)
                {
                    for (int j = 0; j < m2; j++)
                    {
                        for (int k = 0; k < m1; k++)
                            sum += matr1[i, k] * matr2[k, j];
                        rez[i, j] = sum;
                        sum = 0;
                    }
                }
            }
            return rez;
        }


        public static double Determinant(double[,] matr, int n, int m)
        {
            return matr[0, 0] * matr[1, 1] * matr[2, 2] + matr[0, 1] * matr[1, 2] * matr[2, 0] + matr[1, 0] * matr[2, 1] * matr[0, 2] -
                matr[0, 2] * matr[1, 1] * matr[2, 0] - matr[1, 0] * matr[0, 1] * matr[2, 2] - matr[0, 0] * matr[2, 1] * matr[1, 2];
        }
    }

    //class of exporting data from exel
    class ExelWork
    {
        int column;
        int row;
        double[,] matr;

        public ExelWork(int row, int column)
        {
            this.row = row;
            this.column = column;
            matr = new double[row, column];
        }

        public double[,] ExportFile()
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Title = "Выбор документа";
            fileDialog.DefaultExt = "*.xls;*.xlsx";

            if (!(fileDialog.ShowDialog() == DialogResult.OK))
                return matr;

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var package = new ExcelPackage(fileDialog.FileName);
            ExcelWorksheet sheet = package.Workbook.Worksheets[0];

            for (int i = 0; i < row; i++)
            {
                for (int j = 0; j < column; j++)
                {
                    matr[i, j] = (double)sheet.Cells[i + 1, j + 1].Value;
                }
            }

            return matr;
        }
    }
}
