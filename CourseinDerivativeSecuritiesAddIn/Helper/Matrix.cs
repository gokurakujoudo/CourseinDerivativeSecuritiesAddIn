using System;

namespace DerivativeSecuritiesAddIn.Helper {
    public class Matrix {
        protected bool Equals(Matrix other) => Equals(_data, other._data);

        public override bool Equals(object obj) {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            return obj.GetType() == GetType() && Equals((Matrix) obj);
        }

        public override int GetHashCode() => (_data != null ? _data.GetHashCode() : 0);

        private readonly double[,] _data;

        public Matrix(int size) { _data = new double[size, size]; }

        public Matrix(int rows, int cols) { _data = new double[rows, cols]; }

        public Matrix(double[,] data) { _data = data; }

        public static Matrix operator +(Matrix matrix1, Matrix matrix2) {
            var matrix1Rows = matrix1.Data.GetLength(0);
            var matrix1Columns = matrix1.Data.GetLength(1);

            var matrix2Rows = matrix2.Data.GetLength(0);
            var matrix2Columns = matrix2.Data.GetLength(1);

            if (matrix1Rows != matrix2Rows || matrix1Columns != matrix2Columns) {
                throw new Exception("Matrix Dimensions Don't Agree!");
            }

            var result = new double[matrix1Rows, matrix1Columns];
            for (var i = 0; i < matrix1Rows; i++) {
                for (var j = 0; j < matrix1Columns; j++) {
                    result[i, j] = matrix1.Data[i, j] + matrix2.Data[i, j];
                }
            }

            return new Matrix(result);
        }

        public static Matrix operator -(Matrix matrix1, Matrix matrix2) {
            var matrix1Rows = matrix1.Data.GetLength(0);
            var matrix1Columns = matrix1.Data.GetLength(1);

            var matrix2Rows = matrix2.Data.GetLength(0);
            var matrix2Columns = matrix2.Data.GetLength(1);

            if (matrix1Rows != matrix2Rows || matrix1Columns != matrix2Columns) {
                throw new Exception("Matrix Dimensions Don't Agree!");
            }

            var result = new double[matrix1Rows, matrix1Columns];
            for (var i = 0; i < matrix1Rows; i++) {
                for (var j = 0; j < matrix1Columns; j++) {
                    result[i, j] = matrix1.Data[i, j] - matrix2.Data[i, j];
                }
            }

            return new Matrix(result);
        }

        public static Matrix operator *(Matrix matrix1, Matrix matrix2) {
            var matrix1Rows = matrix1.Data.GetLength(0);
            var matrix1Columns = matrix1.Data.GetLength(1);

            var matrix2Rows = matrix2.Data.GetLength(0);
            var matrix2Columns = matrix2.Data.GetLength(1);

            if (matrix1Columns != matrix2Rows) {
                throw new Exception("Matrix Dimensions Don't Agree!");
            }

            var result = new double[matrix1Rows, matrix2Columns];
            for (var i = 0; i < matrix1Rows; i++) {
                for (var j = 0; j < matrix2Columns; j++) {
                    for (var k = 0; k < matrix2Rows; k++) {
                        result[i, j] += matrix1.Data[i, k] * matrix2.Data[k, j];
                    }
                }
            }

            return new Matrix(result);
        }

        public static Matrix operator /(double i, Matrix matrix) { return new Matrix(ScaleBy(i, Inv(matrix.Data))); }

        public static bool operator ==(Matrix matrix1, Matrix matrix2) {
            var result = true;

            if (matrix1 == null || matrix2 == null) return false;

            var matrix1Rows = matrix1.Data.GetLength(0);
            var matrix1Columns = matrix1.Data.GetLength(1);

            var matrix2Rows = matrix2.Data.GetLength(0);
            var matrix2Columns = matrix2.Data.GetLength(1);

            if (matrix1Rows != matrix2Rows || matrix1Columns != matrix2Columns) {
                result = false;
            }
            else {
                for (var i = 0; i < matrix1Rows; i++) {
                    for (var j = 0; j < matrix1Columns; j++) {
                        if (matrix1.Data[i, j] != matrix2.Data[i, j]) result = false;
                    }
                }
            }
            return result;
        }

        public static bool operator !=(Matrix matrix1, Matrix matrix2) => !(matrix1 == matrix2);

        public void Display() {
            var rows = this.Data.GetLength(0);
            var columns = this.Data.GetLength(1);
            for (var i = 0; i < rows; i++) {
                for (var j = 0; j < columns; j++) {
                    Console.Write(this.Data[i, j].ToString("N2") + "    ");
                }
                Console.WriteLine();
            }
            Console.WriteLine();
        }

        public void Display(string format) {
            var rows = this.Data.GetLength(0);
            var columns = this.Data.GetLength(1);
            for (var i = 0; i < rows; i++) {
                for (var j = 0; j < columns; j++) {
                    Console.Write(this.Data[i, j].ToString(format) + "   ");
                }
                Console.WriteLine();
            }
            Console.WriteLine();
        }

        public Matrix Inverse() {
            if (this.IsSquare && !this.IsSingular) {
                return new Matrix(Inv(this.Data));
            }
            throw new Exception(@"Cannot find inverse for non square /singular matrix");
        }

        public Matrix Transpose() {
            var d = Transpose(this.Data);
            return new Matrix(d);
        }

        public static Matrix Zeros(int size) {
            var d = new double[size, size];
            return new Matrix(d);
        }

        public static Matrix Zeros(int rows, int cols) {
            var d = new double[rows, cols];
            return new Matrix(d);
        }

        public static Matrix LinearSolve(Matrix cof, Matrix con) { return cof.Inverse() * con; }

        public double Det() {
            if (this.IsSquare) {
                return Determinant(this.Data);
            }
            throw new Exception("Cannot Determine the DET for a non square matrix");
        }

        public bool IsSquare => this.Data.GetLength(0) == this.Data.GetLength(1);

        public bool IsSingular => (int) Det() == 0;

        public int Rows => this.Data.GetLength(0);

        public int Columns => this.Data.GetLength(1);

        private static double[,] Inv(double[,] srcMatrix) {
            var rows = srcMatrix.GetLength(0);
            var columns = srcMatrix.GetLength(1);
            try {
                if (rows != columns)
                    throw new Exception();
            }
            catch {
                Console.WriteLine("Cannot find inverse for an non-square matrix");
            }

            int q;
            var desMatrix = new double[rows, columns];
            var unitMatrix = UnitMatrix(rows);
            for (var p = 0; p < rows; p++) {
                for (q = 0; q < columns; q++) {
                    desMatrix[p, q] = srcMatrix[p, q];
                }
            }

            if (srcMatrix[0, 0] == 0) {
                var i = 1;
                while (i < rows) {
                    if (srcMatrix[i, 0] != 0) {
                        InterRow(srcMatrix, 0, i);
                        InterRow(unitMatrix, 0, i);
                        break;
                    }

                    i++;
                }
            }

            RowDiv(unitMatrix, 0, srcMatrix[0, 0]);
            RowDiv(srcMatrix, 0, srcMatrix[0, 0]);
            for (var p = 1; p < rows; p++) {
                q = 0;
                while (q < p) {
                    RowSub(unitMatrix, p, q, srcMatrix[p, q]);
                    RowSub(srcMatrix, p, q, srcMatrix[p, q]);
                    q++;
                }

                if (srcMatrix[p, p] != 0) {
                    RowDiv(unitMatrix, p, srcMatrix[p, p]);
                    RowDiv(srcMatrix, p, srcMatrix[p, p]);
                }

                if (srcMatrix[p, p] == 0) {
                    for (var j = p + 1; j < columns; j++) {
                        if (srcMatrix[p, j] != 0) {
                            for (var k = 0; k < rows; k++) {
                                for (q = 0; q < columns; q++) {
                                    srcMatrix[k, q] = desMatrix[k, q];
                                }
                            }

                            return Inverse(desMatrix);
                        }
                    }
                }
            }

            for (var p = rows - 1; p > 0; p--) {
                for (q = p - 1; q >= 0; q--) {
                    RowSub(unitMatrix, q, p, srcMatrix[q, p]);
                    RowSub(srcMatrix, q, p, srcMatrix[q, p]);
                }
            }

            for (var p = 0; p < rows; p++) {
                for (q = 0; q < columns; q++) {
                    srcMatrix[p, q] = desMatrix[p, q];
                }
            }

            return unitMatrix;
        }

        /// <summary>
        /// Inverse the Matrix
        /// </summary>
        /// <param name="srcMatrix">The Matrix to be Inversed</param>
        /// <returns>The Inversed Matrix</returns>
        private static double[,] Inverse(double[,] srcMatrix) {
            var rows = srcMatrix.GetLength(0);
            var columns = srcMatrix.GetLength(1);
            var desMatrix = new double[rows, columns];

            for (var i = 0; i < rows; i++) {
                for (var j = 0; j < columns; j++) {
                    desMatrix[i, j] = 0;
                }
            }

            try {
                if (rows != columns) throw new Exception();
            }
            catch {
                Console.WriteLine("Cannot Find Inverse for an non-square Matrix");
            }

            var determine = Determinant(srcMatrix);

            try {
                if (determine == 0) {
                    throw new Exception("Cannot Perform Inversion. Matrix Singular");
                }
            }
            catch (Exception ex) {
                Console.WriteLine(ex.Message);
            }

            for (var p = 0; p < rows; p++) {
                for (var q = 0; q < columns; q++) {
                    var tmp = FilterMatrix(srcMatrix, p, q);
                    var determineTmp = Determinant(tmp);
                    desMatrix[p, q] = Math.Pow(-1, p + q + 2) * determineTmp / determine;
                }
            }
            desMatrix = Transpose(desMatrix);
            return desMatrix;
        }

        /// <summary>
        /// Calculate the Determinant
        /// </summary>
        /// <param name="srcMatrix">The Matrix used to calculate</param>
        /// <returns>The Determinant</returns>
        private static double Determinant(double[,] srcMatrix) {
            int q;
            var rows = srcMatrix.GetLength(0);
            var columns = srcMatrix.GetLength(1);
            var desMatrix = new double[rows, columns];
            for (var p = 0; p < rows; p++) {
                for (q = 0; q < columns; q++) {
                    desMatrix[p, q] = srcMatrix[p, q];
                }
            }

            double det = 1;
            try {
                if (rows != columns) {
                    throw new Exception("Error: Matrix not Square");
                }
            }
            catch (Exception ex) {
                Console.WriteLine(ex.Message);
            }

            try {
                if (rows == 0) {
                    throw new Exception("Dimension of the Matrix 0X0");
                }
            }
            catch (Exception ex) {
                Console.WriteLine(ex.Message);
            }

            if (rows == 2) {
                return srcMatrix[0, 0] * srcMatrix[1, 1] - srcMatrix[0, 1] * srcMatrix[1, 0];
            }

            if (srcMatrix[0, 0] == 0) {
                var i = 1;
                while (i < rows) {
                    if (srcMatrix[i, 0] != 0) {
                        InterRow(srcMatrix, 0, i);
                        det *= -1;
                        break;
                    }
                    i++;
                }
            }

            if (srcMatrix[0, 0] == 0) return 0;

            det *= srcMatrix[0, 0];
            RowDiv(srcMatrix, 0, srcMatrix[0, 0]);
            for (var p = 1; p < rows; p++) {
                q = 0;
                while (q < p) {
                    RowSub(srcMatrix, p, q, srcMatrix[p, q]);
                    q++;
                }
                if (srcMatrix[p, p] != 0) {
                    det *= srcMatrix[p, p];
                    RowDiv(srcMatrix, p, srcMatrix[p, p]);
                }
                if (srcMatrix[p, p] == 0) {
                    for (var j = p + 1; j < columns; j++) {
                        if (srcMatrix[p, j] != 0) {
                            ColumnSub(srcMatrix, p, j, -1);
                            det *= srcMatrix[p, p];
                            RowDiv(srcMatrix, p, srcMatrix[p, p]);
                            break;
                        }
                    }
                }

                if (srcMatrix[p, p] == 0) return 0;
            }

            for (var p = 0; p < rows; p++) {
                for (q = 0; q < columns; q++) {
                    srcMatrix[p, q] = desMatrix[p, q];
                }
            }

            return det;
        }

        /// <summary>
        /// Scale the Matrix by a specified ratio
        /// </summary>
        /// <param name="scalar">The Ratio for the Scale</param>
        /// <param name="srcMatrix">The Matrix which will be scaled</param>
        /// <returns>A new Matrix which is scaled with a specified ratio</returns>
        private static double[,] ScaleBy(double scalar, double[,] srcMatrix) {
            var rows = srcMatrix.GetLength(0);
            var columns = srcMatrix.GetLength(1);
            var desMatrix = new double[rows, columns];
            for (var i = 0; i < rows; i++) {
                for (var j = 0; j < columns; j++) {
                    desMatrix[i, j] = scalar * srcMatrix[i, j];
                }
            }

            return desMatrix;
        }

        /// <summary>
        /// To get a unit matrix
        /// </summary>
        /// <param name="dimension">Dimension</param>
        /// <returns>The unit double array</returns>
        private static double[,] UnitMatrix(int dimension) {
            var a = new double[dimension, dimension];
            for (var i = 0; i < dimension; i++) {
                for (var j = 0; j < dimension; j++) {
                    if (i == j) a[i, j] = 1;
                    else a[i, j] = 0;
                }
            }

            return a;
        }

        /// <summary>
        /// Scale a specified Row
        /// </summary>
        /// <param name="srcMatrix">The Matrix to be scaled</param>
        /// <param name="row">The Specified Row</param>
        /// <param name="scaleRatio">The Scale Ratio</param>
        private static void RowDiv(double[,] srcMatrix, int row, double scaleRatio) {
            var columns = srcMatrix.GetLength(1);
            for (var i = 0; i < columns; i++) {
                srcMatrix[row, i] /= scaleRatio;
            }
        }

        /// <summary>
        /// Subtract a specified Row from another Row
        /// </summary>
        /// <param name="srcMatrix">The Matrix to be subtract</param>
        /// <param name="row1">The Row Index to be subtracted</param>
        /// <param name="row2">The Row Index to subtract</param>
        /// <param name="scaleRatio">Scale Ratio</param>
        private static void RowSub(double[,] srcMatrix, int row1, int row2, double scaleRatio) {
            var columns = srcMatrix.GetLength(1);
            for (var q = 0; q < columns; q++) {
                srcMatrix[row1, q] -= scaleRatio * srcMatrix[row2, q];
            }
        }

        /// <summary>
        /// Subtract a specified Column from another Column
        /// </summary>
        /// <param name="srcMatrix">The Matrix to be subtracted</param>
        /// <param name="column1">The Column Index to be subtracted</param>
        /// <param name="column2">The Column Index to Subtract</param>
        /// <param name="scaleRatio">Scale Ratio</param>
        private static void ColumnSub(double[,] srcMatrix, int column1, int column2, double scaleRatio) {
            var rows = srcMatrix.GetLength(0);
            for (var i = 0; i < rows; i++) {
                srcMatrix[i, column1] -= scaleRatio * srcMatrix[i, column2];
            }
        }

        /// <summary>
        /// Exchange a specified row's value
        /// </summary>
        /// <param name="srcMatrix">The Matrix to be exchanged</param>
        /// <param name="row1">Row index</param>
        /// <param name="row2">Row index</param>
        private static void InterRow(double[,] srcMatrix, int row1, int row2) {
            var columns = srcMatrix.GetLength(1);
            for (var k = 0; k < columns; k++) {
                var tmp = srcMatrix[row1, k];
                srcMatrix[row1, k] = srcMatrix[row2, k];
                srcMatrix[row2, k] = tmp;
            }
        }

        /// <summary>
        /// To Filter the Matrix
        /// </summary>
        /// <param name="srcMatrix">The Matrix to be Filtered</param>
        /// <param name="row">A specified Row</param>
        /// <param name="column">A specified Column</param>
        /// <returns>The Filtered Matrix</returns>
        private static double[,] FilterMatrix(double[,] srcMatrix, int row, int column) {
            var rows = srcMatrix.GetLength(0);
            var desMatrix = new double[rows - 1, rows - 1];
            var i = 0;
            for (var p = 0; p < rows; p++) {
                var j = 0;
                if (p != row) {
                    for (var q = 0; q < rows; q++) {
                        if (q != column) {
                            desMatrix[i, j] = srcMatrix[p, q];
                            j++;
                        }
                    }

                    i++;
                }
            }

            return desMatrix;
        }

        /// <summary>
        /// Transpose Matrix
        /// </summary>
        /// <param name="srcMatrix">The Matrix to be Transposed</param>
        /// <returns>The Transposed Matrix</returns>
        private static double[,] Transpose(double[,] srcMatrix) {
            var rows = srcMatrix.GetLength(0);
            var columns = srcMatrix.GetLength(1);
            var desMatrix = new double[rows, columns];
            for (var i = 0; i < rows; i++) {
                for (var j = 0; j < columns; j++) {
                    desMatrix[j, i] = srcMatrix[i, j];
                }
            }

            return desMatrix;
        }

        /// <summary>
        /// Gets or Sets the specified value of the Matrix
        /// </summary>
        /// <param name="i">The row position in the Matrix</param>
        /// <param name="j">The Column position in the Matrix</param>
        /// <returns>Double, The specified value of the Matrix</returns>
        public double this[int i, int j] {
            get => _data[i, j];
            set => _data[i, j] = value;
        }

        /// <summary>
        /// Gets the Matrix
        /// </summary>
        public double[,] Data => _data;
    }
}