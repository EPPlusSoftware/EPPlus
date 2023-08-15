using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers
{
    internal static class BlockMatrix
    {
        public const int BLOCK_SIZE = 52;

        public static double[][] CreateBlockMatrix(int rows, int columns)
        {
            int blockRows = (rows + BLOCK_SIZE - 1) / BLOCK_SIZE;
            int blockColumns = (columns + BLOCK_SIZE - 1) / BLOCK_SIZE;

            double[][] blocks = new double[blockRows * blockColumns][];
            int blockIndex = 0;
            for (int iBlock = 0; iBlock < blockRows; ++iBlock)
            {
                int pStart = iBlock * BLOCK_SIZE;
                int pEnd = FastMathHelper.Min(pStart + BLOCK_SIZE, rows);
                int iHeight = pEnd - pStart;
                for (int jBlock = 0; jBlock < blockColumns; ++jBlock)
                {
                    int qStart = jBlock * BLOCK_SIZE;
                    int qEnd = FastMathHelper.Min(qStart + BLOCK_SIZE, columns);
                    int jWidth = qEnd - qStart;
                    blocks[blockIndex] = new double[iHeight * jWidth];
                    ++blockIndex;
                }
            }

            return blocks;
        }
    }
}