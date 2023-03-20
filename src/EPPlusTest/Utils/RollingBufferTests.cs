using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace EPPlusTest.Utils
{
    [TestClass]
    public class RollingBufferTests
    {
        [TestMethod]    
        public void WriteSmallerBuffer()
        {
            var rb = new RollingBuffer(8);
            rb.Write(new byte[] { 1, 2, 3, 4 });
            var b = rb.GetBuffer();

            Assert.AreEqual(4, b.Length);
            Assert.AreEqual(1, b[0]);
            Assert.AreEqual(2, b[1]);
            Assert.AreEqual(3, b[2]);
            Assert.AreEqual(4, b[3]);
        }
        [TestMethod]
        public void WriteSmallerBufferTwoWrites()
        {
            var rb = new RollingBuffer(8);
            rb.Write(new byte[] { 1, 2, 3, 4 });
            rb.Write(new byte[] { 5, 6, 7 });
            var b = rb.GetBuffer();

            Assert.AreEqual(7, b.Length);
            Assert.AreEqual(1, b[0]);
            Assert.AreEqual(2, b[1]);
            Assert.AreEqual(3, b[2]);
            Assert.AreEqual(4, b[3]);
            Assert.AreEqual(5, b[4]);
            Assert.AreEqual(6, b[5]);
            Assert.AreEqual(7, b[6]);
        }
        [TestMethod]
        public void WriteSmallerBufferOverWriteFullBuffer()
        {
            var rb = new RollingBuffer(8);
            rb.Write(new byte[] { 1, 2, 3, 4 });
            rb.Write(new byte[] { 5, 6, 7, 8 });
            rb.Write(new byte[] { 9, 10 });
            var b = rb.GetBuffer();

            Assert.AreEqual(8, b.Length);
            Assert.AreEqual(3, b[0]);
            Assert.AreEqual(4, b[1]);
            Assert.AreEqual(5, b[2]);
            Assert.AreEqual(6, b[3]);
            Assert.AreEqual(7, b[4]);
            Assert.AreEqual(8, b[5]);
            Assert.AreEqual(9, b[6]);
            Assert.AreEqual(10, b[7]);
        }
        [TestMethod]
        public void WriteSmallerBufferOverWrite()
        {
            var rb = new RollingBuffer(8);
            rb.Write(new byte[] { 1, 2, 3, 4 });
            rb.Write(new byte[] { 5, 6, 7, 8, 9 });
            rb.Write(new byte[] { 10,11,12 });
            var b = rb.GetBuffer();

            Assert.AreEqual(8, b.Length);
            Assert.AreEqual(5, b[0]);
            Assert.AreEqual(6, b[1]);
            Assert.AreEqual(7, b[2]);
            Assert.AreEqual(8, b[3]);
            Assert.AreEqual(9, b[4]);
            Assert.AreEqual(10, b[5]);
            Assert.AreEqual(11, b[6]);
            Assert.AreEqual(12, b[7]);
        }
        [TestMethod]
        public void WriteSmallerBufferOverWriteWithBufferOfSameSize()
        {
            var rb = new RollingBuffer(8);
            rb.Write(new byte[] { 1, 2, 3, 4 });
            rb.Write(new byte[] { 5, 6, 7, 8, 9 });
            rb.Write(new byte[] { 10, 11, 12,13,14,15,16,17 });
            var b = rb.GetBuffer();

            Assert.AreEqual(8, b.Length);
            Assert.AreEqual(10, b[0]);
            Assert.AreEqual(11, b[1]);
            Assert.AreEqual(12, b[2]);
            Assert.AreEqual(13, b[3]);
            Assert.AreEqual(14, b[4]);
            Assert.AreEqual(15, b[5]);
            Assert.AreEqual(16, b[6]);
            Assert.AreEqual(17, b[7]);
        }
        [TestMethod]
        public void WriteSmallerBufferOverWriteWithBufferOfLargerSize()
        {
            var rb = new RollingBuffer(8);
            rb.Write(new byte[] { 1, 2, 3, 4 });
            rb.Write(new byte[] { 5, 6, 7, 8, 9 });
            rb.Write(new byte[] { 10, 11, 12, 13, 14, 15, 16, 17, 18 });
            var b = rb.GetBuffer();

            Assert.AreEqual(8, b.Length);
            Assert.AreEqual(11, b[0]);
            Assert.AreEqual(12, b[1]);
            Assert.AreEqual(13, b[2]);
            Assert.AreEqual(14, b[3]);
            Assert.AreEqual(15, b[4]);
            Assert.AreEqual(16, b[5]);
            Assert.AreEqual(17, b[6]);
            Assert.AreEqual(18, b[7]);
        }

    }
}
