using System.Reflection;
using ModbusExcel;
using System;
using System.Net.Sockets;
using NUnit.Framework;

namespace ModbusExcel.Tests
{
    
    
    /// <summary>
    ///This is a test class for RTDTest and is intended
    ///to contain all RTDTest Unit Tests
    ///</summary>
    [TestFixture]
    public class RTDTest
    {


        private TestContext testContextInstance;

        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }

        #region Additional test attributes
        // 
        //You can use the following additional attributes as you write your tests:
        //
        //Use ClassInitialize to run code before running the first test in the class
        //[ClassInitialize()]
        //public static void MyClassInitialize(TestContext testContext)
        //{
        //}
        //
        //Use ClassCleanup to run code after all tests in a class have run
        //[ClassCleanup()]
        //public static void MyClassCleanup()
        //{
        //}
        //
        //Use TestInitialize to run code before running each test
        //[TestInitialize()]
        //public void MyTestInitialize()
        //{
        //}
        //
        //Use TestCleanup to run code after each test has run
        //[TestCleanup()]
        //public void MyTestCleanup()
        //{
        //}
        //
        #endregion


        /// <summary>
        ///A test for RTD Constructor
        ///</summary>
        [Test]
        public void RTDConstructorTest()
        {
            RTD target = new RTD();
            target.ToString();
            Assert.AreEqual(target.ToString(),"ModbusExcel.RTD");
        }

        /// <summary>
        ///A test for ConnectData
        ///</summary>
        [Test]
        public void ConnectDataTest()
        {
            RTD target = new RTD(); // TODO: Initialize to an appropriate value
            int topicId = 0; // TODO: Initialize to an appropriate value

            Object[] topics = new Object[6];

            topics[0] = "127.0.0.1";
            topics[1] = "1";
            topics[2] = "N";
            topics[3] = "4849";
            topics[4] = "1";
            topics[5] = "2000";

            bool getNewValues = false; // TODO: Initialize to an appropriate value
            bool getNewValuesExpected = true; // TODO: Initialize to an appropriate value
            object actual;
            actual = target.ConnectData(topicId, ref topics, ref getNewValues);
            Assert.AreEqual(getNewValuesExpected, getNewValues);
        }

        /// <summary>
        ///A test for DisconnectData
        ///</summary>
        [Test]
        public void DisconnectDataTest()
        {
            RTD target = new RTD(); // TODO: Initialize to an appropriate value
            int topicId = 0; // TODO: Initialize to an appropriate value
            target.DisconnectData(topicId);
            Assert.Inconclusive("A method that does not return a value cannot be verified.");
        }

        /// <summary>
        ///A test for Heartbeat
        ///</summary>
        [Test]
        public void HeartbeatTest()
        {
            RTD target = new RTD(); // TODO: Initialize to an appropriate value
            int expected = 0; // TODO: Initialize to an appropriate value
            int actual;
            actual = target.Heartbeat();
            Assert.AreEqual(expected, actual);
            Assert.Inconclusive("Verify the correctness of this test method.");
        }

        /// <summary>
        ///A test for ModbusScan
        ///</summary>
        [Test]
        public void ModbusScanTest()
        {
            RTD target = new RTD(); // TODO: Initialize to an appropriate value
      // add event args back later      target.ModbusScan();
            Assert.Inconclusive("A method that does not return a value cannot be verified.");
        }

        /// <summary>
        ///A test for Now
        ///</summary>
        [Test]
        public void NowTest()
        {
            // If first vaule no more or less than 5 different from test value I guess we are good.
            double actual = RTD.Now();
            var comparer = new UnitTestHelpers.DoubleComparer(5);
            Assert.IsTrue(comparer.Compare(actual, RTD.Now()) == 0);
        }

        /// <summary>
        ///A test for RefreshData
        ///</summary>
        [Test]
        public void RefreshDataTest()
        {
            RTD target = new RTD(); // TODO: Initialize to an appropriate value
            int topicCount = 0; // TODO: Initialize to an appropriate value
            int topicCountExpected = 0; // TODO: Initialize to an appropriate value
            object[,] expected = null; // TODO: Initialize to an appropriate value
            object[,] actual;
            actual = target.RefreshData(ref topicCount);
            Assert.AreEqual(topicCountExpected, topicCount);
            Assert.AreEqual(expected, actual);
            Assert.Inconclusive("Verify the correctness of this test method.");
        }

        /// <summary>
        ///A test for ServerStart
        ///</summary>
        [Test]
        public void ServerStartTest()
        {

            RTD target = new RTD(); // TODO: Initialize to an appropriate value
            IRTDUpdateEvent callbackObject = null; // TODO: Initialize to an appropriate value
            int expected = 0; // TODO: Initialize to an appropriate value
            int actual;
            actual = target.ServerStart(callbackObject);
            Assert.AreEqual(expected, actual);
            Assert.Inconclusive("Verify the correctness of this test method.");
        }

        /// <summary>
        ///A test for ServerTerminate
        ///</summary>
        [Test]
        public void ServerTerminateTest()
        {
            RTD target = new RTD(); // TODO: Initialize to an appropriate value
            target.ServerTerminate();
            Assert.Inconclusive("A method that does not return a value cannot be verified.");
        }
    }
}
