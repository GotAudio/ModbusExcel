using ModbusExcel;
using System;
using NUnit.Framework;

namespace ModbusExcel.Tests
{
    
    
    /// <summary>
    ///This is a test class for IRtdServerTest and is intended
    ///to contain all IRtdServerTest Unit Tests
    ///</summary>
    [TestFixture]
    public class IRtdServerTest
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


        internal virtual IRtdServer CreateIRtdServer()
        {
            // TODO: Instantiate an appropriate concrete class.
            IRtdServer target = null;
            return target;
        }

        /// <summary>
        ///A test for ConnectData
        ///</summary>
        [Test]
        public void ConnectDataTest()
        {
            IRtdServer target = CreateIRtdServer(); // TODO: Initialize to an appropriate value
            int topicId = 0; // TODO: Initialize to an appropriate value
            object[] parameters = null; // TODO: Initialize to an appropriate value
            object[] parametersExpected = null; // TODO: Initialize to an appropriate value
            bool newValue = false; // TODO: Initialize to an appropriate value
            bool newValueExpected = false; // TODO: Initialize to an appropriate value
            object expected = null; // TODO: Initialize to an appropriate value
            object actual;
            actual = target.ConnectData(topicId: topicId, parameters: ref parameters, newValue: ref newValue);
            Assert.AreEqual(parametersExpected, parameters);
            Assert.AreEqual(newValueExpected, newValue);
            Assert.AreEqual(expected, actual);
            Assert.Inconclusive("Verify the correctness of this test method.");
        }

        /// <summary>
        ///A test for DisconnectData
        ///</summary>
        [Test]
        public void DisconnectDataTest()
        {
            IRtdServer target = CreateIRtdServer(); // TODO: Initialize to an appropriate value
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
            IRtdServer target = CreateIRtdServer(); // TODO: Initialize to an appropriate value
            int expected = 0; // TODO: Initialize to an appropriate value
            int actual;
            actual = target.Heartbeat();
            Assert.AreEqual(expected, actual);
            Assert.Inconclusive("Verify the correctness of this test method.");
        }

        /// <summary>
        ///A test for RefreshData
        ///</summary>
        [Test]
        public void RefreshDataTest()
        {
            IRtdServer target = CreateIRtdServer(); // TODO: Initialize to an appropriate value
            int topicCount = 0; // TODO: Initialize to an appropriate value
            int topicCountExpected = 0; // TODO: Initialize to an appropriate value
            object[,] expected = null; // TODO: Initialize to an appropriate value
            object[,] actual;
            actual = target.RefreshData(topicCount: ref topicCount);
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
            IRtdServer target = CreateIRtdServer(); // TODO: Initialize to an appropriate value
            IRTDUpdateEvent callback = null; // TODO: Initialize to an appropriate value
            int expected = 0; // TODO: Initialize to an appropriate value
            int actual;
            actual = target.ServerStart(callback);
            Assert.AreEqual(expected, actual);
            Assert.Inconclusive("Verify the correctness of this test method.");
        }

        /// <summary>
        ///A test for ServerTerminate
        ///</summary>
        [Test]
        public void ServerTerminateTest()
        {
            IRtdServer target = CreateIRtdServer(); // TODO: Initialize to an appropriate value
            target.ServerTerminate();
            Assert.Inconclusive("A method that does not return a value cannot be verified.");
        }
    }
}
