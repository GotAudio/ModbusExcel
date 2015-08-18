using NUnit.Framework;

namespace ModbusExcel.Tests
{
    
    
    /// <summary>
    ///This is a test class for IRTDUpdateEventTest and is intended
    ///to contain all IRTDUpdateEventTest Unit Tests
    ///</summary>
    [TestFixture]
    public class IRTDUpdateEventTest
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


        internal virtual IRTDUpdateEvent CreateIRTDUpdateEvent()
        {
            // TODO: Instantiate an appropriate concrete class.
            IRTDUpdateEvent target = null;
            return target;
        }

        /// <summary>
        ///A test for Disconnect
        ///</summary>
        [Test]
        public void DisconnectTest()
        {
            IRTDUpdateEvent target = CreateIRTDUpdateEvent(); // TODO: Initialize to an appropriate value
            target.Disconnect();
            Assert.Inconclusive("A method that does not return a value cannot be verified.");
        }

        /// <summary>
        ///A test for UpdateNotify
        ///</summary>
        [Test]
        public void UpdateNotifyTest()
        {
            IRTDUpdateEvent target = CreateIRTDUpdateEvent(); // TODO: Initialize to an appropriate value
            target.UpdateNotify();
            Assert.Inconclusive("A method that does not return a value cannot be verified.");
        }

        /// <summary>
        ///A test for HeartbeatInterval
        ///</summary>
        [Test]
        public void HeartbeatIntervalTest()
        {
            IRTDUpdateEvent target = CreateIRTDUpdateEvent(); // TODO: Initialize to an appropriate value
            int expected = 0; // TODO: Initialize to an appropriate value
            int actual;
            target.HeartbeatInterval = expected;
            actual = target.HeartbeatInterval;
            Assert.AreEqual(expected, actual);
            Assert.Inconclusive("Verify the correctness of this test method.");
        }
    }
}
