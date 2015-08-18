using NUnit.Framework;

namespace ModbusExcel.Tests
{
    
    
    /// <summary>
    ///This is a test class for UpdateEventTest and is intended
    ///to contain all UpdateEventTest Unit Tests
    ///</summary>
    [TestFixture]
    public class UpdateEventTest
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
        ///A test for UpdateEvent Constructor
        ///</summary>
        [Test]
        public void UpdateEventConstructorTest()
        {
            UpdateEvent target = new UpdateEvent();
            Assert.Inconclusive("TODO: Implement code to verify target");
        }

        /// <summary>
        ///A test for Disconnect
        ///</summary>
        [Test]
        public void DisconnectTest()
        {
            UpdateEvent target = new UpdateEvent(); // TODO: Initialize to an appropriate value
            target.Disconnect();
            Assert.Inconclusive("A method that does not return a value cannot be verified.");
        }

        /// <summary>
        ///A test for UpdateNotify
        ///</summary>
        [Test]
        public void UpdateNotifyTest()
        {
            UpdateEvent target = new UpdateEvent(); // TODO: Initialize to an appropriate value
            target.UpdateNotify();
            Assert.Inconclusive("A method that does not return a value cannot be verified.");
        }

        /// <summary>
        ///A test for HeartbeatInterval
        ///</summary>
        [Test]
        public void HeartbeatIntervalTest()
        {
            UpdateEvent target = new UpdateEvent(); // TODO: Initialize to an appropriate value
            int expected = 0; // TODO: Initialize to an appropriate value
            int actual;
            target.HeartbeatInterval = expected;
            actual = target.HeartbeatInterval;
            Assert.AreEqual(expected, actual);
            Assert.Inconclusive("Verify the correctness of this test method.");
        }
    }
}
