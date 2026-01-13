using NUnit.Framework;

namespace GoContactSyncMod.UnitTests
{
    [TestFixture]
    public class SameEmailTests
    {

        [OneTimeSetUp]
        public void Init()
        {
        }

        [SetUp]
        public void SetUp()
        {
        }

        [OneTimeTearDown]
        public void TearDown()
        {
        }

        [Test]
        public void Test_IsSameEmail()
        {
            Assert.That(OutlookPropertiesUtils.IsSameEmail("john.smith@gmail.com", "johnsmith@gmail.com"));
            Assert.That(OutlookPropertiesUtils.IsSameEmail("john.smith+newsletter@gmail.com", "johnsmith@gmail.com"));
            Assert.That(OutlookPropertiesUtils.IsSameEmail("Joh.n.Smith+newsletter@gmail.com", "johnsmith+..@gmail.com"));
            Assert.That(OutlookPropertiesUtils.IsSameEmail("john.smith@gmail.com", "johnsmith@googlemail.com"));
            Assert.That(OutlookPropertiesUtils.IsSameEmail("john.smith+newsletter@gmail.com", "johnsmith@googlemail.com"));
            Assert.That(OutlookPropertiesUtils.IsSameEmail("Joh.n.Smith+newsletter@gmail.com", "johnsmith+..@googlemail.com"));
        }
    }
}
