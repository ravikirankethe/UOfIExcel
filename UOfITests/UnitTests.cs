using System.Collections.Generic;
using System.Linq;
using UOfISpace;

namespace UOfITests
{
    [TestClass]
    public class UnitTests
    {
        [TestMethod]
        public void TestSubLocationValidation()
        {
            var validLocs = new List<string> { "Indiana", "Mayo", "Purdue", "Florida" };
            Subrecipient sub = new Subrecipient();

            List<SubrecipientData> data = sub.getSubrecipientDataFromFile("SubawardBudgetExample1.xlsx");

            Assert.IsTrue(data != null && data.Count> 0 && data.All(o => validLocs.Any(w => w == o.SubLocation)));
        }

        [TestMethod]
        public void TestDataExists()
        {
            Subrecipient sub = new Subrecipient();

            List<SubrecipientData> data = sub.getSubrecipientDataFromFile("SubawardBudgetExample1.xlsx");

            Assert.IsTrue(data != null && data.Count> 0);
        }
    }
}