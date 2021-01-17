using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExShift.Util.Tests
{
    [TestClass]
    public class QueryNodeTests
    {
        [TestMethod("Standard case")]
        public void EvaluateExpressionTest()
        {
            QueryNode queryNode = new QueryNode("Person.vorname = 'Jan'", QueryOperator.AND);
            Assert.IsTrue(queryNode.EvaluateExpression("Jan"));
        }

        [TestMethod("Escape characters")]
        public void EvalutateExpressionEscapeCharTest()
        {
            QueryNode queryNode = new QueryNode("storage.item = 'Jens' tv'", QueryOperator.OR);
            Assert.IsTrue(queryNode.EvaluateExpression("Jens' tv"));
        }

        [TestMethod("Digits")]
        public void EvalutateExpressionDigitsTest()
        {
            QueryNode queryNode = new QueryNode("storage.count = 2", QueryOperator.OR);
            Assert.IsTrue(queryNode.EvaluateExpression(2));
        }

        [TestMethod("Equal sign")]
        public void EvalutateExpressionTestEqualSign()
        {
            QueryNode queryNode = new QueryNode("test.testField = '='", QueryOperator.ROOT);
            Assert.IsTrue(queryNode.EvaluateExpression("="));
        }
    }
}