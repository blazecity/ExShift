using System.Text.RegularExpressions;

namespace ExShift.Mapping
{
    /// <summary>
    /// Represents a node in a resolved query.
    /// All condition which are applied (where, or, and) are converted to a QueryNode.
    /// </summary>
    public class QueryNode
    {
        /// <summary>
        /// Boolean operator, <see cref="QueryOperator"/>
        /// </summary>
        public QueryOperator Operator { get; }
        /// <summary>
        /// Attribute which has to be looked at.
        /// </summary>
        public dynamic Attribute { get; }
        /// <summary>
        /// Expected value of the attribute
        /// </summary>
        public dynamic Expected { get; }

        /// <summary>
        /// Constructor for new QueryNode object.
        /// It also takes the condition, which is provided as a string and splits it apart.
        /// </summary>
        /// <param name="expression">Condition as string</param>
        /// <param name="queryOperator">Boolean operator as <see cref="QueryOperator"/></param>
        public QueryNode(string expression, QueryOperator queryOperator)
        {
            Regex rgx = new Regex("=");
            
            string[] splitExpression = rgx.Split(expression, 2);
            Attribute = splitExpression[0].Trim();
            Expected = splitExpression[1];
            if (!double.TryParse(Expected, out double number))
            {
                Expected = Regex.Match(Expected, @"(?<=').*(?=')").Value;
            }
            else
            {
                Expected = number;
            }
            Operator = queryOperator;
        }

        /// <summary>
        /// Compares the expected value with the actual value.
        /// </summary>
        /// <param name="actual">Actual value</param>
        /// <returns><c>true</c> if they match, else <c>false</c></returns>
        public bool EvaluateExpression(dynamic actual)
        {
            return Expected == actual;
        }
    }
}
