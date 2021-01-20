using System.Text.RegularExpressions;

namespace ExShift.Util
{
    public class QueryNode
    {
        public QueryOperator Operator { get; }
        public dynamic Attribute { get; }
        public dynamic ExpressionResult { get; }

        public QueryNode(string expression, QueryOperator queryOperator)
        {
            Regex rgx = new Regex("=");
            
            string[] splitExpression = rgx.Split(expression, 2);
            Attribute = splitExpression[0].Trim();
            ExpressionResult = splitExpression[1];
            if (!double.TryParse(ExpressionResult, out double number))
            {
                ExpressionResult = Regex.Match(ExpressionResult, @"(?<=').*(?=')").Value;
            }
            else
            {
                ExpressionResult = number;
            }
            Operator = queryOperator;
        }

        public bool EvaluateExpression(dynamic actual)
        {
            return ExpressionResult == actual;
        }
    }
}
