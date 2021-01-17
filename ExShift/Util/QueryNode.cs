using System.Text.RegularExpressions;

namespace ExShift.Util
{
    public class QueryNode
    {
        public QueryOperator Operator { get; set; }
        public dynamic ExpressionResult { get; set; }
        public bool Subresult { get; set; }
        public bool EvalutationResult { get; set; }

        public QueryNode(string expression, QueryOperator queryOperator)
        {
            Regex rgx = new Regex("=");
            ExpressionResult = rgx.Split(expression, 2)[1];
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
            return EvalutationResult = ExpressionResult == actual;
        }
    }
}
