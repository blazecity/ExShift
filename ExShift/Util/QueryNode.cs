using System.Text.RegularExpressions;

namespace ExShift.Mapping
{
    public class QueryNode
    {
        public QueryOperator Operator { get; }
        public dynamic Attribute { get; }
        public dynamic Expected { get; }

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

        public bool EvaluateExpression(dynamic actual)
        {
            return Expected == actual;
        }
    }
}
