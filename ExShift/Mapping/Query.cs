using ExShift.Util;
using System.Collections.Generic;
using System.Text.Json;

namespace ExShift.Mapping
{
    public class Query<T> where T : IPersistable, new()
    {
        private bool join;
        private Queue<QueryNode> queryNodes;

        private Query()
        {
            queryNodes = new Queue<QueryNode>();
        }

        public static Query<T> Select()
        {
            Query<T> query = new Query<T>();
            return query;
        }

        public Query<T> Where(string whereClause)
        {
            QueryNode qn = new QueryNode(whereClause, QueryOperator.ROOT);
            queryNodes.Enqueue(qn);
            return this;
        }

        public Query<T> And(string expression)
        {
            QueryNode qn = new QueryNode(expression, QueryOperator.AND);
            queryNodes.Enqueue(qn);
            return this;
        }

        public Query<T> Or(string expression)
        {
            QueryNode qn = new QueryNode(expression, QueryOperator.OR);
            queryNodes.Enqueue(qn);
            return this;
        }

        public Query<T> Join()
        {
            join = true;
            return this;
        }

        public List<T> Run()
        {
            List<T> resultList = new List<T>();
            ObjectPackager objectPackager = new ObjectPackager(null);
            ExcelObjectMapper eom = new ExcelObjectMapper();
            eom.Initialize();

            foreach (string rawJson in eom.GetAll<T>())
            {
                JsonElement jsonElement = objectPackager.DeserializeTupel(rawJson);
                bool elementIsQualified = false;

                foreach (QueryNode qn in queryNodes)
                {
                    JsonElement jsonProperty = jsonElement.GetProperty(qn.Attribute);
                    
                    bool evaluationResult;
                    if (jsonProperty.ValueKind.Equals(JsonValueKind.Number))
                    {
                        jsonProperty.TryGetDouble(out double value);
                        evaluationResult = qn.EvaluateExpression(value);
                    }
                    else
                    {
                        evaluationResult = qn.EvaluateExpression(jsonProperty.GetString());
                    }

                    switch (qn.Operator)
                    {
                        case QueryOperator.ROOT:
                            elementIsQualified = evaluationResult;
                            break;

                        case QueryOperator.AND:
                            elementIsQualified = elementIsQualified && evaluationResult;
                            break;

                        case QueryOperator.OR:
                            elementIsQualified = elementIsQualified || evaluationResult;
                            break;
                    }
                }
                if (elementIsQualified)
                {
                    resultList.Add(objectPackager.Unpackage<T>(rawJson));
                }
            }
            return resultList;
        }
    }
}
