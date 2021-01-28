using ExShift.Util;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text.Json;

namespace ExShift.Mapping
{
    public class Query<T> where T : IPersistable, new()
    {
        private List<QueryNode> queryNodes;

        private Query()
        {
            queryNodes = new List<QueryNode>();
        }

        public static Query<T> Select()
        {
            Query<T> query = new Query<T>();
            return query;
        }

        public Query<T> Where(string whereClause)
        {
            QueryNode qn = new QueryNode(whereClause, QueryOperator.ROOT);
            queryNodes.Add(qn);
            return this;
        }

        public Query<T> And(string expression)
        {
            QueryNode qn = new QueryNode(expression, QueryOperator.AND);
            queryNodes.Add(qn);
            return this;
        }

        public Query<T> Or(string expression)
        {
            QueryNode qn = new QueryNode(expression, QueryOperator.OR);
            queryNodes.Add(qn);
            return this;
        }

        public List<T> Run()
        {
            List<T> resultList = new List<T>();
            ObjectPackager objectPackager = new ObjectPackager();

            // If there are no query nodes return all elements in the table.
            if (queryNodes.Count == 0)
            {
                foreach (string rawJson in ExcelObjectMapper.GetAll<T>())
                {
                    resultList.Add(objectPackager.Unpackage<T>(rawJson));
                }
                return resultList;
            }

            // Only one existing query node
            if (queryNodes.Count == 1)
            {
                QueryNode qn = queryNodes[0];
                return SingleEvaluation(qn);
            }

            // Two query nodes
            if (queryNodes.Count == 2)
            {
                return PairwiseEvaluation(queryNodes[0], queryNodes[1]);
            }


            // More than two query nodes
            QueryNode qn1 = queryNodes[0];
            QueryNode qn2 = queryNodes[1];
            List<T> intermediateResult = PairwiseEvaluation(qn1, qn2);
            for (int i = 2; i < queryNodes.Count; i++)
            {
                QueryNode nextNode = queryNodes[i];
                if (nextNode.Operator == QueryOperator.OR)
                {
                    List<T> result = SingleEvaluation(nextNode);
                    intermediateResult = intermediateResult.Union(result).ToList();
                }
                else if (nextNode.Operator == QueryOperator.AND)
                {
                    List<T> result = new List<T>();
                    foreach (T element in intermediateResult)
                    {
                        PropertyInfo property = element.GetType().GetProperty(nextNode.Attribute);
                        if (nextNode.EvaluateExpression(property.GetValue(element)))
                        {
                            result.Add(element);
                        }
                    }
                    intermediateResult = result;
                }
            }
            return intermediateResult;
        }

        private List<T> SingleEvaluation(QueryNode qn)
        {
            List<T> resultList = new List<T>();
            ObjectPackager objectPackager = new ObjectPackager();
            if (ExcelObjectMapper.IsIndexed<T>(qn.Attribute))
            {
                return GetIndexedResults(qn);
            }
            else
            {
                foreach (string rawJson in ExcelObjectMapper.GetAll<T>())
                {
                    JsonElement jsonElement = ObjectPackager.DeserializeTupel(rawJson);

                    JsonElement jsonProperty;
                    try
                    {
                        jsonProperty = jsonElement.GetProperty(qn.Attribute);
                    }
                    catch (KeyNotFoundException)
                    {
                        return new List<T>();
                    }

                    bool elementIsQualified;
                    if (jsonProperty.ValueKind.Equals(JsonValueKind.Number))
                    {
                        jsonProperty.TryGetDouble(out double value);
                        elementIsQualified = qn.EvaluateExpression(value);
                    }
                    else
                    {
                        elementIsQualified = qn.EvaluateExpression(jsonProperty.GetString());
                    }

                    if (elementIsQualified)
                    {
                        resultList.Add(objectPackager.Unpackage<T>(rawJson));
                    }
                }
                return resultList;
            }
        }

        private List<T> PairwiseEvaluation(QueryNode firstNode, QueryNode secondNode)
        {
            List<T> resultList = new List<T>();
            ObjectPackager objectPackager = new ObjectPackager();
            
            if (ExcelObjectMapper.IsIndexed<T>(secondNode.Attribute))
            {
                Dictionary<string, List<int>> secondIndex = ExcelObjectMapper.FindIndex<T>(secondNode.Attribute);
                secondIndex.TryGetValue(secondNode.Expected.ToString(), out List<int> secondRows);
                if (ExcelObjectMapper.IsIndexed<T>(firstNode.Attribute))
                {
                    // First and second are indexed
                    Dictionary<string, List<int>> firstIndex = ExcelObjectMapper.FindIndex<T>(firstNode.Attribute);
                    firstIndex.TryGetValue(firstNode.Expected.ToString(), out List<int> firstRows);
                    List<int> subset;
                    if (secondNode.Operator == QueryOperator.AND)
                    {
                        subset = firstRows.Intersect(secondRows).ToList();
                    }
                    else
                    {
                        firstRows.Union(secondRows);
                        subset = firstRows;
                    }
                    foreach (int i in subset)
                    {
                        resultList.Add(ExcelObjectMapper.Find<T>(i));
                    }
                    return resultList;
                }

                // Only second is indexed
                foreach (int i in secondRows)
                {
                    string rawJson = ExcelObjectMapper.GetRawEntry<T>(i);
                    if (CheckRawJson(firstNode, rawJson))
                    {
                        resultList.Add(objectPackager.Unpackage<T>(rawJson));
                    }
                    return resultList;

                }
                return resultList;
            }

            else if (ExcelObjectMapper.IsIndexed<T>(firstNode.Attribute))
            {
                // Only first is indexed
                Dictionary<string, List<int>> firstIndex = ExcelObjectMapper.FindIndex<T>(secondNode.Attribute);
                firstIndex.TryGetValue(firstNode.Expected.ToString(), out List<int> firstRows);
                foreach (int i in firstRows)
                {
                    string rawJson = ExcelObjectMapper.GetRawEntry<T>(i);
                    if (CheckRawJson(firstNode, rawJson))
                    {
                        resultList.Add(objectPackager.Unpackage<T>(rawJson));
                    }
                    return resultList;
                }
            }

            else
            {
                // No index
                foreach (string rawJson in ExcelObjectMapper.GetAll<T>())
                {
                    bool elementIsQualified = false;
                    foreach (QueryNode qn in queryNodes)
                    {
                        elementIsQualified = CheckRawJson(qn, rawJson);
                        bool evaluationResult = CheckRawJson(qn, rawJson);
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
            return resultList;
        }

        private bool CheckRawJson(QueryNode qn, string rawJson)
        {
            JsonElement jsonElement = ObjectPackager.DeserializeTupel(rawJson);
            JsonElement jsonProperty = jsonElement.GetProperty(qn.Attribute);

            bool elementIsQualified;
            if (jsonProperty.ValueKind.Equals(JsonValueKind.Number))
            {
                jsonProperty.TryGetDouble(out double value);
                elementIsQualified = qn.EvaluateExpression(value);
            }
            else
            {
                elementIsQualified = qn.EvaluateExpression(jsonProperty.GetString());
            }
            return elementIsQualified;
        }

        private List<T> GetIndexedResults(QueryNode qn)
        {
            List<T> resultList = new List<T>();
            Dictionary<string, List<int>> idx = ExcelObjectMapper.FindIndex<T>(qn.Attribute);
            idx.TryGetValue(qn.Expected.ToString(), out List<int> rows);
            if (rows != null && rows.Count != 0)
            {
                foreach (int row in rows)
                {
                    T obj = ExcelObjectMapper.Find<T>(row);
                    resultList.Add(obj);
                }
            }
            return resultList;
        }
    }
}
