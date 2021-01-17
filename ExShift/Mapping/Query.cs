using ExShift.Util;
using System.Collections.Generic;

namespace ExShift.Mapping
{
    public class Query<T> where T : IPersistable
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
            ExcelObjectMapper eom = new ExcelObjectMapper();
            return resultList;
        }
    }
}
