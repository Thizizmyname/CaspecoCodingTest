using System;
using System.Collections.Generic;

namespace KPartition.Data
{
    public class Box : IWeighted
    {
        internal int WeightSum;
        List<Article> BoxItems;

        public Box()
        {
            WeightSum = 0;
            BoxItems = new List<Article>();
        }

        public override string ToString()
        {
            var join = String.Join(", ", BoxItems);
            return String.Format("[ {0} ]  - {1}", join, WeightSum.ToString());
        }

        public void Add(Article article)
        {
            BoxItems.Add(article);
            WeightSum += article.WeightInGrams;
        }

        public int Weight => WeightSum;
    }
}