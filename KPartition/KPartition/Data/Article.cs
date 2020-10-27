namespace KPartition.Data
{
    public struct Article
    {
        internal readonly int WeightInGrams; // interval is 100g – 1kg

        public Article(int weightInGrams)
        {
            WeightInGrams = weightInGrams;
        }

        public override string ToString() => WeightInGrams.ToString();
    }
}