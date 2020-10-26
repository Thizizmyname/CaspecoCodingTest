using System.Collections.Generic;
using System.Linq;
using KPartition.Data;

namespace KPartition
{
    
    /* Box Packer is an implementation of the K-partition optimization problem.
     * The K-partition problem is an NP-hard problem
     * 
     * Provided down below is one of the simplest solutions
     * The greedy algorithm:
     * keep placing the largest article in the smallest bin until all articles are placed.
     * This provides a simple, fast alternative but might provide a large approximation error
     * from the theoretical optimal solution.
     *
     * This solution, utilises a priority queue to obtain the Box with the least weight.
     * This would provide a worst case complexity of O(n log n)
     * Unfortunately, .net does not provide a priority queue, or the underlying data structures
     * (e.g. binomial heap) usually used, So instead of spending longer time implementing a example structure,
     * or installing external dependencies, a simple solution based on a list is used which raises the
     * worst case complexity to O(n^2)
     *
     * There are solutions that have a lower approximation error, even a pseudo-polynomial solution
     * using dynamic programming, (although the one I looked at entails an exponential memory usage).
     * Most of these solutions though would be quite extensive work for a coding test.
     *
     */
    public static class BoxPacker
    {
        public static List<Box> Pack(List<Article> articles, int numBoxes)
        {
            // Algorithm cannot operate on non-positive amount of boxes
            if (numBoxes <= 0)
            {
                return new List<Box>();
            }
            
            //Creates a shallow-copy to nondestructively sort the elements
            var shallowCopy = articles.ToList();
            shallowCopy.Sort((article, article1) => article1.WeightInGrams - article.WeightInGrams);
            
            // Creates a priority queue and initialize it with empty boxes
            var pq = new DumbPriorityQueue<Box>(numBoxes);
            for (int i = 0; i < numBoxes; i++)
            {
                pq.Add(new Box());
            }

            // Greedy Algorithm
            foreach (var article in shallowCopy)
            {
                var minBox = pq.ExtractMin();
                minBox.Add(article);
                pq.Add(minBox);
            }

            return pq.ToList();
        }

    }
}