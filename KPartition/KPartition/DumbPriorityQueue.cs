using System.Collections.Generic;

namespace KPartition
{
    public interface IWeighted
    {
        int Weight { get; }
    }
    
    /* Priority queue - Dumb version
     *
     * This is a very simple implementation of a priority queue.
     * This should be implemented as a binomial heap or other more optimized data structures.
     * Consequence of this is while GetMin still have a complexity of O(1), Additions have O(n) instead of O(log n)  
     * 
     */
    public class DumbPriorityQueue<T> where T : IWeighted
    {
        private List<T> pq;
        public DumbPriorityQueue(int cap = 0)
        {
            pq = new List<T>(cap);
        }

        public void Add(T elem)
        {

            var index = pq.FindIndex(e => e.Weight >= elem.Weight);
            if (index >= 0)
            {
                pq.Insert(index, elem);
            }
            else
            {
                pq.Add(elem);
            }
        }

        public T ExtractMin()
        {
            var elem = pq[0];
            pq.RemoveAt(0);
            return elem;
        }

        public List<T> ToList() => pq;
    }
}