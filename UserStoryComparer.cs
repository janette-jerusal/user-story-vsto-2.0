// TF-IDF + Cosine Similarity implementation
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace UserStorySimilarityAddIn
{
    public static class UserStoryComparer
    {
        public static DataTable CompareUserStories(DataTable table1, DataTable table2, double threshold)
        {
            var desc1 = table1.AsEnumerable().Select(row => row.Field<string>("Desc")).ToList();
            var desc2 = table2.AsEnumerable().Select(row => row.Field<string>("Desc")).ToList();

            var allDescriptions = desc1.Concat(desc2).ToList();

            var tfidf = new TfIdfVectorizer();
            var tfidfMatrix = tfidf.FitTransform(allDescriptions);

            var tfidf1 = tfidfMatrix.Take(desc1.Count).ToList();
            var tfidf2 = tfidfMatrix.Skip(desc1.Count).ToList();

            var results = new DataTable();
            results.Columns.Add("Story A ID");
            results.Columns.Add("Story A Desc");
            results.Columns.Add("Story B ID");
            results.Columns.Add("Story B Desc");
            results.Columns.Add("Similarity Score");

            for (int i = 0; i < tfidf1.Count; i++)
            {
                for (int j = 0; j < tfidf2.Count; j++)
                {
                    double similarity = CosineSimilarity(tfidf1[i], tfidf2[j]);
                    if (similarity >= threshold)
                    {
                        results.Rows.Add(
                            table1.Rows[i]["ID"],
                            table1.Rows[i]["Desc"],
                            table2.Rows[j]["ID"],
                            table2.Rows[j]["Desc"],
                            Math.Round(similarity, 3).ToString()
                        );
                    }
                }
            }

            return results;
        }

        private static double CosineSimilarity(Dictionary<string, double> vec1, Dictionary<string, double> vec2)
        {
            var commonKeys = vec1.Keys.Intersect(vec2.Keys);
            double dot = commonKeys.Sum(k => vec1[k] * vec2[k]);

            double mag1 = Math.Sqrt(vec1.Values.Sum(v => v * v));
            double mag2 = Math.Sqrt(vec2.Values.Sum(v => v * v));

            if (mag1 == 0 || mag2 == 0) return 0.0;

            return dot / (mag1 * mag2);
        }
    }
}
