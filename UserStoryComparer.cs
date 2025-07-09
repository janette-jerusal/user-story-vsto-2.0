// TF-IDF + Cosine Similarity implementation
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace UserStorySimilarityAddIn
{
    public static class UserStoryComparer
    {
        public static DataTable CompareUserStories(DataTable tableA, DataTable tableB, double threshold)
        {
            var results = new DataTable();
            results.Columns.Add("Story A ID");
            results.Columns.Add("Story A Desc");
            results.Columns.Add("Story B ID");
            results.Columns.Add("Story B Desc");
            results.Columns.Add("Similarity Score");

            var tfidfA = TFIDF(tableA.AsEnumerable().Select(r => r["Desc"].ToString()).ToList());
            var tfidfB = TFIDF(tableB.AsEnumerable().Select(r => r["Desc"].ToString()).ToList());

            for (int i = 0; i < tfidfA.Count; i++)
            {
                for (int j = 0; j < tfidfB.Count; j++)
                {
                    double similarity = CosineSimilarity(tfidfA[i], tfidfB[j]);
                    if (similarity >= threshold)
                    {
                        var row = results.NewRow();
                        row["Story A ID"] = tableA.Rows[i]["ID"];
                        row["Story A Desc"] = tableA.Rows[i]["Desc"];
                        row["Story B ID"] = tableB.Rows[j]["ID"];
                        row["Story B Desc"] = tableB.Rows[j]["Desc"];
                        row["Similarity Score"] = Math.Round(similarity, 3);
                        results.Rows.Add(row);
                    }
                }
            }

            return results;
        }

        private static List<Dictionary<string, double>> TFIDF(List<string> documents)
        {
            var tfidfList = new List<Dictionary<string, double>>();
            var allTerms = new HashSet<string>();
            var docFrequencies = new Dictionary<string, int>();

            // Tokenize and build term frequencies
            var termFrequencies = documents.Select(doc =>
            {
                var tf = new Dictionary<string, double>();
                var words = doc.ToLower().Split(new[] { ' ', '.', ',', '!', '?', ';', ':', '-', '\n', '\r', '\t' }, StringSplitOptions.RemoveEmptyEntries);

                foreach (var word in words)
                {
                    if (!tf.ContainsKey(word)) tf[word] = 0;
                    tf[word]++;
                }

                foreach (var word in tf.Keys)
                {
                    allTerms.Add(word);
                }

                return tf;
            }).ToList();

            // Document frequency for each term
            foreach (var term in allTerms)
            {
                docFrequencies[term] = termFrequencies.Count(tf => tf.ContainsKey(term));
            }

            // Compute TF-IDF
            for (int i = 0; i < documents.Count; i++)
            {
                var tfidf = new Dictionary<string, double>();
                foreach (var term in allTerms)
                {
                    double tf = termFrequencies[i].ContainsKey(term) ? termFrequencies[i][term] : 0;
                    double idf = Math.Log((double)documents.Count / (1 + docFrequencies[term]));
                    tfidf[term] = tf * idf;
                }
                tfidfList.Add(tfidf);
            }

            return tfidfList;
        }

        private static double CosineSimilarity(Dictionary<string, double> vec1, Dictionary<string, double> vec2)
        {
            var commonTerms = vec1.Keys.Intersect(vec2.Keys);
            double dotProduct = commonTerms.Sum(term => vec1[term] * vec2[term]);
            double normA = Math.Sqrt(vec1.Values.Sum(val => val * val));
            double normB = Math.Sqrt(vec2.Values.Sum(val => val * val));

            if (normA == 0 || normB == 0) return 0;
            return dotProduct / (normA * normB);
        }
    }
}
