// TF-IDF + Cosine Similarity implementation
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;

namespace UserStorySimilarityAddIn
{
    public class UserStoryComparer
    {
        public DataTable Compare(string path1, string path2, double threshold)
        {
            var table1 = LoadExcel(path1);
            var table2 = LoadExcel(path2);

            var results = new DataTable();
            results.Columns.Add("Story A ID");
            results.Columns.Add("Story A Desc");
            results.Columns.Add("Story B ID");
            results.Columns.Add("Story B Desc");
            results.Columns.Add("Similarity Score");

            var descs1 = table1.AsEnumerable().Select(row => row["Desc"].ToString()).ToList();
            var descs2 = table2.AsEnumerable().Select(row => row["Desc"].ToString()).ToList();

            var vectorizer = new SimpleTFIDF();
            vectorizer.Fit(descs1.Concat(descs2).ToList());

            var vectors1 = descs1.Select(d => vectorizer.Transform(d)).ToList();
            var vectors2 = descs2.Select(d => vectorizer.Transform(d)).ToList();

            for (int i = 0; i < vectors1.Count; i++)
            {
                for (int j = 0; j < vectors2.Count; j++)
                {
                    double score = CosineSimilarity(vectors1[i], vectors2[j]);
                    if (score >= threshold)
                    {
                        results.Rows.Add(
                            table1.Rows[i]["ID"],
                            table1.Rows[i]["Desc"],
                            table2.Rows[j]["ID"],
                            table2.Rows[j]["Desc"],
                            Math.Round(score, 3)
                        );
                    }
                }
            }

            return results;
        }

        private double CosineSimilarity(Dictionary<string, double> vec1, Dictionary<string, double> vec2)
        {
            var allKeys = new HashSet<string>(vec1.Keys.Concat(vec2.Keys));
            double dot = 0, norm1 = 0, norm2 = 0;

            foreach (var key in allKeys)
            {
                double v1 = vec1.ContainsKey(key) ? vec1[key] : 0;
                double v2 = vec2.ContainsKey(key) ? vec2[key] : 0;
                dot += v1 * v2;
                norm1 += v1 * v1;
                norm2 += v2 * v2;
            }

            return dot / (Math.Sqrt(norm1) * Math.Sqrt(norm2) + 1e-10);
        }

        private DataTable LoadExcel(string filePath)
        {
            var app = new Microsoft.Office.Interop.Excel.Application();
            Workbook wb = app.Workbooks.Open(filePath);
            Worksheet ws = wb.Sheets[1];
            Range usedRange = ws.UsedRange;

            object[,] data = usedRange.Value2;
            DataTable dt = new DataTable();
            dt.Columns.Add("ID");
            dt.Columns.Add("Desc");

            for (int row = 2; row <= data.GetLength(0); row++)
            {
                string id = data[row, 1]?.ToString();
                string desc = data[row, 2]?.ToString();
                if (!string.IsNullOrEmpty(id) && !string.IsNullOrEmpty(desc))
                    dt.Rows.Add(id, desc);
            }

            wb.Close(false);
            app.Quit();
            return dt;
        }
    }

    // Simple TF-IDF vectorizer
    public class SimpleTFIDF
    {
        private Dictionary<string, double> idf = new Dictionary<string, double>();
        private List<string[]> documents = new List<string[]>();

        public void Fit(List<string> texts)
        {
            int N = texts.Count;
            var df = new Dictionary<string, int>();

            foreach (var text in texts)
            {
                var tokens = Tokenize(text);
                documents.Add(tokens);
                foreach (var word in tokens.Distinct())
                    df[word] = df.GetValueOrDefault(word, 0) + 1;
            }

            foreach (var kv in df)
                idf[kv.Key] = Math.Log((double)N / (kv.Value + 1));
        }

        public Dictionary<string, double> Transform(string text)
        {
            var tf = new Dictionary<string, double>();
            var tokens = Tokenize(text);
            foreach (var word in tokens)
                tf[word] = tf.GetValueOrDefault(word, 0) + 1;

            int total = tokens.Length;
            var tfidf = new Dictionary<string, double>();
            foreach (var kv in tf)
            {
                double idfVal = idf.GetValueOrDefault(kv.Key, 0);
                tfidf[kv.Key] = (kv.Value / total) * idfVal;
            }

            return tfidf;
        }

        private string[] Tokenize(string text)
        {
            return Regex.Split(text.ToLower(), @"\W+").Where(w => w.Length > 1).ToArray();
        }
    }

    public static class Extensions
    {
        public static TValue GetValueOrDefault<TKey, TValue>(this Dictionary<TKey, TValue> dict, TKey key, TValue def = default)
        {
            return dict.ContainsKey(key) ? dict[key] : def;
        }
    }
}

