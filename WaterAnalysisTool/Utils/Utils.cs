using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WaterAnalysisTool.Utils
{
    class Utils
    {
        #region Public Methods
        public static int LevenshteinDistance(String s, String t)
        {
            if(s.Equals("") || s == null)
            {
                if (t.Equals("") || t == null)
                    return 0;

                return t.Length;
            }

            if (t.Equals("") || t == null)
                return s.Length;

            int n = s.Length;
            int m = t.Length;
            int[,] d = new int[n + 1, m + 1];

            for(int i = 1; i <= n; i++)
            {
                for(int j = 1; j <=m; j++)
                {
                    int cost = (t[j - 1] == s[i - 1]) ? 0 : 1;
                    int min1 = d[i - 1, j] + 1;
                    int min2 = d[i, j - 1] + 1;
                    int min3 = d[i - 1, j - 1] + cost;
                    d[i, j] = Math.Min(Math.Min(min1, min2), min3);
                }
            }

            return d[n, m];
        }
        
        public static int LongestCommonSubstring(String s, String t)
        {
            if(s.Equals("") || s == null)
                return 0;
            
            if(t.Equals("") || s == null)
                return 0;
            
            int n = s.Length;
            int m = t.Length;
            int maxlen = 0;
            int[,] d = new int[n, m];
            
            for(int i = 0; i < n; i++)
            {
                for(int j = 0; j < m; j++)
                {
                    if(s[i] != t[j])
                        d[i, j] = 0;
                    
                    else
                    {
                        if(i == 0 || j == 0)
                            d[i, j] = 1;
                        
                        else
                            d[i, j] = d[i - 1, j - 1] + 1;
                        
                        if(d[i, j] > maxlen)
                            maxlen = d[i, j];
                    }
                }
            }
            
            return maxlen;
        }
        #endregion
    }
}
