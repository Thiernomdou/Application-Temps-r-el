using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace VisibiliteCasting
{
    public class Rendement
    {
        public int verresLances;
        public int verresBons;
        public string date;
        public double yield;
        public int[] topDefauts = new int[5] { 0, 0, 0, 0, 0 };
        public decimal[] topDefautsPrc = new decimal[5] { 0, 0, 0, 0, 0 };
        public string[] topNomDefauts = new string[5] { "", "", "", "", "" };
        public decimal[] Heure = new decimal[] { };
        DateTime heure = DateTime.Now;

        public int[] tabValeursDefauts = new int[22];

        public double getRendement()
        {
            
            yield = ((double)verresBons / (double)verresLances) * 100;
            yield = Math.Round(yield, 2);
            return yield;
        }

        


        public void calculTopDefauts(string[] nomDefauts)
        {
            int j = 0, k = 0, l = 0, m = 0, n = 0;
            topDefauts[0] = 0;
            topDefauts[1] = 0;
            topDefauts[2] = 0;
            topDefauts[3] = 0;
            topDefauts[4] = 0;

            for (int i = 0; i < 22; i++)
            {
                if (topDefauts[j] < tabValeursDefauts[i])
                {
                    topDefauts[j] = tabValeursDefauts[i];
                    topDefautsPrc[j] = ((decimal)tabValeursDefauts[i] / (decimal)verresLances) * 100;
                    topNomDefauts[j] = nomDefauts[i];
                    k = i;
                }
            }

            j = j + 1;
            for (int i = 0; i < 22; i++)
            {
                if (topDefauts[j] < tabValeursDefauts[i] && i != k)
                {
                    topDefauts[j] = tabValeursDefauts[i];
                    topDefautsPrc[j] = ((decimal)tabValeursDefauts[i] / (decimal)verresLances) * 100;
                    topNomDefauts[j] = nomDefauts[i];
                    l = i;
                }
            }

            j = j + 1;
            for (int i = 0; i < 22; i++)
            {
                if (topDefauts[j] < tabValeursDefauts[i] && i != k && i != l)
                {
                    topDefauts[j] = tabValeursDefauts[i];
                    topDefautsPrc[j] = ((decimal)tabValeursDefauts[i] / (decimal)verresLances) * 100;
                    topNomDefauts[j] = nomDefauts[i];
                    m = i;
                }
            }

            j = j + 1;
            for (int i = 0; i < 22; i++)
            {
                if (topDefauts[j] < tabValeursDefauts[i] && i != k && i != l && i != m)
                {
                    topDefauts[j] = tabValeursDefauts[i];
                    topDefautsPrc[j] = ((decimal)tabValeursDefauts[i] / (decimal)verresLances) * 100;
                    topNomDefauts[j] = nomDefauts[i];
                    n = i;
                }
            }

            j = j + 1;
            for (int i = 0; i < 22; i++)
            {
                if (topDefauts[j] < tabValeursDefauts[i] && i != k && i != l && i != m && i != n)
                {
                    topDefauts[j] = tabValeursDefauts[i];
                    topDefautsPrc[j] = ((decimal)tabValeursDefauts[i] / (decimal)verresLances) * 100;
                    topNomDefauts[j] = nomDefauts[i];
                }
            }
        }
        public void Calcultop3DefautHeure()
        {
            for (int i = 0; i < 10; i++)
            {
                if (verresLances != 0)
                    Heure[i] = (tabValeursDefauts[i] / verresLances) * 100;
                else
                    Heure[i] = 0;
            }
        }


        public Rendement( int verreL, int verreB)
        {
            
            this.verresBons = verreB;
            this.verresLances = verreL;
        }



        public Rendement()
        {
        }


    }
}
