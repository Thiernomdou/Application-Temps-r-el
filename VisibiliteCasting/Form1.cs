using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace VisibiliteCasting
{
    public partial class RendementVerres : Form
    {
        

        SqlConnection myConnection;
        SqlCommand myCommand;
        string myConnectionString;
        string strRequete;
        DateTime DateSQL = DateTime.Now;
        int vartest = 0;

        




        public RendementVerres()
        {
            InitializeComponent();
           

        }


        private void RendementsVerres(object sender, EventArgs e)
        {

            //variables locales qui permettra de récupérer les valeurs de la table
            string DateP2;
            double GLances1 = 0, GBons1 = 0, GLances2 = 0, GBons2 = 0, GLances3 = 0, GBons3 = 0, GLances4 = 0, GBons4 = 0;
            double GLances5 = 0, GBons5 = 0, GLances6 = 0, GBons6 = 0, GLances7 = 0, GBons7 = 0, GLances8 = 0, GBons8 = 0;
            double rendement1 = 0, rendement2 = 0, rendement3 = 0, rendement4 = 0, rendement5 = 0, rendement6 = 0, rendement7 = 0, rendement8 = 0;
            DateTime heure = DateTime.Now;
            DateSQL = DateTime.Now; //Obtention de la date

            //------------------------------------------------------------------------------------------------------

            //EQUIPE DU MATIN
            //------------------------------------------------------------------------------------------------------


            
            if (DateTime.Now.Hour > 5 && DateTime.Now.Hour < 24)
            {
                //1.6


                DateSQL = DateTime.Now;
               
                DateP2 = DateSQL.ToString("yyyy-MM-dd");
                hp1.Text = heure.ToShortTimeString();
                //Code permettant d'acquérir les valeurs globales
                myConnectionString = "user id=OFGDIJON; password=eb71e22c2a851c491d07b700ad6742811f2e3d34fb4a71b0102b30702c77a461; data source=dijon-lispfr10";
                strRequete = "Select CONVERT(int,sum(CASTING_LAUNCH)) as CASTINGLances,CONVERT(int,sum(CASTING_GOOD)) as CASTINGBons FROM [DIJON_FLX_PROD].[dbo].[ESS_MES_RPT_Org_Production] WITH (nolock) WHERE CASTINGInspectionFinishOn between '" + DateP2 + " 05:30:00:00' and '" + DateP2 + " 13:00:00:00' and ProdIndex = '1.6' and Process <> '3T24'";

                try
                {
                    myConnection = new SqlConnection(myConnectionString);
                    myConnection.Open();
                    myCommand = new SqlCommand(strRequete, myConnection);
                    SqlDataReader mySqDataReader = myCommand.ExecuteReader();
                    while (mySqDataReader.Read())
                    {
                     
                        GLances1 = (int)mySqDataReader["CASTINGLances"];
                        GBons1 = (int)mySqDataReader["CASTINGBons"];
                    }
                }
                catch (Exception eMsg1)
                {
                    Console.WriteLine("Erreur durant l’execution de la requete : " + eMsg1.Message);
                }
                finally
                {
                    myConnection.Close();
                }

                if (GLances1 > 0 || GBons1 > 0)
                {
                    lp1.Text = GLances1.ToString();
                    bp1.Text = GBons1.ToString();
                    rendement1 = (Math.Round(((double)GBons1 / (double)GLances1), 4) * 100);
                    rp1.Text = rendement1.ToString() + "%";

                    


                }
                else
                { lp1.Text = bp1.Text = rp1.Text = "0"; }

                
                   
                

                if (rendement1 < 96.5)
                {
                    rp1.BackColor = Color.Red;
                    
                }
                else if (rendement1 >= 96.5 && rendement1 < 97)
                {
                    rp1.BackColor = Color.Orange;
                   
                }
                else
                {
                    rp1.BackColor = Color.YellowGreen;
                    
                }

                

                //-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

                //1.67

                DateSQL = DateTime.Now;
                DateP2 = DateSQL.ToString("yyyy-MM-dd");
                hp2.Text = heure.ToShortTimeString();
                //Code permettant d'acquérir les valeurs globales
                myConnectionString = "user id=OFGDIJON; password=eb71e22c2a851c491d07b700ad6742811f2e3d34fb4a71b0102b30702c77a461; data source=dijon-lispfr10";
                strRequete = "Select CONVERT(int,sum(CASTING_LAUNCH)) as CASTINGLances,CONVERT(int,sum(CASTING_GOOD)) as CASTINGBons FROM [DIJON_FLX_PROD].[dbo].[ESS_MES_RPT_Org_Production] WITH (nolock) WHERE CASTINGInspectionFinishOn between '" + DateP2 + " 05:30:00:00' and '" + DateP2 + " 13:00:00:00'  and ProdIndex = '1.67' ";

                try
                {
                    myConnection = new SqlConnection(myConnectionString);
                    myConnection.Open();
                    myCommand = new SqlCommand(strRequete, myConnection);
                    SqlDataReader mySqDataReader = myCommand.ExecuteReader();
                    while (mySqDataReader.Read())
                    {
                        GLances2 = (int)mySqDataReader["CASTINGLances"];
                        GBons2 = (int)mySqDataReader["CASTINGBons"];
                    }
                }
                catch (Exception eMsg1)
                {
                    Console.WriteLine("Erreur durant l’execution de la requete : " + eMsg1.Message);
                }
                finally
                {
                    myConnection.Close();
                }

                
                
                if (GLances2 > 0 || GBons2 > 0)
                {
                    lp2.Text = GLances2.ToString();
                    bp2.Text = GBons2.ToString();
                    rendement2 = (Math.Round(((double)GBons2 / (double)GLances2), 4) * 100);
                    rp2.Text = rendement2.ToString() + "%";
                }
                else
                {
                     lp2.Text = bp2.Text = rp2.Text = "0";
                }
                
                if (rendement2 < 95)
                {
                    rp2.BackColor = Color.Red;
                }
                else if (rendement2 >= 95 && rendement2 < 96)
                {
                    rp2.BackColor = Color.Orange;
                }
                else
                {
                    rp2.BackColor = Color.YellowGreen;
                }
           

            

            //------------------------------------------------------------------------------------------------------------------------------------------------------------

            //BCT

                DateSQL = DateTime.Now;
                DateP2 = DateSQL.ToString("yyyy-MM-dd");
                hp3.Text = heure.ToShortTimeString();
                
                //Code permettant d'acquérir les valeurs globales
                myConnectionString = "user id=OFGDIJON; password=eb71e22c2a851c491d07b700ad6742811f2e3d34fb4a71b0102b30702c77a461; data source=dijon-lispfr10";
                strRequete = "Select CONVERT(int,sum(CASTING_LAUNCH)) as CASTINGLances,CONVERT(int,sum(CASTING_GOOD)) as CASTINGBons FROM [DIJON_FLX_PROD].[dbo].[ESS_MES_RPT_Org_Production] WITH (nolock) WHERE CASTINGInspectionFinishOn between '" + DateP2 + " 05:30:00:00' and '" + DateP2 + " 13:00:00:00' and RPL like '3T%'  ";

                try
                {
                    myConnection = new SqlConnection(myConnectionString);
                    myConnection.Open();
                    myCommand = new SqlCommand(strRequete, myConnection);
                    SqlDataReader mySqDataReader = myCommand.ExecuteReader();
                    while (mySqDataReader.Read())
                    {
                        GLances3 = (int)mySqDataReader["CASTINGLances"];
                        GBons3 = (int)mySqDataReader["CASTINGBons"];
                    }
                }
                catch (Exception eMsg1)
                {
                    Console.WriteLine("Erreur durant l’execution de la requete : " + eMsg1.Message);
                }
                finally
                {
                    myConnection.Close();
                }

                if (GLances3 > 0 || GBons3 > 0)
                {
                    lp3.Text = GLances3.ToString();
                    bp3.Text = GBons3.ToString();
                    rendement3 = (Math.Round(((double)GBons3 / (double)GLances3), 4) * 100);
                    rp3.Text = rendement3.ToString() + "%";
                }
                else
                { lp3.Text = bp3.Text = rp3.Text = "0"; }

                if (rendement3 < 96.5)
                {
                    rp3.BackColor = Color.Red;
                }
                else if (rendement3 >= 96.5 && rendement3 < 97)
                {
                    rp3.BackColor = Color.Orange;
                }
                else
                {
                    rp3.BackColor = Color.YellowGreen;
                }

                //--------------------------------------------------------------------------------------------------------------------------------------------


                //GLOBAL

                DateSQL = DateTime.Now;
                DateP2 = DateSQL.ToString("yyyy-MM-dd");
                hp4.Text = heure.ToShortTimeString();
                
                //Code permettant d'acquérir les valeurs globales
                myConnectionString = "user id=OFGDIJON; password=eb71e22c2a851c491d07b700ad6742811f2e3d34fb4a71b0102b30702c77a461; data source=dijon-lispfr10";
                strRequete = "Select CONVERT(int,sum(CASTING_LAUNCH)) as CASTINGLances,CONVERT(int,sum(CASTING_GOOD)) as CASTINGBons FROM [DIJON_FLX_PROD].[dbo].[ESS_MES_RPT_Org_Production] WITH (nolock) WHERE CASTINGInspectionFinishOn between '" + DateP2 + " 05:30:00:00' and '" + DateP2 + " 13:00:00:00'";

                try
                {
                    myConnection = new SqlConnection(myConnectionString);
                    myConnection.Open();
                    myCommand = new SqlCommand(strRequete, myConnection);
                    SqlDataReader mySqDataReader = myCommand.ExecuteReader();
                    while (mySqDataReader.Read())
                    {
                        GLances4 = (int)mySqDataReader["CASTINGLances"];
                        GBons4 = (int)mySqDataReader["CASTINGBons"];
                    }
                }
                catch (Exception eMsg1)
                {
                    Console.WriteLine("Erreur durant l’execution de la requete : " + eMsg1.Message);
                }
                finally
                {
                    myConnection.Close();
                }

                if (GLances4 > 0 && GBons4 > 0)
                {
                    lpg.Text = GLances4.ToString();
                    bpg.Text = GBons4.ToString();
                    rendement4 = (Math.Round(((double)GBons4 / (double)GLances4), 4) * 100);
                    rpg.Text = rendement4.ToString() + "%";

                    int x = 0;
                    
                    for (int i = 0; i <= x; i++)
                    {
                      
                            this.graph1.Series["RendementsVerresMatin"].Points.AddXY(heure.ToShortTimeString(), rendement4);
                            this.graph1.Series["RendementsVerresMatin"].IsValueShownAsLabel = true;

                    }
                             

                }
                else
                { lpg.Text = bpg.Text = rpg.Text = "0"; }

                if (rendement4 < 96.5)
                {
                    rpg.BackColor = Color.Red;
                }
                else if (rendement4 >= 96.5 && rendement4 < 97)
                {
                    rpg.BackColor = Color.Orange;
                }
                else
                {
                    rpg.BackColor = Color.YellowGreen;
                }

                


       
            }

            else
            {

                hp1.Text = hp2.Text = hp3.Text = hp4.Text = heure.ToShortTimeString();

                
            }
            //------------------------------------------------------------------------------------------------------

            //EQUIPE D'APRES-MIDI
            //------------------------------------------------------------------------------------------------------

            //1.6
            if (DateTime.Now.Hour >= 5  && DateTime.Now.Hour < 24)
            {
                //Restructure d'un format de date pour la requeteSQL pour le global


                DateSQL = DateTime.Now;
                DateP2 = DateSQL.ToString("yyyy-MM-dd");
                hp5.Text = heure.ToShortTimeString();
                //Code permettant d'acquérir les valeurs globales
                myConnectionString = "user id=OFGDIJON; password=eb71e22c2a851c491d07b700ad6742811f2e3d34fb4a71b0102b30702c77a461; data source=dijon-lispfr10";
                strRequete = "Select CONVERT(int,sum(CASTING_LAUNCH)) as CASTINGLances,CONVERT(int,sum(CASTING_GOOD)) as CASTINGBons FROM [DIJON_FLX_PROD].[dbo].[ESS_MES_RPT_Org_Production] WITH (nolock) WHERE CASTINGInspectionFinishOn between '" + DateP2 + " 13:00:00:00' and '" + DateP2 + " 20:00:00:00'  and ProdIndex = '1.6' and RPL not like '3T%' ";

                try
                {
                    myConnection = new SqlConnection(myConnectionString);
                    myConnection.Open();
                    myCommand = new SqlCommand(strRequete, myConnection);
                    SqlDataReader mySqDataReader = myCommand.ExecuteReader();
                    while (mySqDataReader.Read())
                    {
                        GLances5 = (int)mySqDataReader["CASTINGLances"];
                        GBons5 = (int)mySqDataReader["CASTINGBons"];
                    }
                }
                catch (Exception eMsg1)
                {
                    Console.WriteLine("Erreur durant l’execution de la requete : " + eMsg1.Message);
                }
                finally
                {
                    myConnection.Close();
                }

                if (GLances5 > 0 || GBons5 > 0)
                {
                    lp4.Text = GLances5.ToString();
                    bp4.Text = GBons5.ToString();
                    rendement5 = (Math.Round(((double)GBons5 / (double)GLances5), 4) * 100);
                    rp4.Text = rendement5.ToString() + "%";

                    
                }
                else
                { lp4.Text = bp4.Text = rp4.Text = "0"; }

                if (rendement5 < 96.5)
                {
                    rp4.BackColor = Color.Red;
                }
                else if (rendement5 >= 96.5 && rendement5 < 97)
                {
                    rp4.BackColor = Color.Orange;
                }
                else
                {
                    rp4.BackColor = Color.YellowGreen;
                }
            

            //-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            //1.67


              //Restructure d'un format de date pour la requeteSQL pour le global


                DateSQL = DateTime.Now;
                DateP2 = DateSQL.ToString("yyyy-MM-dd");
                hp6.Text = heure.ToShortTimeString();
                //Code permettant d'acquérir les valeurs globales
                myConnectionString = "user id=OFGDIJON; password=eb71e22c2a851c491d07b700ad6742811f2e3d34fb4a71b0102b30702c77a461; data source=dijon-lispfr10";
                strRequete = "Select CONVERT(int,sum(CASTING_LAUNCH)) as CASTINGLances,CONVERT(int,sum(CASTING_GOOD)) as CASTINGBons FROM [DIJON_FLX_PROD].[dbo].[ESS_MES_RPT_Org_Production] WITH (nolock) WHERE CASTINGInspectionFinishOn between '" + DateP2 + " 13:00:00:00' and '" + DateP2 + " 20:00:00:00'  and ProdIndex = '1.67'";

                try
                {
                    myConnection = new SqlConnection(myConnectionString);
                    myConnection.Open();
                    myCommand = new SqlCommand(strRequete, myConnection);
                    SqlDataReader mySqDataReader = myCommand.ExecuteReader();
                    while (mySqDataReader.Read())
                    {
                        GLances6 = (int)mySqDataReader["CASTINGLances"];
                        GBons6 = (int)mySqDataReader["CASTINGBons"];
                    }
                }
                catch (Exception eMsg1)
                {
                    Console.WriteLine("Erreur durant l’execution de la requete : " + eMsg1.Message);
                }
                finally
                {
                    myConnection.Close();
                }

                if (GLances6 > 0 || GBons6 > 0)
                {
                    lp5.Text = GLances6.ToString();
                    bp5.Text = GBons6.ToString();
                    rendement6 = (Math.Round(((double)GBons6 / (double)GLances6), 4) * 100);
                    rp5.Text = rendement6.ToString() + "%";
                }
                else
                { lp5.Text = bp5.Text = rp5.Text = "0"; }

                if (rendement6 < 95)
                {
                    rp5.BackColor = Color.Red;
                }
                else if (rendement6 >= 95 && rendement6 < 96)
                {
                    rp5.BackColor = Color.Orange;
                }
                else
                {
                    rp5.BackColor = Color.YellowGreen;
                }
            

            //------------------------------------------------------------------------------------------------------------------------------------------------------------

            //BCT

              //Restructure d'un format de date pour la requeteSQL pour le global


                DateSQL = DateTime.Now;
                DateP2 = DateSQL.ToString("yyyy-MM-dd");
                hp7.Text = heure.ToShortTimeString();
                //Code permettant d'acquérir les valeurs globales
                myConnectionString = "user id=OFGDIJON; password=eb71e22c2a851c491d07b700ad6742811f2e3d34fb4a71b0102b30702c77a461; data source=dijon-lispfr10";
                strRequete = "Select CONVERT(int,sum(CASTING_LAUNCH)) as CASTINGLances,CONVERT(int,sum(CASTING_GOOD)) as CASTINGBons FROM [DIJON_FLX_PROD].[dbo].[ESS_MES_RPT_Org_Production] WITH (nolock) WHERE CASTINGInspectionFinishOn between '" + DateP2 + " 13:00:00:00' and '" + DateP2 + " 20:00:00:00'  and ProdIndex = '1.6' and RPL like '3T%'";

                try
                {
                    myConnection = new SqlConnection(myConnectionString);
                    myConnection.Open();
                    myCommand = new SqlCommand(strRequete, myConnection);
                    SqlDataReader mySqDataReader = myCommand.ExecuteReader();
                    while (mySqDataReader.Read())
                    {
                        GLances7 = (int)mySqDataReader["CASTINGLances"];
                        GBons7 = (int)mySqDataReader["CASTINGBons"];
                    }
                }
                catch (Exception eMsg1)
                {
                    Console.WriteLine("Erreur durant l’execution de la requete : " + eMsg1.Message);
                }
                finally
                {
                    myConnection.Close();
                }

                if (GLances7 > 0 && GBons7 > 0)
                {
                    lp6.Text = GLances7.ToString();
                    bp6.Text = GBons7.ToString();
                    rendement7 = (Math.Round(((double)GBons7 / (double)GLances7), 4) * 100);
                    rp6.Text = rendement7.ToString() + "%";
                }
                else
                { lp6.Text = bp6.Text = rp6.Text = "0"; }

                if (rendement7 < 96.5)
                {
                    rp6.BackColor = Color.Red;
                }
                else if (rendement7 >= 96.5 && rendement7 < 97)
                {
                    rp6.BackColor = Color.Orange;
                }
                else
                {
                    rp6.BackColor = Color.YellowGreen;
                }
           

            //--------------------------------------------------------------------------------------------------------------------------------------------

            //GLOBAL

              //Restructure d'un format de date pour la requeteSQL pour le global


                DateSQL = DateTime.Now;
                DateP2 = DateSQL.ToString("yyyy-MM-dd");
                hp8.Text = heure.ToShortTimeString();
                //Code permettant d'acquérir les valeurs globales
                myConnectionString = "user id=OFGDIJON; password=eb71e22c2a851c491d07b700ad6742811f2e3d34fb4a71b0102b30702c77a461; data source=dijon-lispfr10";
                strRequete = "Select CONVERT(int,sum(CASTING_LAUNCH)) as CASTINGLances,CONVERT(int,sum(CASTING_GOOD)) as CASTINGBons FROM [DIJON_FLX_PROD].[dbo].[ESS_MES_RPT_Org_Production] WITH (nolock) WHERE CASTINGInspectionFinishOn between '" + DateP2 + " 13:00:00:00' and '" + DateP2 + " 20:00:00:00'";

                try
                {
                    myConnection = new SqlConnection(myConnectionString);
                    myConnection.Open();
                    myCommand = new SqlCommand(strRequete, myConnection);
                    SqlDataReader mySqDataReader = myCommand.ExecuteReader();
                    while (mySqDataReader.Read())
                    {
                        GLances8 = (int)mySqDataReader["CASTINGLances"];
                        GBons8 = (int)mySqDataReader["CASTINGBons"];
                    }
                }
                catch (Exception eMsg1)
                {
                    Console.WriteLine("Erreur durant l’execution de la requete : " + eMsg1.Message);
                }
                finally
                {
                    myConnection.Close();
                }

                if (GLances8 > 0 && GBons8 > 0)
                {
                    lpg1.Text = GLances8.ToString();
                    bpg1.Text = GBons8.ToString();
                    rendement8 = (Math.Round(((double)GBons8 / (double)GLances8), 4) * 100);
                    rpg1.Text = rendement8.ToString() + "%";


                    int y = 0;

                    for (int i = 0; i <= y; i++)
                    {

                        this.graph2.Series["RendementsVerresAprès-Midi"].Points.AddXY(heure.ToShortTimeString(), rendement8);
                        this.graph2.Series["RendementsVerresAprès-Midi"].IsValueShownAsLabel = true;


                    }

                    

                }
                    else{ lpg1.Text = bpg1.Text = rpg1.Text = "0"; }

                if (rendement8 < 96.5)
                {
                    rpg1.BackColor = Color.Red;
                }
                else if (rendement8 >= 96.5 && rendement8 < 97)
                {
                    rpg1.BackColor = Color.Orange;
                }
                else
                {
                    rpg1.BackColor = Color.YellowGreen;
                }
            }

            else
            {

                hp5.Text = hp6.Text = hp7.Text = hp8.Text = heure.ToShortTimeString();
                
            }
        }



        //-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        // Remplissage des parois

        private void Remplis_Parois(object sender, EventArgs e)
        {

            //variables locales qui permettra de récupérer les valeurs de la table
            string DateP2;
            int GLances1 = 0, GLances2 = 0, GLances3 = 0, GLances4 = 0, GLances5 = 0, GLances6 = 0, GLances7 = 0, GLances8 = 0, GLances9 = 0, GLances10 = 0, GLances11 = 0, GLances12 = 0, GLances13 = 0, GLances14 = 0;
            DateTime heure = DateTime.Now;
            DateSQL = DateTime.Now; //Obtention de la date


            //------------------------------------------------------------------------------------------------------
            

            //EQUIPE DU MATIN
            //------------------------------------------------------------------------------------------------------



            //si l'heure 

            if (DateTime.Now.Hour > 5 && DateTime.Now.Hour < 24)
            {

                //les diamètres 65



                DateSQL = DateTime.Now;
                DateP2 = DateSQL.ToString("yyyy-MM-dd");
                hp9.Text = heure.ToShortTimeString();
                //Code permettant d'acquérir les valeurs globales
                myConnectionString = "user id=OFGDIJON; password=eb71e22c2a851c491d07b700ad6742811f2e3d34fb4a71b0102b30702c77a461; data source=dijon-lispfr10";
                strRequete = "SELECT CONVERT(int,sum(CASTING_LAUNCH)) as CASTINGLances FROM [DIJON_FLX_PROD].[dbo].[ESS_MES_RPT_Org_Production] where CASTINGInspectionFinishOn between '" + DateP2 + " 05:30:00:00' and '" + DateP2 + " 13:00:00:00' and   Diameter = '65' ";

                try
                {
                    myConnection = new SqlConnection(myConnectionString);
                    myConnection.Open();
                    myCommand = new SqlCommand(strRequete, myConnection);
                    SqlDataReader mySqDataReader = myCommand.ExecuteReader();
                    while (mySqDataReader.Read())
                    {
                        GLances3 = (int)mySqDataReader["CASTINGLances"];
                    }
                }
                catch (Exception eMsg1)
                {
                    Console.WriteLine("Erreur durant l’execution de la requete : " + eMsg1.Message);
                }
                finally
                {
                    myConnection.Close();
                }

                if (GLances3 > 0)
                {

                    qp1.Text = GLances3.ToString();

                }
                else { qp1.Text = "0"; }



                //les diamètres 70


                DateSQL = DateTime.Now;
                DateP2 = DateSQL.ToString("yyyy-MM-dd");
                hp10.Text = heure.ToShortTimeString();
                //Code permettant d'acquérir les valeurs globales
                myConnectionString = "user id=OFGDIJON; password=eb71e22c2a851c491d07b700ad6742811f2e3d34fb4a71b0102b30702c77a461; data source=dijon-lispfr10";
                strRequete = "SELECT CONVERT(int,sum(CASTING_LAUNCH)) as CASTINGLances FROM [DIJON_FLX_PROD].[dbo].[ESS_MES_RPT_Org_Production] where  CASTINGInspectionFinishOn between '" + DateP2 + " 05:30:00:00' and '" + DateP2 + " 13:00:00:00' and   Diameter = '70' ";

                try
                {
                    myConnection = new SqlConnection(myConnectionString);
                    myConnection.Open();
                    myCommand = new SqlCommand(strRequete, myConnection);
                    SqlDataReader mySqDataReader = myCommand.ExecuteReader();
                    while (mySqDataReader.Read())
                    {
                        GLances4 = (int)mySqDataReader["CASTINGLances"];
                    }
                }
                catch (Exception eMsg1)
                {
                    Console.WriteLine("Erreur durant l’execution de la requete : " + eMsg1.Message);
                }
                finally
                {
                    myConnection.Close();
                }

                if (GLances4 > 0)
                {

                    qp2.Text = GLances4.ToString();

                }
                else { qp2.Text = "0"; }


                //les diamètres 75



                DateSQL = DateTime.Now;
                DateP2 = DateSQL.ToString("yyyy-MM-dd");
                hp11.Text = heure.ToShortTimeString();
                //Code permettant d'acquérir les valeurs globales
                myConnectionString = "user id=OFGDIJON; password=eb71e22c2a851c491d07b700ad6742811f2e3d34fb4a71b0102b30702c77a461; data source=dijon-lispfr10";
                strRequete = "SELECT CONVERT(int,sum(CASTING_LAUNCH)) as CASTINGLances FROM [DIJON_FLX_PROD].[dbo].[ESS_MES_RPT_Org_Production] where  CASTINGInspectionFinishOn between '" + DateP2 + " 05:30:00:00' and '" + DateP2 + " 13:00:00:00' and   Diameter = '75' ";

                try
                {
                    myConnection = new SqlConnection(myConnectionString);
                    myConnection.Open();
                    myCommand = new SqlCommand(strRequete, myConnection);
                    SqlDataReader mySqDataReader = myCommand.ExecuteReader();
                    while (mySqDataReader.Read())
                    {
                        GLances5 = (int)mySqDataReader["CASTINGLances"];
                    }
                }
                catch (Exception eMsg1)
                {
                    Console.WriteLine("Erreur durant l’execution de la requete : " + eMsg1.Message);
                }
                finally
                {
                    myConnection.Close();
                }

                if (GLances5 > 0)
                {

                    qp3.Text = GLances5.ToString();

                }
                else { qp3.Text = "0"; }



                //les diamètres 80



                DateSQL = DateTime.Now;
                DateP2 = DateSQL.ToString("yyyy-MM-dd");
                hp12.Text = heure.ToShortTimeString();
                //Code permettant d'acquérir les valeurs globales
                myConnectionString = "user id=OFGDIJON; password=eb71e22c2a851c491d07b700ad6742811f2e3d34fb4a71b0102b30702c77a461; data source=dijon-lispfr10";
                strRequete = "SELECT CONVERT(int,sum(CASTING_LAUNCH)) as CASTINGLances FROM [DIJON_FLX_PROD].[dbo].[ESS_MES_RPT_Org_Production] where  CASTINGInspectionFinishOn between '" + DateP2 + " 05:30:00:00' and '" + DateP2 + " 13:00:00:00' and   Diameter = '80' ";

                try
                {
                    myConnection = new SqlConnection(myConnectionString);
                    myConnection.Open();
                    myCommand = new SqlCommand(strRequete, myConnection);
                    SqlDataReader mySqDataReader = myCommand.ExecuteReader();
                    while (mySqDataReader.Read())
                    {
                        GLances6 = (int)mySqDataReader["CASTINGLances"];
                    }
                }
                catch (Exception eMsg1)
                {
                    Console.WriteLine("Erreur durant l’execution de la requete : " + eMsg1.Message);
                }
                finally
                {
                    myConnection.Close();
                }

                if (GLances6 > 0)
                {

                    qp4.Text = GLances6.ToString();

                }
                else { qp4.Text = "0"; }




                //GLOBAL

                

                    DateSQL = DateTime.Now;
                    DateP2 = DateSQL.ToString("yyyy-MM-dd");
                    hp17.Text = heure.ToShortTimeString();
                    //Code permettant d'acquérir les valeurs globales
                    myConnectionString = "user id=OFGDIJON; password=eb71e22c2a851c491d07b700ad6742811f2e3d34fb4a71b0102b30702c77a461; data source=dijon-lispfr10";
                    strRequete = "SELECT CONVERT(int,sum(CASTING_LAUNCH)) as CASTINGLances FROM [DIJON_FLX_PROD].[dbo].[ESS_MES_RPT_Org_Production] where  CASTINGInspectionFinishOn between '" + DateP2 + " 05:30:00:00' and '" + DateP2 + " 13:00:00:00' ";

                    try
                    {
                        myConnection = new SqlConnection(myConnectionString);
                        myConnection.Open();
                        myCommand = new SqlCommand(strRequete, myConnection);
                        SqlDataReader mySqDataReader = myCommand.ExecuteReader();
                        while (mySqDataReader.Read())
                        {
                            GLances7 = (int)mySqDataReader["CASTINGLances"];
                        }
                    }
                    catch (Exception eMsg1)
                    {
                        Console.WriteLine("Erreur durant l’execution de la requete : " + eMsg1.Message);
                    }
                    finally
                    {
                        myConnection.Close();
                    }

                    if (GLances7 > 0)
                    {
                        int attendu = 1300;
                        qpg.Text = GLances7.ToString();

                        if (GLances7 < attendu)
                        {
                            qpg.BackColor = Color.Red;
                        }

                        else
                        {
                            qpg.BackColor = Color.YellowGreen;
                        }

                        int z = 0;

                        for (int i = 0; i <= z; i++)
                        {

                        this.graph3.Series["QuantitéParoisMatin"].Points.AddXY(heure.ToShortTimeString(), GLances7);
                        this.graph3.Series["QuantitéParoisMatin"].IsValueShownAsLabel = true;


                        }

                        z += 1;

                    }
                        else { qpg.Text = "0"; qpg.BackColor = Color.Red; }

            }

                else
                {

                     hp9.Text = hp10.Text = hp11.Text = hp12.Text = hp17.Text  = heure.ToShortTimeString();
                      
                }

                //------------------------------------------------------------------------------------------------------

                //EQUIPE D'APRES MIDI
                //------------------------------------------------------------------------------------------------------



                //les diamètres 55
                if (DateTime.Now.Hour >= 5 && DateTime.Now.Hour < 24)
                {
                   //les diamètres 65


                    DateSQL = DateTime.Now;
                    DateP2 = DateSQL.ToString("yyyy-MM-dd");
                    hp13.Text = heure.ToShortTimeString();
                    //Code permettant d'acquérir les valeurs globales
                    myConnectionString = "user id=OFGDIJON; password=eb71e22c2a851c491d07b700ad6742811f2e3d34fb4a71b0102b30702c77a461; data source=dijon-lispfr10";
                    strRequete = "SELECT CONVERT(int,sum(CASTING_LAUNCH)) as CASTINGLances FROM [DIJON_FLX_PROD].[dbo].[ESS_MES_RPT_Org_Production] where  CASTINGInspectionFinishOn between '" + DateP2 + " 13:00:00:00' and '" + DateP2 + " 20:00:00:00' and   Diameter = '65' ";

                    try
                    {
                        myConnection = new SqlConnection(myConnectionString);
                        myConnection.Open();
                        myCommand = new SqlCommand(strRequete, myConnection);
                        SqlDataReader mySqDataReader = myCommand.ExecuteReader();
                        while (mySqDataReader.Read())
                        {
                            GLances10 = (int)mySqDataReader["CASTINGLances"];
                        }
                    }
                    catch (Exception eMsg1)
                    {
                        Console.WriteLine("Erreur durant l’execution de la requete : " + eMsg1.Message);
                    }
                    finally
                    {
                        myConnection.Close();
                    }

                    if (GLances10 > 0)
                    {

                        qp5.Text = GLances10.ToString();

                    }
                    else { qp5.Text = "0"; }



                    //les diamètres 70


                    DateSQL = DateTime.Now;
                    DateP2 = DateSQL.ToString("yyyy-MM-dd");
                    hp14.Text = heure.ToShortTimeString();
                    //Code permettant d'acquérir les valeurs globales
                    myConnectionString = "user id=OFGDIJON; password=eb71e22c2a851c491d07b700ad6742811f2e3d34fb4a71b0102b30702c77a461; data source=dijon-lispfr10";
                    strRequete = "SELECT CONVERT(int,sum(CASTING_LAUNCH)) as CASTINGLances FROM [DIJON_FLX_PROD].[dbo].[ESS_MES_RPT_Org_Production] where  CASTINGInspectionFinishOn between '" + DateP2 + " 13:00:00:00' and '" + DateP2 + " 20:00:00:00' and   Diameter = '70' ";

                    try
                    {
                        myConnection = new SqlConnection(myConnectionString);
                        myConnection.Open();
                        myCommand = new SqlCommand(strRequete, myConnection);
                        SqlDataReader mySqDataReader = myCommand.ExecuteReader();
                        while (mySqDataReader.Read())
                        {
                            GLances11 = (int)mySqDataReader["CASTINGLances"];
                        }
                    }
                    catch (Exception eMsg1)
                    {
                        Console.WriteLine("Erreur durant l’execution de la requete : " + eMsg1.Message);
                    }
                    finally
                    {
                        myConnection.Close();
                    }

                    if (GLances11 > 0)
                    {

                        qp6.Text = GLances11.ToString();

                    }
                    else { qp6.Text = "0"; }

                    //les diamètres 75



                    DateSQL = DateTime.Now;
                    DateP2 = DateSQL.ToString("yyyy-MM-dd");
                    hp15.Text = heure.ToShortTimeString();
                    //Code permettant d'acquérir les valeurs globales
                    myConnectionString = "user id=OFGDIJON; password=eb71e22c2a851c491d07b700ad6742811f2e3d34fb4a71b0102b30702c77a461; data source=dijon-lispfr10";
                    strRequete = "SELECT CONVERT(int,sum(CASTING_LAUNCH)) as CASTINGLances FROM [DIJON_FLX_PROD].[dbo].[ESS_MES_RPT_Org_Production] where  CASTINGInspectionFinishOn between '" + DateP2 + " 13:00:00:00' and '" + DateP2 + " 20:00:00:00' and   Diameter = '75' ";

                    try
                    {
                        myConnection = new SqlConnection(myConnectionString);
                        myConnection.Open();
                        myCommand = new SqlCommand(strRequete, myConnection);
                        SqlDataReader mySqDataReader = myCommand.ExecuteReader();
                        while (mySqDataReader.Read())
                        {
                            GLances12 = (int)mySqDataReader["CASTINGLances"];
                        }
                    }
                    catch (Exception eMsg1)
                    {
                        Console.WriteLine("Erreur durant l’execution de la requete : " + eMsg1.Message);
                    }
                    finally
                    {
                        myConnection.Close();
                    }

                    if (GLances12 > 0)
                    {

                        qp7.Text = GLances12.ToString();

                    }
                    else { qp7.Text = "0"; }



                    //les diamètres 80


                    DateSQL = DateTime.Now;
                    DateP2 = DateSQL.ToString("yyyy-MM-dd");
                    hp16.Text = heure.ToShortTimeString();
                    //Code permettant d'acquérir les valeurs globales
                    myConnectionString = "user id=OFGDIJON; password=eb71e22c2a851c491d07b700ad6742811f2e3d34fb4a71b0102b30702c77a461; data source=dijon-lispfr10";
                    strRequete = "SELECT CONVERT(int,sum(CASTING_LAUNCH)) as CASTINGLances FROM [DIJON_FLX_PROD].[dbo].[ESS_MES_RPT_Org_Production] where  CASTINGInspectionFinishOn between '" + DateP2 + " 13:00:00:00' and '" + DateP2 + " 20:00:00:00' and   Diameter = '80' ";

                    try
                    {
                        myConnection = new SqlConnection(myConnectionString);
                        myConnection.Open();
                        myCommand = new SqlCommand(strRequete, myConnection);
                        SqlDataReader mySqDataReader = myCommand.ExecuteReader();
                        while (mySqDataReader.Read())
                        {
                            GLances13 = (int)mySqDataReader["CASTINGLances"];
                        }
                    }
                    catch (Exception eMsg1)
                    {
                        Console.WriteLine("Erreur durant l’execution de la requete : " + eMsg1.Message);
                    }
                    finally
                    {
                        myConnection.Close();
                    }

                    if (GLances13 > 0)
                    {

                        qp8.Text = GLances13.ToString();

                    }
                    else { qp8.Text = "0"; }



                    //GLOBAL


                    DateSQL = DateTime.Now;
                    DateP2 = DateSQL.ToString("yyyy-MM-dd");
                    hp18.Text = heure.ToShortTimeString();
                    //Code permettant d'acquérir les valeurs globales
                    myConnectionString = "user id=OFGDIJON; password=eb71e22c2a851c491d07b700ad6742811f2e3d34fb4a71b0102b30702c77a461; data source=dijon-lispfr10";
                    strRequete = "SELECT CONVERT(int,sum(CASTING_LAUNCH)) as CASTINGLances FROM [DIJON_FLX_PROD].[dbo].[ESS_MES_RPT_Org_Production] where  CASTINGInspectionFinishOn between '" + DateP2 + " 13:00:00:00' and '" + DateP2 + " 20:00:00:00' ";

                    try
                    {
                        myConnection = new SqlConnection(myConnectionString);
                        myConnection.Open();
                        myCommand = new SqlCommand(strRequete, myConnection);
                        SqlDataReader mySqDataReader = myCommand.ExecuteReader();
                        while (mySqDataReader.Read())
                        {
                            GLances14 = (int)mySqDataReader["CASTINGLances"];
                        }
                    }
                    catch (Exception eMsg1)
                    {
                        Console.WriteLine("Erreur durant l’execution de la requete : " + eMsg1.Message);
                    }
                    finally
                    {
                        myConnection.Close();
                    }

                    if (GLances14 > 0)
                    {
                        int attendu = 1300;
                        qpg1.Text = GLances14.ToString();

                        if (GLances14 < attendu)
                        {
                            qpg1.BackColor = Color.Red;
                        }

                        else
                        {
                            qpg1.BackColor = Color.YellowGreen;
                        }


                        int b = 0;

                        for (int i = 0; i <= b; i++)
                        {

                        this.graph4.Series["QuantitéParoisAprès-Midi"].Points.AddXY(heure.ToShortTimeString(), GLances14);
                        this.graph4.Series["QuantitéParoisAprès-Midi"].IsValueShownAsLabel = true;


                        }

                        b += 1;

                    }
                        else { qpg1.Text = "0"; qpg1.BackColor = Color.Red; }


                }

                else
                {

                    hp13.Text = hp14.Text = hp15.Text = hp16.Text = hp18.Text = heure.ToShortTimeString();
                   
                }





        }
        


        private void myEvent(object source, EventArgs e)
        {
            vartest = vartest + 1;
            if (vartest == 2)
                vartest = 0;
           
            this.Rendements_verres.SelectedIndex = vartest;
        }

        //Event relançant les requetes
        private void myEvent2(object source, EventArgs e)
        {
            this.RendementsVerres(source, e);
            this.Remplis_Parois(source, e);
            
        }

        private void button_start_Click(object sender, EventArgs e)
        {

          

            System.Windows.Forms.Timer myTimer;
            // Create a timer
            myTimer = new System.Windows.Forms.Timer();
            // Tell the timer what to do when it elapses
            myTimer.Tick += new EventHandler(myEvent);
            // Set it to go off every 60 seconds
            myTimer.Interval = 30000; //secondes
            
            

            // And start it        
            myTimer.Enabled = true;

            System.Windows.Forms.Timer myTimer2;
            myTimer2 = new System.Windows.Forms.Timer();
            myTimer2.Tick += new EventHandler(myEvent2);
            myTimer2.Interval = 60000; //60000 millisecondes équivaut à 1 minute 
            myTimer2.Enabled = true;


         

            this.Rendements_verres.SelectedIndex = 0;
            myEvent2(this, e);
            


        }

        private void button__Click(object sender, EventArgs e)
        {
          

            this.Rendements_verres.SelectedIndex = 0;
            myEvent2(this, e);
        }

        private void RendementVerres_Load(object sender, EventArgs e)
        {

        }
    }


}
