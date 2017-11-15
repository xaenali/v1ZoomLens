﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using ZOSAPI;
using OSAPI_NetHelper = ZOSAPI_NetHelper;
using ZOSAPI.Tools.General;
using ZOSAPI.Editors.MFE;
using ZOSAPI.Editors.MCE;
using ZOSAPI.Editors;
using ZOSAPI.Editors.LDE;
using ZOSAPI.SystemData;

namespace BXwithZemax
{
    class Program
    {
        public static double InputMax, InputMin, InputbeamDia, EPDConstrain, Mind1, Mind2, SelF1, SelF2, SelF3;
        public static List<double> Temp1 = new List<double>();
        public static List<double> Temp2 = new List<double>();
        //public static double[] Temp1 = new double[1000000];
        //public static double[] Temp2 = new double[1000000];
        public static IList<double> MaxtrackList = new List<double>();
        public static IList<double> F1List = new List<double>();
        public static IList<double> F2List = new List<double>();
        public static IList<double> F3List = new List<double>();
        public static double[] Mx = new double[1000000];
        public static double[] My = new double[1000000];
        public static List<double> MxList = new List<double>();
        public static List<double> MyList = new List<double>();
        public static double[] a1 = new double[1000000];
        public static double[] a2 = new double[1000000];
        public static double[] b1 = new double[1000000];
        public static double[] b2 = new double[1000000];
        public static double[] MxratioMy = new double[1000000];
        public static IList<double> templist = new List<double>();
        public static double[] Maxtrack = new double[1000000];
        public static List<double> focallength1 = new List<double>(); // Initialize array for focal length 1
        public static List<double> focallength2 = new List<double>(); // Initialize array for focal length 2
        public static List<double> focallength3 = new List<double>(); // Initialize array for focal length 3
        public static List<double> EPD1 = new List<double>(); // Initialize List for Entrance pupil Dia 1
        public static List<double> EPD2 = new List<double>(); // Initialize List for Entrance pupil Dia  2
        public static List<double> EPD3 = new List<double>(); // Initialize List for fEntrance pupil Dia  3



        static void Main(string[] args)
        {
            double[] d1forMx = new double[1000000];
            double[] d2forMx = new double[1000000];
            double[] d1forMy = new double[1000000];
            double[] d2forMy = new double[1000000];
            double[] d1forMxratioMy = new double[1000000];
            double[] d2forMxratioMy = new double[1000000];
            double[] d1forInputMax = new double[1000000];
            double[] d2forInputMax = new double[1000000];
            double[] d1forInputMin = new double[1000000];
            double[] d2forInputMin = new double[1000000];

            Console.WriteLine("Accessing Focal length database \n");

            getExcelFile();

            Console.WriteLine("Access Complete \n");


            int k, l, m, n = 0;

            while (true)
            {

                n = 0; // Make n(number of combination) = 0, in start and at the end of all the combination 

                Console.WriteLine("Enter Max Magnification upto 4 decimal point \n");

                while (!Double.TryParse(Console.ReadLine(), out InputMax))
                {
                    Console.WriteLine("Please only input numeric value \n\n");

                    Console.WriteLine("Enter Max Magnification upto 4 decimal point \n");
                }

                Console.WriteLine("\n");

                Console.WriteLine("Enter Min Magnification upto 4 decimal point \n");

                // Check for value other than numerics

                while (!Double.TryParse(Console.ReadLine(), out InputMin))
                {
                    Console.WriteLine("Please only input numeric value \n\n");

                    Console.WriteLine("Enter Min Magnification upto 4 decimal point \n");
                }

                Console.WriteLine("\nEnter Input Beam Diameter \n");

                while (!Double.TryParse(Console.ReadLine(), out InputbeamDia))
                {
                    Console.WriteLine("Please only input numeric value \n\n");

                    Console.WriteLine("Enter Input Beam Diameter \n");

                }

                //Introduce Inputbeamdia and EPD constrain

                EPDConstrain = (double)1.5 * InputbeamDia;

                Console.WriteLine("\n");

                Console.WriteLine("Enter 0 or 1 to see all combinations with each resutls or d1 and d2 for Max Mag respectively \n");

                string input = Console.ReadLine();

                Console.WriteLine("\n");

                switch (input)
                {
                    case "0":

                        //take f1, f2 and f3 to calculate Max magnification and compare with Mx and My magnifications

                        for (k = 0; k < focallength1.Count; k++) // take f1
                        {

                            for (l = 0; l < focallength2.Count; l++) // take f2
                            {


                                for (m = 0; m < focallength3.Count; m++) // take f3
                                {

                                    a1[m] = Math.Round((double)focallength1[k] + focallength2[l], 4);

                                    MxratioMy[m] = Math.Round((double)focallength1[k] / focallength3[m], 4);

                                    a2[m] = Math.Round((double)focallength2[l] + focallength3[m], 4);

                                    b1[m] = Math.Round((double)(focallength1[k] * focallength2[l]) / focallength3[m], 4);

                                    b2[m] = Math.Round((double)(focallength2[l] * focallength3[m]) / focallength1[k], 4);

                                    Mx[m] = Math.Round((double)-a2[m] / b2[m], 4);

                                    My[m] = Math.Round((double)-b1[m] / a1[m], 4);


                                    //comparison with Mx and My magnifications

                                    if ((Mx[m] > MxratioMy[m]) && (MxratioMy[m] > My[m]) && (InputMax <= Mx[m]) && (InputMin >= My[m]) && (InputMax > InputMin) && (EPD1[k] > EPDConstrain) && (EPD1[k] == EPD2[k]))
                                    {
                                        n = n + 1;

                                        Console.WriteLine("number of combination = {0} ", n);

                                        Console.WriteLine("Conditions satified for comination {0}", n);

                                        Console.WriteLine("take Mx as {0} and My as {1} with F1 as {2}, F2 as {3} and F3 as {4}", Mx[m], My[m], focallength1[k], focallength2[l], focallength3[m]);



                                        //Calculate d1 and d2 for the Input Max and Min Magnification

                                        d1forInputMax[m] = Math.Round((double)focallength1[k] + focallength2[l] + ((focallength1[k] * focallength2[l]) / (InputMax * focallength3[m])), 4);

                                        d2forInputMax[m] = Math.Round((double)focallength2[l] + focallength3[m] + ((focallength2[l] * focallength3[m] * InputMax) / (focallength1[l])), 4);

                                        // Check for very small negative distance value and convert them to zero

                                        if ((d1forInputMax[m] >= -0.012) && (d1forInputMax[m] < 0))
                                        {
                                            d1forInputMax[m] = 0;
                                        }
                                        else

                                            if ((d2forInputMax[m] >= -0.012) && (d2forInputMax[m] < 0))
                                            {
                                                d2forInputMax[m] = 0;
                                            }

                                        Console.WriteLine("The system has d1 = {0} and d2 = {1} for Max magnification Input = {2} ", d1forInputMax[m], d2forInputMax[m], InputMax);

                                        d1forInputMin[m] = Math.Round((double)focallength1[k] + focallength2[l] + ((focallength1[k] * focallength2[l]) / (InputMin * focallength3[m])), 4);

                                        d2forInputMin[m] = Math.Round((double)focallength2[l] + focallength3[m] + ((focallength2[l] * focallength3[m] * InputMin) / (focallength1[k])), 4);

                                        if ((d1forInputMin[m] >= -0.012) && (d1forInputMin[m] < 0))
                                        {
                                            d1forInputMin[m] = 0;
                                        }
                                        else

                                            if ((d2forInputMin[m] >= -0.012) && (d2forInputMin[m] < 0))
                                            {
                                                d2forInputMin[m] = 0;
                                            }

                                        Console.WriteLine("The system has d1 = {0} and d2 = {1} for Min magnification Input = {2} ", d1forInputMin[m], d2forInputMin[m], InputMin);


                                        // Calculate Max d1 and d2 for Max magnification

                                        d1forMx[m] = Math.Round((double)focallength1[k] + focallength2[l] + ((focallength1[k] * focallength2[l]) / (Mx[m] * focallength3[m])), 4);

                                        d2forMx[m] = Math.Round((double)focallength2[l] + focallength3[m] + ((focallength2[l] * focallength3[m] * Mx[m]) / (focallength1[k])), 4);

                                        if ((d1forMx[m] >= -0.012) && (d1forMx[m] < 0))
                                        {
                                            d1forMx[m] = 0;
                                        }
                                        else

                                            if ((d2forMx[m] >= -0.012) && (d2forMx[m] < 0))
                                            {
                                                d2forMx[m] = 0;
                                            }

                                        Console.WriteLine("The system has d1 = {0} and d2 = {1} for Maximum Magnification possible = {2} ", d1forMx[m], d2forMx[m], Mx[m]);


                                        // Calculate d1 and d2 for Minimum magnification

                                        d1forMy[m] = Math.Round((double)focallength1[k] + focallength2[l] + ((focallength1[k] * focallength2[l]) / (My[m] * focallength3[m])), 4);

                                        d2forMy[m] = Math.Round((double)focallength2[l] + focallength3[m] + ((focallength2[l] * focallength3[m] * My[m]) / (focallength1[k])), 4);

                                        if ((d1forMy[m] >= -0.012) && (d1forMy[m] < 0))
                                        {
                                            d1forMy[m] = 0;
                                        }
                                        else

                                            if ((d2forMy[m] >= -0.012) && (d2forMy[m] < 0))
                                            {
                                                d2forMy[m] = 0;
                                            }

                                        Console.WriteLine("The system has d1 = {0} and d2 = {1} for Mimimum Magnification possible = {2} ", d1forMy[m], d2forMy[m], My[m]);


                                        // Calculate Max track length and d1 and d2 for that

                                        Maxtrack[m] = Math.Round((double)focallength1[k] + 2 * focallength2[l] + focallength3[m] + (focallength2[l] * (((focallength3[m] * MxratioMy[m]) / focallength1[k]) + focallength1[k] / (focallength3[m] * MxratioMy[m]))), 4);


                                        Console.WriteLine("The total system length (d1+d2) = {0} for Magnification = {1} with F1 = {2}, F2 = {3} and F3 = {4} ", Maxtrack[m], MxratioMy[m], focallength1[k], focallength2[l], focallength3[m]);

                                        d1forMxratioMy[m] = Math.Round((double)focallength1[k] + focallength2[l] + ((focallength1[k] * focallength2[l]) / (MxratioMy[m] * focallength3[m])), 4);

                                        d2forMxratioMy[m] = Math.Round((double)focallength2[l] + focallength3[m] + ((focallength2[l] * focallength3[m] * MxratioMy[m]) / (focallength1[k])), 4);

                                        if ((d1forMxratioMy[m] >= -0.012) && (d1forMxratioMy[m] < 0))
                                        {
                                            d1forMxratioMy[m] = 0;
                                        }
                                        else

                                            if ((d2forMxratioMy[m] >= -0.012) && (d2forMxratioMy[m] < 0))
                                            {
                                                d2forMxratioMy[m] = 0;
                                            }

                                        Console.WriteLine("The system has maximum length with d1 = {0} and d2 = {1} and Magnification = {2} ", d1forMxratioMy[m], d2forMxratioMy[m], MxratioMy[m]);


                                    }


                                    else

                                        if ((MxratioMy[m] > Mx[m]) || (My[m] > MxratioMy[m]) || (InputMax > Mx[m]) || (InputMin < My[m]) || (InputMax < InputMin) || (EPD1[k] < EPDConstrain) || (EPD1[k] != EPD2[k]))
                                        {
                                            n = n + 1;

                                            Console.WriteLine("number of combination = {0} ", n);

                                            Console.WriteLine("Conditions didn't satified for comination {0}", n);

                                            Console.WriteLine("Can't choose InputMax = {0} and InputMin = {1} as InputMax ({0}) > Calculated Mx {2} or InputMin ({1}) < calculated My {3} with F1 as {4}, F2 as {5} and F3 as {6} ", InputMax, InputMin, Mx[m], My[m], focallength1[k], focallength2[l], focallength3[m]);

                                        }

                                }

                            }

                        }


                        break;

                    case "1":

                        perm(focallength1, focallength2, focallength3, EPD1, EPD2, EPD3);

                        break;

                    default:

                        Console.WriteLine("Please choose from (0) or (1) \n");

                        break;
                }

            }

        }

        public static void getExcelFile()
        {

            double F1rngCount;
            double F2rngCount;
            double F3rngCount;
            double EPD1rngCount;
            double EPD2rngCount;
            double EPD3rngCount;

            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();

            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Users\Ali\Source\Repos\v1ZoomLens\BXwithZemax\BXwithZemax\focal2.xlsx");

            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];

            Excel.Range xlRange = xlWorksheet.UsedRange;

            // get used range of column Focal length columns

            Excel.Range F1range = xlWorksheet.UsedRange.Columns["A", Type.Missing];

            Excel.Range F2range = xlWorksheet.UsedRange.Columns["C", Type.Missing];

            Excel.Range F3range = xlWorksheet.UsedRange.Columns["E", Type.Missing];

            Excel.Range EPD1range = xlWorksheet.UsedRange.Columns["B", Type.Missing];

            Excel.Range EPD2range = xlWorksheet.UsedRange.Columns["D", Type.Missing];

            Excel.Range EPD3range = xlWorksheet.UsedRange.Columns["F", Type.Missing];



            // get number of used rows in column A, B and C

            F1rngCount = F1range.Rows.Count;

            F2rngCount = F2range.Rows.Count;

            F3rngCount = F3range.Rows.Count;

            EPD1rngCount = EPD1range.Rows.Count;

            EPD2rngCount = EPD2range.Rows.Count;

            EPD3rngCount = EPD3range.Rows.Count;


            // iterate over column A, C and E's used row count and store values to the list for Focal lengths

            for (int i = 2; i <= F1rngCount; i++)
            {
                focallength1.Add(xlWorksheet.Cells[i, "A"].Value());
            }

            for (int j = 2; j <= F2rngCount; j++)
            {
                focallength2.Add(xlWorksheet.Cells[j, "C"].Value());
            }

            for (int k = 2; k <= F3rngCount; k++)
            {
                focallength3.Add(xlWorksheet.Cells[k, "E"].Value());
            }

            // iterate over column B, D and F's used row count and store values to the list for Entracne puplil Dia


            for (int i = 2; i <= EPD1rngCount; i++)
            {
                EPD1.Add(xlWorksheet.Cells[i, "B"].Value());
            }

            for (int j = 2; j <= EPD2rngCount; j++)
            {
                EPD2.Add(xlWorksheet.Cells[j, "D"].Value());
            }

            for (int k = 2; k <= EPD3rngCount; k++)
            {
                EPD3.Add(xlWorksheet.Cells[k, "F"].Value());
            }


            //cleanup

            GC.Collect();

            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background

            Marshal.ReleaseComObject(xlRange);

            Marshal.ReleaseComObject(xlWorksheet);

            //close and release

            xlWorkbook.Close();

            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release

            xlApp.Quit();

            Marshal.ReleaseComObject(xlApp);


        }

        public static double perm(List<double> F1, List<double> F2, List<double> F3, List<double> EP1, List<double> EP2, List<double> EP3)
        {
            int i, j, k;

            for (i = 0; i < F1.Count; i++)
            {

                for (j = 0; j < F2.Count; j++)
                {
                    for (k = 0; k < F3.Count; k++)
                    {

                        a1[k] = Math.Round((double)F1[i] + F2[j], 4);

                        MxratioMy[k] = Math.Round((double)F1[i] / F3[k], 4);

                        a2[k] = Math.Round((double)F2[j] + F3[k], 4);

                        b1[k] = Math.Round((double)(F1[i] * F2[j]) / F3[k], 4);

                        b2[k] = Math.Round((double)(F2[j] * F3[k]) / F1[i], 4);

                        Mx[k] = Math.Round((double)-a2[k] / b2[k], 4);

                        My[k] = Math.Round((double)-b1[k] / a1[k], 4);



                        Maxtrack[k] = Math.Round((double)F1[i] + 2 * F2[j] + F3[k] + (F2[j] * (((F3[k] * MxratioMy[k]) / F1[i]) + F1[i] / (F3[k] * MxratioMy[k]))), 4);

                        templist.Add(Maxtrack[k]);

                        if ((Mx[k] > MxratioMy[k]) && (MxratioMy[k] > My[k]) && (InputMax <= Mx[k]) && (InputMin >= My[k]) && (InputMax > InputMin) && (InputMin < InputMax) && (EPD1[k] > EPDConstrain) && (EPD1[k] == EPD2[k]))
                        {
                            F1List.Add(F1[i]);

                            F2List.Add(F2[j]);

                            F3List.Add(F3[k]);


                            MaxtrackList.Add(Maxtrack[k]);

                            MxList.Add(Mx[k]);

                            MyList.Add(My[k]);

                        }

                        else

                            if ((MxratioMy[k] > Mx[k]) || (My[k] > MxratioMy[k]) || (InputMax > Mx[k]) || (InputMin < My[k]) || (InputMax < InputMin) || (EPD1[k] < EPDConstrain) || (EPD1[k] != EPD2[k]))
                            {

                                // Do nothing here just ignore the values

                            }
                    }

                }
            }

            //check for emptiness of a List for no suitable combination of focal length

            if (!MaxtrackList.Any())
            {
                Console.WriteLine("There is no suitable focal length in database for this configuration \n");

                return 0;
            }

            else

                // Display Tracklengths only once last element in the list is reached

                for (int p = 0; p < MaxtrackList.Count; p++)
                {

                    if (p == MaxtrackList.Count - 1)
                    {
                        // Get Maximum and Minimum value of Tracklength with respective Focal lengths  

                        Console.WriteLine("Maxtrackvalue = {0} with F1 = {1}, F2 = {2} and F3 = {3} \n", MaxtrackList.Max(), F1List[MaxtrackList.IndexOf(MaxtrackList.Max())], F2List[MaxtrackList.IndexOf(MaxtrackList.Max())], F3List[MaxtrackList.IndexOf(MaxtrackList.Max())]);

                        Console.WriteLine("The Maxtrackvalue can provide Max Magnifiaction = {0} and Min Magnification = {1} \n", MxList[MaxtrackList.IndexOf(MaxtrackList.Max())], MyList[MaxtrackList.IndexOf(MaxtrackList.Max())]);

                        Console.WriteLine("Mintrackvalue = {0} with F1 = {1}, F2 = {2} and F3 = {3} \n", MaxtrackList.Min(), F1List[MaxtrackList.IndexOf(MaxtrackList.Min())], F2List[MaxtrackList.IndexOf(MaxtrackList.Min())], F3List[MaxtrackList.IndexOf(MaxtrackList.Min())]);

                        Console.WriteLine("The Mintrackvalue can provide Max Magnifiaction = {0} and Min Magnification = {1} \n", MxList[MaxtrackList.IndexOf(MaxtrackList.Min())], MyList[MaxtrackList.IndexOf(MaxtrackList.Min())]);
                    }
                }

            Console.WriteLine("\n");

            return userinputs(F1, F2, F3, EP1, EP2, EP3);
        }

        public static double userinputs(List<double> F1, List<double> F2, List<double> F3, List<double> EP1, List<double> EP2, List<double> EP3)
        {
            Console.WriteLine("\n");

            Console.WriteLine("Choose Tracklength \n");

            Console.WriteLine("Press (a) for Maxtrack and (b) for Mintrack\n ");

            Console.WriteLine("\n\nRecommended option is b \n\n");


            string choose = Console.ReadLine();

            Console.WriteLine("\n");


            switch (choose)
            {
                // Choosing Maxtracklength option for focallengths

                case "a":

                    return Maxtractcal(F1, F2, F3, EP1, EP2, EP3);

                case "b":

                    return Mintrackcal(F1, F2, F3, EP1, EP2, EP3);

                default:

                    Console.WriteLine("Please choose from (a) or (b) \n");

                    break;

            }

            return userinputs(F1, F2, F3, EP1, EP2, EP3);


        }

        public static double Maxtractcal(List<double> F1, List<double> F2, List<double> F3, List<double> EP1, List<double> EP2, List<double> EP3)
        {
            double Maxd1, Maxd2, MaxInput, MaxF1, MaxF2, MaxF3, Maxa1, Maxa2, Maxb1, Maxb2, MaxMx, MaxMy, MaxMxratioMy;


            int a = 1;

            if (a == 1)
            {
                Console.WriteLine("Focallength choosed with Maxtrack = {0} are: F1 = {1}, F2 = {2}, F3 = {3} \n", MaxtrackList.Max(), F1List[MaxtrackList.IndexOf(MaxtrackList.Max())], F2List[MaxtrackList.IndexOf(MaxtrackList.Max())], F3List[MaxtrackList.IndexOf(MaxtrackList.Max())]);

                Console.WriteLine("Please choose values between or equal to Max and Min Magnification \n");

                a = a + 1;
            }

            MaxF1 = F1List[MaxtrackList.IndexOf(MaxtrackList.Max())];

            MaxF2 = F2List[MaxtrackList.IndexOf(MaxtrackList.Max())];

            MaxF3 = F3List[MaxtrackList.IndexOf(MaxtrackList.Max())];

            SelF1 = MaxF1;

            SelF2 = MaxF2;

            SelF3 = MaxF3;


            Maxa1 = Math.Round((double)MaxF1 + MaxF2, 4);

            MaxMxratioMy = Math.Round((double)MaxF1 / MaxF3, 4);

            Maxa2 = Math.Round((double)MaxF2 + MaxF3, 4);

            Maxb1 = Math.Round((double)(MaxF1 * MaxF2 / MaxF3), 4);

            Maxb2 = Math.Round((double)(MaxF2 * MaxF3) / MaxF1, 4);

            MaxMx = Math.Round((double)-Maxa2 / Maxb2, 4);

            MaxMy = Math.Round((double)-Maxb1 / Maxa1, 4);


            while (true)
            {

                Console.WriteLine("Enter Magnification upto 4 decimal point or Enter (000) to quit and see Zemax File \n");

                // Check for value other than numerics

                while (!Double.TryParse(Console.ReadLine(), out MaxInput))
                {

                    Console.WriteLine("Please enter numeric value \n");

                    Console.WriteLine("Enter Magnification upto 4 decimal point or Enter (000) to quit and see Zemax File \n");
                }


                Console.WriteLine("\n");

                if (MaxInput != 000)
                {
                    if ((MaxInput <= InputMax) && (MaxInput >= InputMin))
                    {

                        Console.WriteLine("Conditions satified \n");


                        //Calculate d1 and d2 for the Input Magnification

                        Maxd1 = Math.Round((double)MaxF1 + MaxF2 + ((MaxF1 * MaxF2) / (MaxInput * MaxF3)), 4);

                        Maxd2 = Math.Round((double)MaxF2 + MaxF3 + ((MaxF2 * MaxF3 * MaxInput) / (MaxF1)), 4);

                        if ((Maxd1 >= -0.012) && (Maxd1 < 0))
                        {
                            Maxd1 = 0;
                        }
                        else

                            if ((Maxd2 >= -0.012) && (Maxd2 < 0))
                            {
                                Maxd2 = 0;
                            }

                        Console.WriteLine("The system has d1 = {0} and d2 = {1} for the Input Magnification = {2} \n", Maxd1, Maxd2, MaxInput);

                        Temp1.Add(Maxd1);

                        Temp2.Add(Maxd2);


                    }

                    else

                        if ((MaxInput > InputMax) || (MaxInput < InputMin))
                        {


                            Console.WriteLine("Conditions didn't satified \n");

                            Console.WriteLine("Please choose values between or equal to Max and Min Magnification \n");

                            return Maxtractcal(F1, F2, F3, EP1, EP2, EP3);

                        }

                }

                else

                    if (MaxInput == 000)
                    {
                        ZemaxInitialize();
                    }



            }

            // return Maxtractcal(F1, F2, F3, MxratioMy, Maxtrack);
        }

        public static double Mintrackcal(List<double> F1, List<double> F2, List<double> F3, List<double> EP1, List<double> EP2, List<double> EP3)
        {
            double MinInput, MinF1, MinF2, MinF3, Mina1, Mina2, Minb1, Minb2, MinMx, MinMy, MinMxratioMy;

            int a = 1;

            if (a == 1)
            {

                Console.WriteLine("Focallength choosed with Mintrack = {0} are: F1 = {1}, F2 = {2}, F3 = {3} \n", MaxtrackList.Min(), F1List[MaxtrackList.IndexOf(MaxtrackList.Min())], F2List[MaxtrackList.IndexOf(MaxtrackList.Min())], F3List[MaxtrackList.IndexOf(MaxtrackList.Min())]);

                Console.WriteLine("Please choose values between or equal to InputMax and InputMin Magnification \n");

                a = a + 1;

            }

            MinF1 = F1List[MaxtrackList.IndexOf(MaxtrackList.Min())];

            MinF2 = F2List[MaxtrackList.IndexOf(MaxtrackList.Min())];

            MinF3 = F3List[MaxtrackList.IndexOf(MaxtrackList.Min())];

            SelF1 = MinF1;

            SelF2 = MinF2;

            SelF3 = MinF3;


            Mina1 = Math.Round((double)MinF1 + MinF2, 4);

            MinMxratioMy = Math.Round((double)MinF1 / MinF3, 4);

            Mina2 = Math.Round((double)MinF2 + MinF3, 4);

            Minb1 = Math.Round((double)(MinF1 * MinF2 / MinF3), 4);

            Minb2 = Math.Round((double)(MinF2 * MinF3) / MinF1, 4);

            MinMx = Math.Round((double)-Mina2 / Minb2, 4);

            MinMy = Math.Round((double)-Minb1 / Mina1, 4);


            while (true)
            {

                Console.WriteLine("Enter Magnification upto 4 decimal point or Enter (000) to quit and see Zemax File \n");

                // Check for value other than numerics

                while (!Double.TryParse(Console.ReadLine(), out MinInput))
                {
                    Console.WriteLine("Please enter numeric value \n");

                    Console.WriteLine("Enter Magnification upto 4 decimal point or Enter (000) to quit and see Zemax File \n");
                }

                if (MinInput != 000)
                {
                    if ((MinInput <= InputMax) && (MinInput >= InputMin))
                    {

                        Console.WriteLine("Conditions satified \n");


                        //Calculate d1 and d2 for the Input Magnification

                        Mind1 = Math.Round((double)MinF1 + MinF2 + ((MinF1 * MinF2) / (MinInput * MinF3)), 4);

                        Mind2 = Math.Round((double)MinF2 + MinF3 + ((MinF2 * MinF3 * MinInput) / (MinF1)), 4);

                        if ((Mind1 >= -0.012) && (Mind1 < 0))
                        {
                            Mind1 = 0;
                        }
                        else

                            if ((Mind2 >= -0.012) && (Mind2 < 0))
                            {
                                Mind2 = 0;
                            }

                        Console.WriteLine("The system has d1 = {0} and d2 = {1} for the Input Magnification = {2} \n", Mind1, Mind2, MinInput);

                        Temp1.Add(Mind1);

                        Temp2.Add(Mind2);

                    }

                    else

                        if ((MinInput > InputMax) || (MinInput < InputMin))
                        {


                            Console.WriteLine("Conditions didn't satified \n");

                            Console.WriteLine("Please choose values between or equal to InputMax and InputMin Magnification \n");

                            return Mintrackcal(F1, F2, F3, EP1, EP2, EP3);

                        }

                    //for (int s = 0; s < 3; s++)
                    //{
                    //    Temp1.Add(Mind1);

                    //    Temp2.Add(Mind2);

                    //}


                }

                else

                    if (MinInput == 000)
                    {
                        // See the output of the stored distances

                        //for (int mn = 0; mn < Temp1.Count; mn++)
                        //{
                        //    Console.WriteLine("D1 = {0} \n", Temp1[mn]); ;

                        //    Console.WriteLine("D2 = {0} \n", Temp2[mn]); ;

                        //}


                        ZemaxInitialize();
                    }

            }

            // return Mintrackcal(F1, F2, F3, MxratioMy, Maxtrack);
        }

        public static void ZemaxInitialize()
        {
            // Find the installed version of OpticStudio
            bool isInitialized = ZOSAPI_NetHelper.ZOSAPI_Initializer.Initialize();
            // Note -- uncomment the following line to use a custom initialization path
            //bool isInitialized = ZOSAPI_NetHelper.ZOSAPI_Initializer.Initialize(@"C:\Program Files\OpticStudio\");
            if (isInitialized)
            {
                LogInfo("Found OpticStudio at: " + ZOSAPI_NetHelper.ZOSAPI_Initializer.GetZemaxDirectory());
            }
            else
            {
                HandleError("Failed to locate OpticStudio!");
                return;
            }

            BeginStandaloneApplication();

        }

        static void BeginStandaloneApplication()
        {
            // Create the initial connection class
            ZOSAPI_Connection TheConnection = new ZOSAPI_Connection();

            // Attempt to create a Standalone connection
            IZOSAPI_Application TheApplication = TheConnection.CreateNewApplication();
            if (TheApplication == null)
            {
                HandleError("An unknown connection error occurred!");
                return;
            }

            // Check the connection status
            if (!TheApplication.IsValidLicenseForAPI)
            {
                HandleError("Failed to connect to OpticStudio: " + TheApplication.LicenseStatus);
                return;
            }
            if (TheApplication.Mode != ZOSAPI_Mode.Server)
            {
                HandleError("User plugin was started in the wrong mode: expected Server, found " + TheApplication.Mode.ToString());
                return;
            }

            IOpticalSystem TheSystem = TheApplication.PrimarySystem;

            // Add your custom code here...

            // creates new directory
            string strPath = System.IO.Path.Combine(TheApplication.SamplesDir, @"API\CS#");
            System.IO.Directory.CreateDirectory(strPath);

            TheSystem.LoadFile(TheApplication.SamplesDir + "\\Sequential\\Objectives\\BX.zmx", false);

            TheSystem.New(false);

            // Open MCE and MFE

            TheSystem.MCE.ShowMCE();

            TheSystem.MFE.ShowMFE();



            //! [e19s01_cs]
            // ISystemData represents the System Explorer in GUI.
            // We access options in System Explorer through ISystemData in ZOS-API
            ISystemData TheSystemData = TheSystem.SystemData;
            TheSystemData.Aperture.ApertureValue = InputbeamDia;
            TheSystemData.Aperture.SemiDiameterMargin = 2;
            TheSystemData.Aperture.AFocalImageSpace = true;
            TheSystemData.Wavelengths.GetWavelength(1).Wavelength = 0.55;
            //! [e19s01_cs]

            // Get interface of Lens Data Editor and add 3 surfaces.
            //------------------------------------
            ILensDataEditor TheLDE = TheSystem.LDE;
            TheLDE.InsertNewSurfaceAt(2);
            TheLDE.InsertNewSurfaceAt(3);
            TheLDE.InsertNewSurfaceAt(4);
            //-----------------------------------



            //! [e18s06_cs]
            // Refocus for each configuration
            //------------------------------------
            //IQuickFocus quickfocus = TheSystem.Tools.OpenQuickFocus();
            //TheMCE.SetCurrentConfiguration(1);
            //quickfocus.RunAndWaitForCompletion();
            //TheMCE.SetCurrentConfiguration(2);
            //quickfocus.RunAndWaitForCompletion();
            //TheMCE.SetCurrentConfiguration(3);
            //quickfocus.RunAndWaitForCompletion();
            //------------------------------------
            //! [e18s06_cs]






            //-------------------------------
            //TheLDE.GetSurfaceAt(4).Thickness = 30;
            //TheLDE.GetSurfaceAt(2).Material = "N-BK7";
            //-------------------------------
            //! [e19s02_cs]

            //! [e19s03_cs]
            // GetSurfaceAt(surface number shown in LDE) will return an interface ILDERow
            // Through property TiltDecenterData of each interface ILDERow, we can modify data in Surface Properties > Tilt/Decenter section
            //-------------------------------------------------------
            //TheLDE.GetSurfaceAt(2).TiltDecenterData.BeforeSurfaceOrder = TiltDecenterOrderType.Decenter_Tilt;
            //TheLDE.GetSurfaceAt(2).TiltDecenterData.BeforeSurfaceTiltX = 15;
            //TheLDE.GetSurfaceAt(2).TiltDecenterData.AfterSurfaceTiltX = -15;
            //TheLDE.GetSurfaceAt(3).TiltDecenterData.BeforeSurfaceTiltX = -15;
            //TheLDE.GetSurfaceAt(3).TiltDecenterData.AfterSurfaceTiltX = 15;
            //----------------------------------------------------------
            //! [e19s03_cs]

            //! [e19s04_cs]
            // To specify an aperture to a surface, we need to first create an ISurfaceApertureType and then assign it.
            //-------------------------------------
            //ISurfaceApertureType Rect_Aper = TheLDE.GetSurfaceAt(2).ApertureData.CreateApertureTypeSettings(SurfaceApertureTypes.RectangularAperture);
            //Rect_Aper._S_RectangularAperture.XHalfWidth = 10;
            //Rect_Aper._S_RectangularAperture.YHalfWidth = 10;
            //TheLDE.GetSurfaceAt(2).ApertureData.ChangeApertureTypeSettings(Rect_Aper);
            //TheLDE.GetSurfaceAt(3).ApertureData.PickupFrom = 2;
            //-----------------------------------------
            //! [e19s04_cs]

            //! [e19s05_cs]
            // To change surface type, we need to first get an ISurfaceTypesettings and then assign it.
            //----------------------------
            //ISurfaceTypeSettings SurfaceType_CB = TheLDE.GetSurfaceAt(4).GetSurfaceTypeSettings(SurfaceType.CoordinateBreak);
            //TheLDE.GetSurfaceAt(4).ChangeType(SurfaceType_CB);

            ISurfaceTypeSettings SurfaceType_Paraxial = TheLDE.GetSurfaceAt(2).GetSurfaceTypeSettings(SurfaceType.Paraxial);
            TheLDE.GetSurfaceAt(2).ChangeType(SurfaceType_Paraxial);
            TheLDE.GetSurfaceAt(3).ChangeType(SurfaceType_Paraxial);
            TheLDE.GetSurfaceAt(4).ChangeType(SurfaceType_Paraxial);


            // Set thickness and material for each surface.

            //List<double> TheLDE.GetSurfaceAt(2).Thickness = new List<double>();
            //List<double> ThicknessAt2 = new List<double>();

            List<double> T1 = new List<double>();
            List<double> T2 = new List<double>();

            TheLDE.GetSurfaceAt(1).Thickness = 10;

            TheLDE.GetSurfaceAt(4).Thickness = 10;


            T1.Add(TheLDE.GetSurfaceAt(2).Thickness);
            T1 = Temp1;

            T2.Add(TheLDE.GetSurfaceAt(3).Thickness);
            T2 = Temp2;

            // Get interface of the Multi-Configuration Editor
            IMultiConfigEditor TheMCE = TheSystem.MCE;
            //! [e18s01_cs]
            // Add two configurations (totally 3)
            TheMCE.AddConfiguration(false);
            TheMCE.AddConfiguration(false);

            //! [e18s01_cs]

            //! [e18s02_cs]
            // Add one operand (totally 2)
            TheMCE.AddOperand();
            //! [e18s02_cs]

            //! [e18s03_cs]
            // Get interface of each operand
            IMCERow MCOperand1 = TheMCE.GetOperandAt(1);
            IMCERow MCOperand2 = TheMCE.GetOperandAt(2);
            // Change both operands' type to THIC
            MCOperand1.ChangeType(MultiConfigOperandType.THIC);
            MCOperand2.ChangeType(MultiConfigOperandType.THIC);
            //! [e18s03_cs]

            //! [e18s04_cs]
            // Set parameters of operands
            // If the type of operand is THIC, the first parameter here means surface number
            MCOperand1.Param1 = 2;
            MCOperand2.Param1 = 3;
            //! [e18s04_cs]

            //! [e18s05_cs]
            // Set values of opeand for each configurations


            //MCOperand1.GetOperandCell(3).DoubleValue = T1[g];
            //MCOperand2.GetOperandCell(3).DoubleValue = T2[g];

            //MCOperand1.GetOperandCell(2).DoubleValue = T1[g];
            //MCOperand2.GetOperandCell(2).DoubleValue = T2[g];


            //! [e18s05_cs]


            //}

            for (int w = 0; w < Temp1.Count; w++)
            {
                MCOperand1.GetOperandCell(3).DoubleValue = T1[0];
                MCOperand2.GetOperandCell(3).DoubleValue = T2[0];

                MCOperand1.GetOperandCell(2).DoubleValue = T1[1];
                MCOperand2.GetOperandCell(2).DoubleValue = T2[1];

                MCOperand1.GetOperandCell(1).DoubleValue = T1[2];
                MCOperand2.GetOperandCell(1).DoubleValue = T2[2];



            }

            TheLDE.GetSurfaceAt(4).Thickness = 10;


            //! [e19s02_cs]

            //-----------------------------------
            //! [e19s05_cs]

            //! [e19s06_cs]
            // Set Chief Ray solves to surface 4, which is Coordinate Break
            // To set a solve to a cell in editor, we need to first create a ISolveData and then assign it.
            //--------------------------------------------
            //ISolveData Solve_ChiefNormal = TheLDE.GetSurfaceAt(4).GetSurfaceCell(SurfaceColumn.Par1).CreateSolveType(SolveType.PickupChiefRay);
            //TheLDE.GetSurfaceAt(4).GetSurfaceCell(SurfaceColumn.Par1).SetSolveData(Solve_ChiefNormal);
            //TheLDE.GetSurfaceAt(4).GetSurfaceCell(SurfaceColumn.Par2).SetSolveData(Solve_ChiefNormal);
            //TheLDE.GetSurfaceAt(4).GetSurfaceCell(SurfaceColumn.Par3).SetSolveData(Solve_ChiefNormal);
            //TheLDE.GetSurfaceAt(4).GetSurfaceCell(SurfaceColumn.Par4).SetSolveData(Solve_ChiefNormal);
            //TheLDE.GetSurfaceAt(4).GetSurfaceCell(SurfaceColumn.Par5).SetSolveData(Solve_ChiefNormal);
            //-----------------------------------------------------
            //! [e19s06_cs]

            //! [e19s07_cs]
            // Copy 3 surfaces starting from surface number 2 in LDE and paste to surface number 5, 
            // which will become surface number 8 after pasting.
            //-------------------------------
            //for (int i = 0; i < 10; i++)
            //{
            //    TheLDE.CopySurfaces(2, 3, 5);
            //}
            //-------------------------------

            //Merit Funtions

            //--------------------------------------------

            // Operands for !st Configuration

            IMeritFunctionEditor TheMFE1 = TheSystem.MFE;

            IMFERow Operand_1 = TheMFE1.GetOperandAt(1);

            Operand_1.ChangeType(MeritOperandType.CONF);

            Operand_1.GetCellAt(2).IntegerValue = 1;

            IMFERow Operand_2 = TheMFE1.InsertNewOperandAt(2);

            Operand_2.ChangeType(MeritOperandType.REAY);

            Operand_2.GetCellAt(2).IntegerValue = 1;

            Operand_2.GetCellAt(3).IntegerValue = 1;

            Operand_2.GetCellAt(7).DoubleValue = 1;

            IMFERow Operand_3 = TheMFE1.InsertNewOperandAt(3);

            Operand_3.ChangeType(MeritOperandType.REAY);

            Operand_3.GetCellAt(2).IntegerValue = 5;

            Operand_3.GetCellAt(3).IntegerValue = 1;

            Operand_3.GetCellAt(7).DoubleValue = 1;

            IMFERow Operand_4 = TheMFE1.InsertNewOperandAt(4);

            Operand_4.ChangeType(MeritOperandType.DIVI);

            Operand_4.GetCellAt(2).IntegerValue = 2;

            Operand_4.GetCellAt(3).IntegerValue = 3;

            //-------------------------------------------

            //Operand for 2nd Configuration

            IMeritFunctionEditor TheMFE2 = TheSystem.MFE;

            IMFERow Operand_5 = TheMFE2.InsertNewOperandAt(5);

            Operand_5.ChangeType(MeritOperandType.CONF);

            Operand_5.GetCellAt(2).IntegerValue = 2;

            IMFERow Operand_6 = TheMFE2.InsertNewOperandAt(6);

            Operand_6.ChangeType(MeritOperandType.REAY);

            Operand_6.GetCellAt(2).IntegerValue = 1;

            Operand_6.GetCellAt(3).IntegerValue = 1;

            Operand_6.GetCellAt(7).DoubleValue = 1;

            IMFERow Operand_7 = TheMFE2.InsertNewOperandAt(7);

            Operand_7.ChangeType(MeritOperandType.REAY);

            Operand_7.GetCellAt(2).IntegerValue = 5;

            Operand_7.GetCellAt(3).IntegerValue = 1;

            Operand_7.GetCellAt(7).DoubleValue = 1;

            IMFERow Operand_8 = TheMFE2.InsertNewOperandAt(8);

            Operand_8.ChangeType(MeritOperandType.DIVI);

            Operand_8.GetCellAt(2).IntegerValue = 6;

            Operand_8.GetCellAt(3).IntegerValue = 7;

            //---------------------------------------------------

            //Operand for 3rd Configuration

            IMeritFunctionEditor TheMFE3 = TheSystem.MFE;

            IMFERow Operand_9 = TheMFE3.InsertNewOperandAt(9);

            Operand_9.ChangeType(MeritOperandType.CONF);

            Operand_9.GetCellAt(2).IntegerValue = 3;

            IMFERow Operand_10 = TheMFE3.InsertNewOperandAt(10);

            Operand_10.ChangeType(MeritOperandType.REAY);

            Operand_10.GetCellAt(2).IntegerValue = 1;

            Operand_10.GetCellAt(3).IntegerValue = 1;

            Operand_10.GetCellAt(7).DoubleValue = 1;

            IMFERow Operand_11 = TheMFE3.InsertNewOperandAt(11);

            Operand_11.ChangeType(MeritOperandType.REAY);

            Operand_11.GetCellAt(2).IntegerValue = 5;

            Operand_11.GetCellAt(3).IntegerValue = 1;

            Operand_11.GetCellAt(7).DoubleValue = 1;

            IMFERow Operand_12 = TheMFE3.InsertNewOperandAt(12);

            Operand_12.ChangeType(MeritOperandType.DIVI);

            Operand_12.GetCellAt(2).IntegerValue = 10;

            Operand_12.GetCellAt(3).IntegerValue = 11;

            //--------------------------------------------

            // Open MCE and MFE

            //TheSystem.MCE.ShowMCE();

            //TheSystem.MFE.ShowMFE();

            //------------------------------------------

            //setup choosen focal lengths (F1, F2 and F3)

            ILDERow Surface2 = TheLDE.GetSurfaceAt(2);

            Surface2.GetCellAt(12).DoubleValue = SelF1;

            ILDERow Surface3 = TheLDE.GetSurfaceAt(3);

            Surface3.GetCellAt(12).DoubleValue = SelF2;

            ILDERow Surface4 = TheLDE.GetSurfaceAt(4);

            Surface4.GetCellAt(12).DoubleValue = SelF3;

            // Save file
            TheSystem.SaveAs(TheApplication.SamplesDir + "\\API\\CS#\\BX.ZMX");
            //! [e19s07_cs]

            // TheSystem.LoadFile(TheApplication.SamplesDir + "\\Sequential\\Objectives\\BX.zmx");




            //! [e19s08_cs]
            // Run tool Convert Local To Global Coordinates to convert surface #2 to surface #35 to be globally referenced to surface #1
            //------------------------------------------
            //TheLDE.RunTool_ConvertLocalToGlobalCoordinates(2, 35, 1);
            //TheSystem.SaveAs(TheApplication.SamplesDir + "\\API\\CS#\\e19_Sample_Prism_Chain_GlobalCoordinate.ZMX");
            //------------------------------------------
            //! [e19s08_cs]

            //! [e19s09_cs]
            // Run tool Conver Global To Local Coordinates to convert surface #1 to surface #57 back to local coordinate.
            //------------------------------------------
            //TheLDE.RunTool_ConvertGlobalToLocalCoordinates(1, 57, 0);
            //TheSystem.SaveAs(TheApplication.SamplesDir + "\\API\\CS#\\e19_Sample_Prism_Chain_BackTo_LocalCoordinate.ZMX");
            //------------------------------------------
            //! [e19s09_cs]

            Console.Write("Press any key to continue...");
            Console.ReadKey();

            // Clean up
            FinishStandaloneApplication(TheApplication);
        }

        static void FinishStandaloneApplication(IZOSAPI_Application TheApplication)
        {
            // Note - TheApplication will close automatically when this application exits, so this isn't strictly necessary in most cases
            if (TheApplication != null)
            {
                TheApplication.CloseApplication();
            }
        }

        static void LogInfo(string message)
        {
            // TODO - add custom logging
            Console.WriteLine(message);
        }

        static void HandleError(string errorMessage)
        {
            // TODO - add custom error handling
            throw new Exception(errorMessage);
        }

    }
}