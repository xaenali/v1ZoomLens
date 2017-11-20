using System;
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
using ZOSAPI.Tools.Optimization;

namespace BXwithZemax
{
    
    // InputMax = user input value for Max magnification input
    // InputMin = user input value for Min magnification input
    // Mx = Max magnification possible
    // My = Min magnification possible

    //Note: where ever Mx, My and MxratioMy is used with different notation then it is for the same calculation as before 

    //foallength 1,2,3 = imported focallentghs
    //MxratioMy = ratiobetween focallength1 and focallength3
    //EPD1,2,3 = Entrance pupil diameter for lens 1, 2 and 3   
     
    class Program
    {
        public static double InputMax, InputMin, InputbeamDia, EPDConstrainF1, EPDConstrainF3, Mind1, Mind2, SelF1, SelF2, SelF3;
        public static List<double> Temp1 = new List<double>();
        public static List<double> Temp2 = new List<double>();
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
        public static List<double> EPD1List = new List<double>(); // Initialize List for Entrance pupil Dia 1
        public static List<double> EPD2List = new List<double>(); // Initialize List for Entrance pupil Dia  2
        public static List<double> EPD3List = new List<double>(); // Initialize List for fEntrance pupil Dia  3


        static void Main(string[] args)
        {

            Console.WriteLine("Accessing Focal length database \n");

            getExcelFile();

            Console.WriteLine("Access Complete \n");

            while (true)
            {

                //n = 0; // Make n(number of combination) = 0, in start and at the end of all the combination 

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

                EPDConstrainF1 = (double)1.5 * InputbeamDia;

                EPDConstrainF3 = (double)(1 / InputMin) * 1.5 * InputbeamDia;

                Console.WriteLine("\n");

                Console.WriteLine("Enter 0 or 1 to see all combinations with each resutls or d1 and d2 for Max Mag respectively \n");

                string input = Console.ReadLine();

                Console.WriteLine("\n");

                switch (input)
                {
                    case "0":

                        Allperm();

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

        public static double Allperm()
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

            int n = 0;


            //take f1, f2 and f3 to calculate Max magnification and compare with Mx and My magnifications

            for (int k = 0; k < focallength1.Count; k++) // take f1
            {

                for (int l = 0; l < focallength2.Count; l++) // take f2
                {


                    for (int m = 0; m < focallength3.Count; m++) // take f3
                    {

                        a1[m] = Math.Round((double)focallength1[k] + focallength2[l], 4);

                        MxratioMy[m] = Math.Round((double)focallength1[k] / focallength3[m], 4);

                        a2[m] = Math.Round((double)focallength2[l] + focallength3[m], 4);

                        b1[m] = Math.Round((double)(focallength1[k] * focallength2[l]) / focallength3[m], 4);

                        b2[m] = Math.Round((double)(focallength2[l] * focallength3[m]) / focallength1[k], 4);

                        Mx[m] = Math.Round((double)-a2[m] / b2[m], 4);

                        My[m] = Math.Round((double)-b1[m] / a1[m], 4);


                        //comparison with Mx and My magnifications

                        if (EPD1[k] >= EPDConstrainF1)
                        {
                            if (EPD3[m] >= EPDConstrainF3 && EPD2[l] != EPD3[m])
                            {
                                if ((Mx[m] > MxratioMy[m]) && (MxratioMy[m] > My[m]) && (InputMax <= Mx[m]) && (InputMin >= My[m]) && (InputMax > InputMin))
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

                            }
                        }



                        else

                            if (EPD1[k] < EPDConstrainF1)
                            {
                                if(EPD3[m] < EPDConstrainF3 || EPD2[l] == EPD3[m])
                                {
                                    if ((MxratioMy[m] > Mx[m]) || (My[m] > MxratioMy[m]) || (InputMax > Mx[m]) || (InputMin < My[m]) || (InputMax < InputMin))
                                    {
                                        n = n + 1;

                                        //Console.WriteLine("number of combination = {0} ", n);

                                        //Console.WriteLine("Conditions didn't satified for comination {0}", n);

                                        //Console.WriteLine("Can't choose InputMax = {0} and InputMin = {1} as InputMax ({0}) > Calculated Mx {2} or InputMin ({1}) < calculated My {3} with F1 as {4}, F2 as {5} and F3 as {6} ", InputMax, InputMin, Mx[m], My[m], focallength1[k], focallength2[l], focallength3[m]);

                                    }

                                }
                            }


                    }

                }

            }



            return 0;
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

                        if (EPD1[i] >= EPDConstrainF1)
                        {
                            if (EPD3[k] >= EPDConstrainF3 && EPD2[j] != EPD3[k])
                            {
                                if ((Mx[k] > MxratioMy[k]) && (MxratioMy[k] > My[k]) && (InputMax <= Mx[k]) && (InputMin >= My[k]) && (InputMax > InputMin) && (InputMin < InputMax))
                                {
                                    F1List.Add(F1[i]);

                                    F2List.Add(F2[j]);

                                    F3List.Add(F3[k]);

                                    EPD1List.Add(EPD1[i]);

                                    EPD2List.Add(EPD2[j]);

                                    EPD3List.Add(EPD3[k]);


                                    MaxtrackList.Add(Maxtrack[k]);

                                    MxList.Add(Mx[k]);

                                    MyList.Add(My[k]);

                                }

                            }
                        }


                        else

                            if (EPD1[i] < EPDConstrainF1)
                            {
                                if (EPD3[k] < EPDConstrainF3 || EPD2[j] == EPD3[k])
                                {
                                    if ((MxratioMy[k] > Mx[k]) || (My[k] > MxratioMy[k]) || (InputMax > Mx[k]) || (InputMin < My[k]) || (InputMax < InputMin))
                                    {

                                        // Do nothing here just ignore the values

                                    }

                                }
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
                        int EPD = MaxtrackList.Count - 1;

                        // Get Maximum and Minimum value of Tracklength with respective Focal lengths  

                        Console.WriteLine("Maxtrackvalue = {0} with F1 = {1}, F2 = {2} and F3 = {3} \n", MaxtrackList.Max(), F1List[MaxtrackList.IndexOf(MaxtrackList.Max())], F2List[MaxtrackList.IndexOf(MaxtrackList.Max())], F3List[MaxtrackList.IndexOf(MaxtrackList.Max())]);

                        Console.WriteLine("The Maxtrackvalue can provide Max Magnifiaction = {0} and Min Magnification = {1} \n", MxList[MaxtrackList.IndexOf(MaxtrackList.Max())], MyList[MaxtrackList.IndexOf(MaxtrackList.Max())]);

                        Console.WriteLine("Mintrackvalue = {0} with F1 = {1}, F2 = {2} and F3 = {3} and with EPD1 = {4}, EPD2= {5}, EPD3 = {6} \n", MaxtrackList.Min(), F1List[MaxtrackList.IndexOf(MaxtrackList.Min())], F2List[MaxtrackList.IndexOf(MaxtrackList.Min())], F3List[MaxtrackList.IndexOf(MaxtrackList.Min())], EPD1List[MaxtrackList.IndexOf(MaxtrackList.Min())], EPD2List[MaxtrackList.IndexOf(MaxtrackList.Min())], EPD3List[MaxtrackList.IndexOf(MaxtrackList.Min())]);

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

            //ILensCatalogs Cataloglenses = TheSystem.Tools.OpenLensCatalogs();

            //Cataloglenses.RunAndWaitForCompletion();

           

         //   MaterialCatalogs.AddCatalog(string) = thor;

            ILensDataEditor TheLDE = TheSystem.LDE;
            TheLDE.InsertNewSurfaceAt(2);
            TheLDE.InsertNewSurfaceAt(3);
            TheLDE.InsertNewSurfaceAt(4);
            TheLDE.InsertNewSurfaceAt(5);
            TheLDE.InsertNewSurfaceAt(6);
            TheLDE.InsertNewSurfaceAt(7);

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
            TheLDE.GetSurfaceAt(5).ChangeType(SurfaceType_Paraxial);
            TheLDE.GetSurfaceAt(6).ChangeType(SurfaceType_Paraxial);
            TheLDE.GetSurfaceAt(7).ChangeType(SurfaceType_Paraxial);



            // Set thickness and material for each surface.

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


            // Setup Variables

            ISolveData configvariable = TheMCE.GetOperandAt(1).GetOperandCell(1).CreateSolveType(ZOSAPI.Editors.SolveType.Variable);

            


            // Set values of opeand for each configurations

            for (int w = 0; w < Temp1.Count; w++)
            {
                MCOperand1.GetOperandCell(3).DoubleValue = T1[0];
                MCOperand2.GetOperandCell(3).DoubleValue = T2[0];
                TheMCE.GetOperandAt(1).GetOperandCell(3).SetSolveData(configvariable);
                TheMCE.GetOperandAt(2).GetOperandCell(3).SetSolveData(configvariable);


               
                MCOperand1.GetOperandCell(2).DoubleValue = T1[1];
                MCOperand2.GetOperandCell(2).DoubleValue = T2[1];
                TheMCE.GetOperandAt(1).GetOperandCell(2).SetSolveData(configvariable);
                TheMCE.GetOperandAt(2).GetOperandCell(2).SetSolveData(configvariable);


                MCOperand1.GetOperandCell(1).DoubleValue = T1[2];
                MCOperand2.GetOperandCell(1).DoubleValue = T2[2];
                TheMCE.GetOperandAt(1).GetOperandCell(1).SetSolveData(configvariable);
                TheMCE.GetOperandAt(2).GetOperandCell(1).SetSolveData(configvariable);

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

            // Operands for 1st Configuration

            // CONF 1


            IMeritFunctionEditor TheMFE = TheSystem.MFE;

            IMFERow Operand_CONF1 = TheMFE.GetOperandAt(1);

            Operand_CONF1.ChangeType(MeritOperandType.CONF);

            Operand_CONF1.GetCellAt(2).IntegerValue = 1;

            IMFERow Operand_REAYop2_CONF1 = TheMFE.InsertNewOperandAt(2);

            Operand_REAYop2_CONF1.ChangeType(MeritOperandType.REAY);

            Operand_REAYop2_CONF1.GetCellAt(2).IntegerValue = 1;

            Operand_REAYop2_CONF1.GetCellAt(7).DoubleValue = 1;

            IMFERow Operand_REAYop3_CONF1 = TheMFE.InsertNewOperandAt(3);

            Operand_REAYop3_CONF1.ChangeType(MeritOperandType.REAY);

            Operand_REAYop3_CONF1.GetCellAt(2).IntegerValue = 5;

            Operand_REAYop3_CONF1.GetCellAt(7).DoubleValue = 1;

            IMFERow Operand_RANGop4_CONF1 = TheMFE.InsertNewOperandAt(4);

            Operand_RANGop4_CONF1.ChangeType(MeritOperandType.RANG);

            Operand_RANGop4_CONF1.GetCellAt(2).IntegerValue = 4;

            Operand_RANGop4_CONF1.GetCellAt(7).DoubleValue = 1;

            Operand_RANGop4_CONF1.Target = 0;

            Operand_RANGop4_CONF1.Weight = 10;

            Operand_RANGop4_CONF1.GetCellAt(7).DoubleValue = 1;

            IMFERow Operand_EFLXop5_CONF1 = TheMFE.InsertNewOperandAt(5); // EFLX for operand 5

            Operand_EFLXop5_CONF1.ChangeType(MeritOperandType.EFLX);

            Operand_EFLXop5_CONF1.GetCellAt(2).IntegerValue = 2;

            Operand_EFLXop5_CONF1.GetCellAt(3).IntegerValue = 3;

            Operand_EFLXop5_CONF1.Target = Operand_EFLXop5_CONF1.GetCellAt(12).DoubleValue;

            Operand_EFLXop5_CONF1.Weight = 1;

            IMFERow Operand_EFLXop6_CONF1 = TheMFE.InsertNewOperandAt(6);

            Operand_EFLXop6_CONF1.ChangeType(MeritOperandType.EFLX);

            Operand_EFLXop6_CONF1.GetCellAt(2).IntegerValue = 4;

            Operand_EFLXop6_CONF1.GetCellAt(3).IntegerValue = 5;

            Operand_EFLXop6_CONF1.Target = Operand_EFLXop6_CONF1.GetCellAt(12).DoubleValue;

            Operand_EFLXop6_CONF1.Weight = 1;

            IMFERow Operand_EFLXop7_CONF1 = TheMFE.InsertNewOperandAt(7);

            Operand_EFLXop7_CONF1.ChangeType(MeritOperandType.EFLX);

            Operand_EFLXop7_CONF1.GetCellAt(2).IntegerValue = 6;

            Operand_EFLXop7_CONF1.GetCellAt(3).IntegerValue = 7;

            Operand_EFLXop7_CONF1.Target = Operand_EFLXop7_CONF1.GetCellAt(12).DoubleValue;

            Operand_EFLXop7_CONF1.Weight = 1;

            IMFERow Operand_CTGTop8_CONF1 = TheMFE.InsertNewOperandAt(8);

            Operand_CTGTop8_CONF1.ChangeType(MeritOperandType.CTGT);

            Operand_CTGTop8_CONF1.GetCellAt(2).IntegerValue = 3;

            Operand_CTGTop8_CONF1.Target = 0.1;

            Operand_CTGTop8_CONF1.Weight = 1;

            IMFERow Operand_CTGTop9_CONF1 = TheMFE.InsertNewOperandAt(9);

            Operand_CTGTop9_CONF1.ChangeType(MeritOperandType.CTGT);

            Operand_CTGTop9_CONF1.GetCellAt(2).IntegerValue = 3;

            Operand_CTGTop9_CONF1.Target = 0.1;

            Operand_CTGTop9_CONF1.Weight = 1;

            IMFERow Operand_DIVIop23_CONF1 = TheMFE.InsertNewOperandAt(10);

            Operand_DIVIop23_CONF1.ChangeType(MeritOperandType.DIVI);

            Operand_DIVIop23_CONF1.GetCellAt(2).IntegerValue = 2;

            Operand_DIVIop23_CONF1.GetCellAt(3).IntegerValue = 3;



            //-------------------------------------------

            //Operand for 2nd Configuration

            // CONF 2

            IMFERow Operand_CONF2 = TheMFE.InsertNewOperandAt(11);

            Operand_CONF2.ChangeType(MeritOperandType.CONF);

            Operand_CONF2.GetCellAt(2).IntegerValue = 2;

            IMFERow Operand_REAYop12_CONF2 = TheMFE.InsertNewOperandAt(12);

            Operand_REAYop12_CONF2.ChangeType(MeritOperandType.REAY);

            Operand_REAYop12_CONF2.GetCellAt(2).IntegerValue = 1;

            Operand_REAYop12_CONF2.GetCellAt(7).DoubleValue = 1;

            IMFERow Operand_REAYop13_CONF2 = TheMFE.InsertNewOperandAt(13);

            Operand_REAYop13_CONF2.ChangeType(MeritOperandType.REAY);

            Operand_REAYop13_CONF2.GetCellAt(2).IntegerValue = 5;

            Operand_REAYop13_CONF2.GetCellAt(7).DoubleValue = 1;

            IMFERow Operand_RANGop14_CONF2 = TheMFE.InsertNewOperandAt(14);

            Operand_RANGop14_CONF2.ChangeType(MeritOperandType.RANG);

            Operand_RANGop14_CONF2.GetCellAt(2).IntegerValue = 4;

            Operand_RANGop14_CONF2.GetCellAt(7).DoubleValue = 1;

            Operand_RANGop14_CONF2.Target = 0;

            Operand_RANGop14_CONF2.Weight = 10;

            Operand_RANGop14_CONF2.GetCellAt(7).DoubleValue = 1;

            IMFERow Operand_EFLXop15_CONF2 = TheMFE.InsertNewOperandAt(15); // EFLX for operand 5

            Operand_EFLXop15_CONF2.ChangeType(MeritOperandType.EFLX);

            Operand_EFLXop15_CONF2.GetCellAt(2).IntegerValue = 2;

            Operand_EFLXop15_CONF2.GetCellAt(3).IntegerValue = 3;

            Operand_EFLXop15_CONF2.Target = Operand_EFLXop15_CONF2.GetCellAt(12).DoubleValue;

            Operand_EFLXop15_CONF2.Weight = 1;

            IMFERow Operand_EFLXop16_CONF2 = TheMFE.InsertNewOperandAt(16);

            Operand_EFLXop16_CONF2.ChangeType(MeritOperandType.EFLX);

            Operand_EFLXop16_CONF2.GetCellAt(2).IntegerValue = 4;

            Operand_EFLXop16_CONF2.GetCellAt(3).IntegerValue = 5;

            Operand_EFLXop16_CONF2.Target = Operand_EFLXop16_CONF2.GetCellAt(12).DoubleValue;

            Operand_EFLXop16_CONF2.Weight = 1;

            IMFERow Operand_EFLXop17_CONF2 = TheMFE.InsertNewOperandAt(17);

            Operand_EFLXop17_CONF2.ChangeType(MeritOperandType.EFLX);

            Operand_EFLXop17_CONF2.GetCellAt(2).IntegerValue = 6;

            Operand_EFLXop17_CONF2.GetCellAt(3).IntegerValue = 7;

            Operand_EFLXop17_CONF2.Target = Operand_EFLXop17_CONF2.GetCellAt(12).DoubleValue;

            Operand_EFLXop17_CONF2.Weight = 1;

            IMFERow Operand_CTGTop18_CONF2 = TheMFE.InsertNewOperandAt(18);

            Operand_CTGTop18_CONF2.ChangeType(MeritOperandType.CTGT);

            Operand_CTGTop18_CONF2.GetCellAt(2).IntegerValue = 3;

            Operand_CTGTop18_CONF2.Target = 0.1;

            Operand_CTGTop18_CONF2.Weight = 1;

            IMFERow Operand_CTGTop19_CONF2 = TheMFE.InsertNewOperandAt(19);

            Operand_CTGTop19_CONF2.ChangeType(MeritOperandType.CTGT);

            Operand_CTGTop19_CONF2.GetCellAt(2).IntegerValue = 3;

            Operand_CTGTop19_CONF2.Target = 0.1;

            Operand_CTGTop19_CONF2.Weight = 1;


            IMFERow Operand_DIVIop20_CONF2 = TheMFE.InsertNewOperandAt(20);

            Operand_DIVIop20_CONF2.ChangeType(MeritOperandType.DIVI);

            Operand_DIVIop20_CONF2.GetCellAt(2).IntegerValue = 12;

            Operand_DIVIop20_CONF2.GetCellAt(3).IntegerValue = 13;


            //---------------------------------------------------

            //Operand for 3rd Configuration

            // CONF 3

            IMFERow Operand_CONF3 = TheMFE.InsertNewOperandAt(21);

            Operand_CONF3.ChangeType(MeritOperandType.CONF);

            Operand_CONF3.GetCellAt(2).IntegerValue = 3;

            IMFERow Operand_REAYop22_CONF3 = TheMFE.InsertNewOperandAt(22);

            Operand_REAYop22_CONF3.ChangeType(MeritOperandType.REAY);

            Operand_REAYop22_CONF3.GetCellAt(2).IntegerValue = 1;

            Operand_REAYop22_CONF3.GetCellAt(7).DoubleValue = 1;

            IMFERow Operand_REAYop33_CONF3 = TheMFE.InsertNewOperandAt(23);

            Operand_REAYop33_CONF3.ChangeType(MeritOperandType.REAY);

            Operand_REAYop33_CONF3.GetCellAt(2).IntegerValue = 5;

            Operand_REAYop33_CONF3.GetCellAt(7).DoubleValue = 1;

            IMFERow Operand_RANGop24_CONF3 = TheMFE.InsertNewOperandAt(24);

            Operand_RANGop24_CONF3.ChangeType(MeritOperandType.RANG);

            Operand_RANGop24_CONF3.GetCellAt(2).IntegerValue = 4;

            Operand_RANGop24_CONF3.GetCellAt(7).DoubleValue = 1;

            Operand_RANGop24_CONF3.Target = 0;

            Operand_RANGop24_CONF3.Weight = 10;

            Operand_RANGop24_CONF3.GetCellAt(7).DoubleValue = 1;

            IMFERow Operand_EFLXop25_CONF3 = TheMFE.InsertNewOperandAt(25); // EFLX for operand 5

            Operand_EFLXop25_CONF3.ChangeType(MeritOperandType.EFLX);

            Operand_EFLXop25_CONF3.GetCellAt(2).IntegerValue = 2;

            Operand_EFLXop25_CONF3.GetCellAt(3).IntegerValue = 3;

            Operand_EFLXop25_CONF3.Target = Operand_EFLXop15_CONF2.GetCellAt(12).DoubleValue;

            Operand_EFLXop25_CONF3.Weight = 1;

            IMFERow Operand_EFLXop26_CONF3 = TheMFE.InsertNewOperandAt(26);

            Operand_EFLXop26_CONF3.ChangeType(MeritOperandType.EFLX);

            Operand_EFLXop26_CONF3.GetCellAt(2).IntegerValue = 4;

            Operand_EFLXop26_CONF3.GetCellAt(3).IntegerValue = 5;

            Operand_EFLXop26_CONF3.Target = Operand_EFLXop26_CONF3.GetCellAt(12).DoubleValue;

            Operand_EFLXop26_CONF3.Weight = 1;

            IMFERow Operand_EFLXop27_CONF3 = TheMFE.InsertNewOperandAt(27);

            Operand_EFLXop27_CONF3.ChangeType(MeritOperandType.EFLX);

            Operand_EFLXop27_CONF3.GetCellAt(2).IntegerValue = 6;

            Operand_EFLXop27_CONF3.GetCellAt(3).IntegerValue = 7;

            Operand_EFLXop27_CONF3.Target = Operand_EFLXop27_CONF3.GetCellAt(12).DoubleValue;

            Operand_EFLXop27_CONF3.Weight = 1;

            IMFERow Operand_CTGTop28_CONF3 = TheMFE.InsertNewOperandAt(28);

            Operand_CTGTop28_CONF3.ChangeType(MeritOperandType.CTGT);

            Operand_CTGTop28_CONF3.GetCellAt(2).IntegerValue = 3;

            Operand_CTGTop28_CONF3.Target = 0.1;

            Operand_CTGTop28_CONF3.Weight = 1;

            IMFERow Operand_CTGTop29_CONF3 = TheMFE.InsertNewOperandAt(29);

            Operand_CTGTop29_CONF3.ChangeType(MeritOperandType.CTGT);

            Operand_CTGTop29_CONF3.GetCellAt(2).IntegerValue = 3;

            Operand_CTGTop29_CONF3.Target = 0.1;

            Operand_CTGTop29_CONF3.Weight = 1;


            IMFERow Operand_DIVIop30_CONF3 = TheMFE.InsertNewOperandAt(30);

            Operand_DIVIop30_CONF3.ChangeType(MeritOperandType.DIVI);

            Operand_DIVIop30_CONF3.GetCellAt(2).IntegerValue = 22;

            Operand_DIVIop30_CONF3.GetCellAt(3).IntegerValue = 23;


            // Local optimisation till completion

            //ILocalOptimization LocalOpt = TheSystem.Tools.OpenLocalOptimization();

            //LocalOpt.Algorithm = OptimizationAlgorithm.DampedLeastSquares;

            //LocalOpt.Cycles = OptimizationCycles.Automatic;

            //LocalOpt.NumberOfCores = 12;

            //LocalOpt.RunAndWaitForCompletion();

            //LocalOpt.Close();

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

            Console.WriteLine("\nPlease close one of the two running Zemax Applications\n");

            throw new Exception(errorMessage);
        }

    }
}
