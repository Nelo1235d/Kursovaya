using System;
using System.Collections;
using System.Linq;
using System.Xml.Linq;
using ClosedXML.Excel;



  
         public class SortShell
{
    public static int k = 0;
    public static int m = 0;
    public  static void Main(string[] args)
        {
             Console.WriteLine("Введите шаг: ");
        int inc = 0;
        inc = Convert.ToInt32(Console.ReadLine());
        Console.WriteLine("Введите элементы: ");
            
            int a = 0;
            a = Convert.ToInt32(Console.ReadLine());
            int[] arr = new int[a];
            int[] arr1 = new int[a];
            int n;
            n = arr.Length;
            string p = "";
            int z1 = 0;
            string[] k1 = { };
            Word(p, k1, arr, arr1, a, z1);
            shellSort(arr, n);
            Console.WriteLine("\nОтсортированные элементы:");
            show_array_elements(arr);
            Console.WriteLine("\n в обратном порядке");
            show_reverse(arr);
            Ex(arr, arr1, a, z1);


        }
        public static void Word(string p, string[] k1, int[] arr, int[] arr1, int a, int z1)
        {
            Aspose.Words.Document doc = new Aspose.Words.Document("Elements.docx");
            p = Convert.ToString(doc.Range.Text);
            int l1 = 0;
            k1 = p.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);
            Console.WriteLine("Начальный массив");
            for (int z = 9; z < a + 9; z++)
            {
                bool isNum = int.TryParse(k1[z], out l1);
                if (isNum)
                {
                    try
                    {
                        arr[z - 9] = Int32.Parse(k1[z]);
                        arr1[z - 9] = Int32.Parse(k1[z]);
                        Console.WriteLine(arr[z - 9]);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e);
                    }
                }
                else
                {
                    z1++;
                }
            }
        }
      public  static void shellSort(int[] arr, int array_size)
        {
            int i, j, inc, temp;
            inc = 3;
            while (inc > 0)
            {
                for (i = 0; i < array_size; i++)
                {
                    j = i;
                    temp = arr[i];
                    while ((j >= inc) && (arr[j - inc] > temp))
                    {
                        arr[j] = arr[j - inc];
                        j = j - inc;
                         arr[j] = temp;
                    k = k + 1;
                }
               
            }
                if (inc / 2 != 0)
                    inc = inc / 2;

                else if (inc == 1)
                    inc = 0;

                else
                    inc = 1;
                 

            }
      
        Console.WriteLine(" кол-во сравнений {0} ", k,m);
        }

      public  static void show_array_elements(int[] arr)
        {
            foreach (var element in arr)
            {
                Console.Write(element + " ");

            }
            Console.Write("\n");

        }
  public      static void show_reverse(int[] arr)
        {

            {

                Array.Reverse(arr);

            }
            Console.Write(String.Join(' ', arr));

        }
        public static void Ex(int[] arr, int[] arr1, int f, int z1)
        {
            int i = 1;
            var path = Path.Combine(Environment.CurrentDirectory, "Export", "Elem.xlsx");
            var wb = new XLWorkbook();
            var sh = wb.Worksheets.Add("Element");
            sh.Cell(1, 1).SetValue("Начальный массив");
            sh.Cell(1, 2).SetValue("Конечный массив");
            sh.Cell(1, 3).SetValue("Кол-во сравнений");
            sh.Cell(1, 4).SetValue("Кол-во перестановок");
            sh.Cell(2, 3).SetValue(k);
            sh.Cell(2, 4).SetValue(m);
        for (i = 1; i < f + z1; i++)
            {
                sh.Cell(i + 1, 1).SetValue(arr1[i]);
                sh.Cell(i + 1, 2).SetValue(arr[i]);
            }
            wb.SaveAs(path);
       
    }
   
}
