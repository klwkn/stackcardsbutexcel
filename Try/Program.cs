using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StackCaravan
{
    internal class Program
    {
        public static int time = 1;
        public static void shuffleList(List<string> list)
        {

        }
        public static void GetuserCards(List<string> ucards, Stack<string> deck)
        {

        }
        public static void displayUserCards(Stack<string> ucards)
        {

        }
        public static void deckOfCards(Stack<string> deck)
        {

        }
        public static void TimerCallBack(Object o)
        {

        }
        static void Main(string[] args)
        {
            List<string> list = new List<string>();
            List<string> usercards = new List<string>();
            Application excelApp = new Application();
            if (excelApp == null)
            {
                Console.WriteLine("Excel is not installed!");
                return;
            }
            Workbook excelBook = excelApp.Workbooks.Open(@"C:\Users\22-0202c\Downloads\deckofcards (1).xlsx");
            _Worksheet excelSheet = excelBook.Sheets[1];
            Range excelRange = excelSheet.UsedRange;
            int rows = excelRange.Rows.Count;
            int cols = excelRange.Columns.Count;
            for(int i = 1; i <= rows; i++)
            {
                for(int j = 1; j <= cols; j++)
                {
                    if (excelRange.Cells[i, j] != null && excelRange.Cells[i, j].Value2 != null)
                        list.Add(excelRange.Cells[i, j].Value2.ToString());
                }
            }
            Console.WriteLine("Shuffling Card in ");
            //setTimer();
            shuffleList(list);
            Stack<string> deckofCards = new Stack<string>(list);
            deckOfCards(deckofCards);
            Console.WriteLine("haha");
            Console.ReadLine();
        }
    }
}