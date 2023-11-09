using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Timers;
using System;
using System.Threading;

namespace StackCaravan
{
    internal class Program
    {
        private static System.Timers.Timer timer;

        public static int time = 1;

        public static void shuffleList(List<string> list)
        {
            Random rng = new Random();
            int n = list.Count;
            while (n > 1)
            {
                n--;
                int k = rng.Next(n + 1);
                string value = list[k];
                list[k] = list[n];
                list[n] = value;
            }
        }

        public static void GetuserCards(List<string> ucards, Stack<string> deck)
        {
            for (int i = 0; i < 13; i++) // Assuming you want to get 5 user cards
            {
                ucards.Add(deck.Pop());
            }
        }

        public static void displayUserCards(List<string> ucards)
        {
            Console.WriteLine("\nUser Cards:");
            int cardsPerRow = 5;
            int count = 0;
            foreach (string card in ucards)
            {
                Console.Write($"{card, -5}");
                count++;
                if (count == cardsPerRow)
                {
                    Console.WriteLine();
                    count = 0;
                }
            }
            Console.WriteLine();
        }

        public static void deckOfCards(Stack<string> deck)
        {
            Console.WriteLine("\nRemaining Cards in the Deck: " + deck.Count);
        }

        public static void TimerCallBack(Object o)
        {
            Console.Write(time + "      ");
            time--;

            if (time == 0)
            {
                Console.WriteLine("\n");
                timer.Enabled = false;
            }
        }

        public static void setTimer()
        {
            time = 5;
            timer = new System.Timers.Timer(1000);
            timer.Elapsed += (sender, e) => TimerCallBack(e);
            timer.AutoReset = true;
            timer.Enabled = true;
            while (time > 0)
            {
                
            }
            timer.Enabled = false;
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
            Workbook excelBook = excelApp.Workbooks.Open(@"C:\Users\rbuen\Downloads\deckofcards.xlsx");
            _Worksheet excelSheet = excelBook.Sheets[1];
            Range excelRange = excelSheet.UsedRange;
            int rows = excelRange.Rows.Count;
            int cols = excelRange.Columns.Count;
            for (int i = 1; i <= rows; i++)
            {
                for (int j = 1; j <= cols; j++)
                {
                    if (excelRange.Cells[i, j] != null && excelRange.Cells[i, j].Value2 != null)
                        list.Add(excelRange.Cells[i, j].Value2.ToString());
                }
            }
            Console.WriteLine("Shuffling Card in ");
            setTimer();
            shuffleList(list);
            Console.WriteLine("Shuffled Cards:");
            int cardsPerRow = 5;
            int count = 0;
            foreach (var card in list)
            {
                Console.Write($"{card,-5}");
                count++;
                if (count == cardsPerRow)
                {
                    Console.WriteLine();
                    count =0;
                }
            }
            Stack<string> deckofCards = new Stack<string>(list);
            deckOfCards(deckofCards);
            Console.Write("\nGenerate User Cards? [y/n]: ");
            string ans = Console.ReadLine().ToLower();
            if (ans == "y")
            {
                Console.WriteLine("Generating user cards in ");
                setTimer();
                GetuserCards(usercards, deckofCards);
                Console.WriteLine("\nShuffled Cards:");
                int ucardsPerRow = 5;
                int ucount = 0;
                foreach (var card in deckofCards)
                {
                    Console.Write($"{card, -5}");
                    ucount++;
                    if (ucount == ucardsPerRow)
                    {
                        Console.WriteLine();
                        ucount = 0;
                    }
                }
                displayUserCards(usercards);
                deckOfCards(deckofCards);
            }
            else
            {
                return;
            }
            Console.Write("Draw First Card? [y/n]: ");
            string ans2 = Console.ReadLine().ToLower();
            if (ans2 == "y")
            {
                usercards.RemoveAt(0);
                displayUserCards(usercards);
                usercards.Add(deckofCards.Pop());
                Console.WriteLine("Shuffled Cards:");
                int mcardsPerRow = 5;
                int mcount = 0;
                foreach (var card in deckofCards)
                {
                    Console.Write($"{card, -5}");
                    mcount++;
                    if (mcount == mcardsPerRow)
                    {
                        Console.WriteLine();
                        mcount = 0;
                    }
                }
                displayUserCards(usercards);
                deckOfCards(deckofCards);
            }
            else
            {
                return;
            }
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            Console.ReadKey();
        }
    }
}