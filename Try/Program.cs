using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Timers;
using System;
using System.Threading;
using System.Linq;
using System.Runtime.InteropServices;

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

        public static void GetuserCards(Dictionary<string, List<string>> playerCards, Queue<string> players, Stack<string> deck)
        {
            foreach (var player in players)
            {
                List<string> cards = new List<string>();

                for (int i = 0; i < 13; i++) // Assuming you want to get 13 user cards
                {
                    cards.Add(deck.Pop());
                }
                playerCards.Add(player, cards);
            }
        }
        public static void displayUserCards(List<string> ucards)
        {
            Console.WriteLine("\nUser Cards:");
            int cardsPerRow = 5;
            int count = 0;
            foreach (string card in ucards)
            {
                Console.Write($"{card,-5}");
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
            time++;

            if (time == 6)
            {
                Console.WriteLine("\n");
                timer.Enabled = false;
            }
        }

        public static void setTimer()
        {
            time = 1;
            timer = new System.Timers.Timer(1000);
            timer.Elapsed += (sender, e) => TimerCallBack(e);
            timer.AutoReset = true;
            timer.Enabled = true;
            while (time < 6)
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
            Workbook excelBook = excelApp.Workbooks.Open(@"C:\Users\22-0202c\Downloads\deckofcards.xlsx"); //make sure this directory is correct
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
            Console.WriteLine("Welcome to Caravan!");
            Thread.Sleep(3000);
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
            Queue<string> players = new Queue<string>();
            Thread.Sleep(2000);
            players.Enqueue("Player 1");
            players.Enqueue("Player 2");
            players.Enqueue("Player 3");
            Console.WriteLine("\nPlayers who are playing the game: ");
            foreach (var player in players)
            {
                Console.WriteLine("\n" + player);
            }
            Dictionary<string, List<string>> playerCards = new Dictionary<string, List<string>>();
            deckOfCards(deckofCards);
            Console.Write("\nGenerate User Cards? [y/n]: ");
            string ans = Console.ReadLine().ToLower();
            if (ans == "y")
            {
                GetuserCards(playerCards, players, deckofCards);
                foreach (var player in playerCards.Reverse())
                {
                    Console.WriteLine($"\nGenerating {player.Key}'s cards in ");
                    setTimer();
                    //Console.WriteLine($"\n{player.Key}'s Cards: ");
                    int ucardsPerRow = 5;
                    int ucount = 0;
                    foreach (var card in player.Value)
                    {
                        Console.Write($"{card,-5}");
                        ucount++;
                        if (ucount == ucardsPerRow)
                        {
                            Console.WriteLine();
                            ucount = 0;
                        }
                    }
                }
                deckOfCards(deckofCards);
                /*Console.WriteLine("\nShuffled Cards:");
                int ucardsPerRow = 5;
                int ucount = 0;
                foreach (var card in deckofCards)
                {
                    Console.Write($"{card,-5}");
                    ucount++;
                    if (ucount == ucardsPerRow)
                    {
                        Console.WriteLine();
                        ucount = 0;
                    }
                }
                displayUserCards(usercards);
                deckOfCards(deckofCards);*/
            }
            else
            {
                return;
            }
            Console.Write("Draw First Card? [y/n]: ");
            string ans2 = Console.ReadLine().ToLower();
            if (ans2 == "y")
            {
                foreach (var player in playerCards)
                {
                    Console.WriteLine($"\nDrawing and getting card for {player.Key}: ");
                    setTimer();
                    player.Value.RemoveAt(0);
                    int mcardsPerRow = 5;
                    int mcount = 0;
                    foreach (var card in deckofCards)
                    {
                        Console.Write($"{card,-5}");
                        mcount++;
                        if (mcount == mcardsPerRow)
                        {
                            Console.WriteLine();
                            mcount = 0;
                        }
                    }
                }
                deckOfCards(deckofCards);
                /*int mcardsPerRow = 5;
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
                }*/
                
                /*playerCards.RemoveAt(0);
                displayUserCards(usercards);
                usercards.Add(deckofCards.Pop());
                Console.WriteLine("Shuffled Cards:");
                int mcardsPerRow = 5;
                int mcount = 0;
                foreach (var card in deckofCards)
                {
                    Console.Write($"{card,-5}");
                    mcount++;
                    if (mcount == mcardsPerRow)
                    {
                        Console.WriteLine();
                        mcount = 0;
                    }
                }
                displayUserCards(usercards);
                deckOfCards(deckofCards);*/
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