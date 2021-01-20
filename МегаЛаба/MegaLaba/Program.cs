using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.IO;
using System.IO.MemoryMappedFiles;
using System.Diagnostics;
using System.ComponentModel;

using Microsoft.VisualBasic;
using COMObj1;


namespace MegaLaba
{
    public class Server
    {
        public static MemoryMappedFile RocketCoord;
        public static Semaphore Semaphore;
        private List<Item> Rockets = new List<Item>();
        private List<Shield> Shields = new List<Shield>();
        private Player Player;
        private int Score;
        private int LVLnumber;
        public static bool Exit;
        public static bool Restart;
        public int[] MapSize = new int[] { 116, 45 }; //116 55
        private List<Enemy> Enemys;
        private bool ShootEnableEnemy;
        TimerCallback tmenemycallback;
        Timer timerenemy;
        dynamic WinDefMessage;

        public Server(Dictionary<int, int> Enemys, int Score, int LVLnumber) ///Entire
        {
            WinDefMessage = (dynamic)Microsoft.VisualBasic.Interaction.GetObject(@"script:C:\Windows\SysWOW64\MegaLabaGame.wsc", null);
            Semaphore = new Semaphore(1, 1);
            Creature.Semaphore = Semaphore;
            Player = new Player(((List<char[]>)Program.Settings[0])[4], new int[] { MapSize[0] / 2, MapSize[1] - 5 }, MapSize, (int)((List<object>)Program.Settings[1])[3], (int)((List<Object>)Program.Settings[1])[2]);
            Player.Shooter += CreateRocket;
            Player.Destroy += Defeat;
            Player.HealthChange += PrintScreen;
            Player.KeyAction += ServerAction;


            MakeShields(MapSize);
            Thread playerth = new Thread(Player.Tick);
            playerth.Start();
            this.Enemys = MakeEnemys(Enemys, MapSize);
            this.Score = Score;
            this.LVLnumber = LVLnumber;
            ShootEnableEnemy = true;
            Creature.BlowPattern = ((List<char[]>)Program.Settings[0])[6];
            Creature.EmptyPattern = ((List<char[]>)Program.Settings[0])[3];

            tmenemycallback = new TimerCallback(ShootEnableMeth);
            timerenemy = new Timer(tmenemycallback, 1, 0, 1000);
            Thread CollisionServer = new Thread(this.ServerTick);
            CollisionServer.Start();
        }

        public void ExternalKeyPress(ConsoleKeyInfo key) => Player.ExternalKeyPress(key);
        public void ServerAction(int act)
        {
            switch(act)
            {
                case 1: Restart = true; break;

                case 2: Exit = true; break;

                case 3:
                    MSScriptControl.ScriptControl sc = new MSScriptControl.ScriptControl();
                    sc.Language = "VBScript";
                    sc.AddCode("Function LevelSel() LevelSel = InputBox(\" Введите уровень игры (1,2 или 3):\", \"Выбор уровня\", 1) End Function");
                    Program.level = int.Parse(sc.Run("LevelSel"));
                    Restart = true;
                    break;
            }
        }

        public List<Enemy> MakeEnemys(Dictionary<int, int> Enemys, int[] MapSize) ///Entire
        {
            int offset = 10;
            List<Enemy> list = new List<Enemy>();

            int last = 0;
            foreach (var elem in Enemys)
            {
                int i;
                for (i = last; i < elem.Key + last; i++)
                {
                    Enemy a = new Enemy(((List<char[]>)Program.Settings[0])[5], new int[] { offset + (i % 5) * offset, (i / 5) * 5 }, MapSize, elem.Value, (int)((List<Object>)Program.Settings[1])[6], (int)((List<Object>)Program.Settings[1])[7]);
                    a.DefeatPattern += Killed;
                    a.Destroy += EnemyDefeat;
                    a.Shooter += CreateRocket;
                    Thread th = new Thread(a.Tick);
                    th.Name = i.ToString() + "Enemy";
                    th.Start();
                    list.Add(a);
                }
                last = i;
            }
            return list;
        }

        public void MakeShields(int[] MapSize) //Entire
        {
            int offset = 20;
            List<Item> list = new List<Item>();

            for (int i = 0; i <= (MapSize[0] - offset) / (5 + offset); i++)
            {
                Shield a = new Shield(((List<char[]>)Program.Settings[0])[1], new int[] { (i % 5 + 1) * offset, Player.Y - 10 }, MapSize, 40);
                a.Destroy += EnemyDefeat;
                a.Print(ConsoleColor.DarkGreen);
                Shields.Add(a);
            }
        }

        public void ShieldDamage(Shield obj) => obj.ChangePattern(((List<char[]>)Program.Settings[0])[2]);

        public void EnemyDefeat(Creature obj) //Enemy
        {
            obj.ChangePattern(((List<char[]>)Program.Settings[0])[3]);
            if (obj as Enemy != null)
            {
                Score += 10;
                Enemys.Remove((Enemy)obj);
                if (Enemys.Count == 0)
                    Win();
                PrintScreen();
            }
            if (obj as Shield != null)
                Shields.Remove((Shield)obj);
        }

        public void ShootEnableMeth(object obj) => ShootEnableEnemy = true;

        public void CreateRocket(int[] cords, int speed) //Player
        {
            if (speed > 0 && ShootEnableEnemy)
            {
                Item a = new Item(((List<char[]>)Program.Settings[0])[0], new int[] { cords[0], cords[1] }, MapSize, speed * (int)((List<Object>)Program.Settings[1])[1], 1);
                a.Destroy += RocketBlow;
                a.ItemDestroy += RocketBlow;

                Thread th = new Thread(a.Tick);
                th.Start();

                Rockets.Add(a);
                ShootEnableEnemy = false;

                Mutex mutex = Mutex.OpenExisting("CollisionMutex");
                mutex.WaitOne();

                using (MemoryMappedFile mmf = MemoryMappedFile.OpenExisting("RocketsMap"))
                {
                    using (MemoryMappedViewStream stream = mmf.CreateViewStream((cords[0]) + MapSize[0] * (cords[1]), 0))
                    {
                        BinaryWriter writer = new BinaryWriter(stream);
                        writer.Write(2);
                    }
                }
                mutex.ReleaseMutex();
            }
            if (speed < 0)
            {
                Item a = new Item(((List<char[]>)Program.Settings[0])[0], new int[] { cords[0], cords[1] }, MapSize, speed * (int)((List<Object>)Program.Settings[1])[1], 1);
                a.Destroy += RocketBlow;
                a.ItemDestroy += RocketBlow;

                Thread th = new Thread(a.Tick);
                th.Start();

                Rockets.Add(a);
            }
        }

        public void ServerTick()
        {
            while (!Exit && !Restart && !Creature.win)
            {
                Semaphore.WaitOne();
                Dictionary<int[], int> RocketsMemoryMap = new Dictionary<int[], int>();  //Список ракет с их координатами из памяти

                Mutex mutex1 = Mutex.OpenExisting("CollisionMutex");
                using (MemoryMappedFile mmf = MemoryMappedFile.OpenExisting("RocketsMap"))
                {
                    mutex1.WaitOne();
                    for (int i = 0; i < MapSize[0] * MapSize[1]; i++)  //Проходим по всей памяти размером с карту X*Y
                    {
                        using (MemoryMappedViewStream stream = mmf.CreateViewStream(i, 0))
                        {
                            BinaryReader reader = new BinaryReader(stream);
                            Byte a = reader.ReadByte();
                            if (a != 0)
                                RocketsMemoryMap.Add(new int[] { i % 116, i / 116 }, a); //Добавляем ракету со значением 1/2 (игрока или врага) и ключем-координатами
                        }
                    }
                    mutex1.ReleaseMutex();
                }
                try
                {
                    foreach (Enemy elem in Enemys)    //Перебор врагов
                    {
                        foreach (KeyValuePair<int[], int> memrocket in RocketsMemoryMap)   //Проверяем каждую ракету, не попала ли она в хитбокс
                        {
                            if (memrocket.Key[0] >= elem.X && memrocket.Key[0] <= elem.X + 4 && memrocket.Key[1] + 1 >= elem.Y && memrocket.Key[1] <= elem.Y + 3 && memrocket.Value == 1)
                            {
                                elem.Health -= (int)((List<Object>)Program.Settings[1])[4]; //Наносим урон врагу


                                foreach (Item rocket in Rockets)
                                    if (rocket.X == memrocket.Key[0] && rocket.Y == memrocket.Key[1])
                                        rocket.Health = 0;
                            }
                        }
                    }
                }
                catch { }
                try
                {
                    foreach (Shield elem in Shields)  //Перебор щитов
                    {
                        foreach (KeyValuePair<int[], int> memrocket in RocketsMemoryMap)   //Проверяем каждую ракету, не попала ли она в хитбокс
                        {
                            if (memrocket.Key[0] >= elem.X && memrocket.Key[0] <= elem.X + 4 && memrocket.Key[1] >= elem.Y && memrocket.Key[1] <= elem.Y + 3)
                            {

                                elem.Health -= (int)((List<Object>)Program.Settings[1])[4];
                                if (elem.Health > 0)
                                {
                                    elem.ChangePattern(((List<char[]>)Program.Settings[0])[2]);
                                    elem.Print();
                                }

                                foreach (Item rocket in Rockets)
                                    if (rocket.X == memrocket.Key[0] && rocket.Y == memrocket.Key[1])
                                        rocket.Health = 0;
                            }
                        }
                    }
                }
                catch { }
                //Игрок
                try
                {
                    foreach (KeyValuePair<int[], int> memrocket in RocketsMemoryMap)   //Проверяем каждую ракету, не попала ли она в наш хитбокс
                    {
                        if (memrocket.Key[0] >= Player.X + 1 && memrocket.Key[0] <= Player.X + 3 && memrocket.Key[1] + 1 >= Player.Y && memrocket.Key[1] <= Player.Y + 4 && memrocket.Value == 2)
                        {
                            Player.Health -= (int)((List<Object>)Program.Settings[1])[5];
                            PrintScreen();

                            foreach (Item rocket in Rockets)
                                if (rocket.X == memrocket.Key[0] && rocket.Y == memrocket.Key[1])
                                    rocket.Health = 0;
                        }
                    }
                }
                catch { }
                try
                {
                    //Ракеты
                    foreach (KeyValuePair<int[], int> memrocket1 in RocketsMemoryMap)   //Проверяем каждую ракету
                    {
                        foreach (KeyValuePair<int[], int> memrocket2 in RocketsMemoryMap)   //Проверяем каждую ракету
                        {
                            if (memrocket1.Value != memrocket2.Value && memrocket1.Key[0] == memrocket2.Key[0] && Math.Abs(memrocket1.Key[1] - memrocket2.Key[1]) == 1)
                            {
                                foreach (Item rocket in Rockets)
                                {
                                    if (rocket.X == memrocket1.Key[0] && Math.Abs(rocket.Y - memrocket1.Key[1]) <= 1 || rocket.X == memrocket2.Key[0] && Math.Abs(rocket.Y - memrocket1.Key[1]) <= 1)
                                    {
                                        rocket.Health = 0;
                                    }
                                }
                            }
                        }
                    }
                }
                catch { }
                Semaphore.Release();
                Thread.Sleep(10);

            }
        }

        public void RocketBlow(Creature objRocket) => Rockets.Remove((Item)objRocket);
        public void Killed() => Defeat(null);
        public void Defeat(Creature obj) ///Entire
        {
            //ВЫ ПРОИГРАЛИ
            if (!Creature.win)
            {
                Creature.win = true;
                Enemys.Clear();
                Rockets.Clear();
                Shields.Clear();
                Console.Clear();
                ServerAction((int)WinDefMessage.CreateDefeat(Score.ToString(), LVLnumber.ToString()));
            }
        }

        public void Win()
        {
            //ВЫ ВЫИГРАЛИ
            if (!Creature.win)
            {
                Creature.win = true;
                Enemys.Clear();
                Rockets.Clear();
                Shields.Clear();
                Console.Clear();
                if (Program.level < 3) Program.level++;
                Program.score = Score;
                ServerAction((int)WinDefMessage.CreateWin(Score.ToString(), LVLnumber.ToString()));
            }
        }

        public void PrintScreen() ///Entire
        {
            Console.SetCursorPosition(112, 1);
            Console.WriteLine("{0,4}", Score);
            Console.SetCursorPosition(113, 2);
            Console.WriteLine("{0,3}", Player.Health);
        }
    }

    class Program
    {
        public static List<object> Settings = new List<object>();
        public static int score;
        public static int level;
        static void excel(int level)
        {
            List<char[]> Patterns = new List<char[]>();
            List<object> Level = new List<object>();
            Dictionary<int, int> Enemys = new Dictionary<int, int>();


            dynamic excelApp = Activator.CreateInstance(
              Type.GetTypeFromProgID("Excel.Application"));
            excelApp.Visible = false;
            var workbook = excelApp.Workbooks.Open(Filename: @"C:\Levels.xlsx", ReadOnly: true);

            dynamic workSheet = excelApp.ActiveSheet;

            Patterns.Add(workSheet.Cells[2, "A"].Value.ToString().ToCharArray(0, 5));  //RoPa
            Patterns.Add(workSheet.Cells[2, "B"].Value.ToString().ToCharArray(0, 25));  //ShPa
            Patterns.Add(workSheet.Cells[2, "C"].Value.ToString().ToCharArray(0, 25));  //ShDmPa
            Patterns.Add(workSheet.Cells[2, "D"].Value.ToString().ToCharArray(0, 25));  //EmptPa
            Patterns.Add(workSheet.Cells[2, "E"].Value.ToString().ToCharArray(0, 25));  //PlPa
            Patterns.Add(workSheet.Cells[2, "F"].Value.ToString().ToCharArray(0, 25));  //EnPa
            Patterns.Add(workSheet.Cells[2, "G"].Value.ToString().ToCharArray(0, 25));  //BlPa

            string enem = workSheet.Cells[level + 3, "A"].Value.ToString();
            string[] enems1 = enem.Split(',');
            foreach (string vaenem in enems1)
            {
                if (vaenem != "")
                {
                    int a = int.Parse(vaenem.Split(' ')[0].ToString());
                    int b = int.Parse(vaenem.Split(' ')[1].ToString());
                    Enemys.Add(a, b);
                }
            }
            Level.Add(Enemys);                                                       //Enemys Dict
            Level.Add(int.Parse(workSheet.Cells[level + 3, "B"].Value.ToString()));  //Speed rocket
            Level.Add(int.Parse(workSheet.Cells[level + 3, "C"].Value.ToString()));  //Cooldown
            Level.Add(int.Parse(workSheet.Cells[level + 3, "D"].Value.ToString()));  //PlayerHP
            Level.Add(int.Parse(workSheet.Cells[level + 3, "E"].Value.ToString()));  //Damage
            Level.Add(int.Parse(workSheet.Cells[level + 3, "F"].Value.ToString()));  //Enemy Damage
            Level.Add(int.Parse(workSheet.Cells[level + 3, "G"].Value.ToString()));  //Speed enemy
            Level.Add(int.Parse(workSheet.Cells[level + 3, "H"].Value.ToString()));  //Enemy chance of shoot
            Settings.Add(Patterns);
            Settings.Add(Level);
            excelApp.Quit();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
        }

        static void Main(string[] args)
        {

            if (args.Length > 0)
            {
                level = int.Parse(args[0]);
                score = int.Parse(args[1]);

            }
            else
            {
                level = 1;
                score = 0;
            }

            Console.Clear();
            Settings.Clear();
            excel(level);
            Console.SetCursorPosition(105, 1);
            Console.WriteLine("Score: {0,4}", score);
            Console.SetCursorPosition(105, 2);
            Console.WriteLine("Health: 100");
            Console.SetCursorPosition(105, 3);
            Console.WriteLine("Level: " + level);

            Server.RocketCoord = MemoryMappedFile.CreateNew("RocketsMap", 10000);
            Mutex Mutex = new Mutex(false, "CollisionMutex");
            Server server = new Server((Dictionary<int, int>)((List<object>)Settings[1])[0], score, level);

            Console.SetWindowSize(server.MapSize[0] + 10, server.MapSize[1] + 10);
            Console.Title = "Space Invaders";
            Console.CursorVisible = false;
            Creature.win = false;
            Server.Restart = false;
            Server.Exit = false;

            while (!Server.Exit && !Server.Restart)
            {
                ConsoleKeyInfo key = Console.ReadKey(true);
                server.ExternalKeyPress(key);
            }
            if (Server.Restart)
            {
                System.Diagnostics.Process p = new System.Diagnostics.Process();
                p.StartInfo.FileName = @"MegaLaba.exe";
                p.StartInfo.Arguments = level.ToString() + " " + score.ToString();
                p.Start();
            }

            Server.RocketCoord.Dispose();
            Process.GetCurrentProcess().Kill();
        }
    }
}
