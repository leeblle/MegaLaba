using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.IO;
using System.IO.MemoryMappedFiles;
using System.Diagnostics;
using System.Windows.Input;
using System.ComponentModel;

namespace COMObj1
{
    [Guid("5FAC6B03-6D95-46ed-A3C1-B0DBEE34D022"),
    InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IMyEvents
    {
    }


    [Guid("01D10028-A89D-47ee-8048-C27B6DD4BE63")]
    public interface ICreature
    {
        [DispId(1)]
        int X { get; set; }
        int Y { get; set; }
        char[] Pattern { get; set; }
        int[] MapSize { get; set; }
        int Health { get; set; }
        void ChangePattern(char[] Pattern);
        void Print();
    }

    [Guid("7A8865A2-A931-4D81-B2DA-F1108A570DA6")]
    public interface IPlayer
    {
        [DispId(2)]
        void Move(int dx);
        void Shoot();
    }

    [Guid("6AF3B5DA-21F3-45C8-9D85-C5CA7B4D5C4A")]
    public interface IEnemy
    {
        [DispId(3)]
        void Move(int dx, int dy);
    }

    [Guid("81C2A6BE-4ABC-4815-B95B-52B510D25C71")]
    public interface IItem
    {
        [DispId(4)]
        int Speed { get; set; }
        bool Move(int dy);
    }





    // создаем класс для обработчика
    class myKeyEventArgs : HandledEventArgs
    {
        // нажатая кнопка
        public ConsoleKeyInfo key;
        public myKeyEventArgs(ConsoleKeyInfo _key)
        {
            key = _key;
        }
    }

    // класс события
    class KeyEvent
    {
        // событие нажатия
        public event EventHandler<myKeyEventArgs> KeyPress;
        // метод запуска события
        public void OnKeyPress(ConsoleKeyInfo _key)
        {
            KeyPress(this, new myKeyEventArgs(_key));
        }
    }

    [Guid("67C74931-2601-465F-A490-7B8CF9E1F13F"),
    ComSourceInterfaces(typeof(IMyEvents)),
    ComVisible(true)]
    public abstract class Creature : ICreature
    {
        public int X { get; set; }
        public int Y { get; set; }
        public char[] Pattern { get; set; }
        public int[] MapSize { get; set; }

        public static Semaphore Semaphore;

        public int Health
        {
            get { return health; }
            set
            {
                health = value;
                if (value <= 0)
                    Destroy?.Invoke(this);
            }
        }

        public delegate void DestroyHandler(Creature obj);
        public event DestroyHandler Destroy;
        public static bool win = false;
        public static char[] BlowPattern, EmptyPattern;
        private int health;

        public Creature(char[] Pattern, int[] StartPoint, int[] MapSize, int Health)
        {
            X = StartPoint[0];
            Y = StartPoint[1];
            this.Pattern = new char[Pattern.Length];
            Array.Copy(Pattern, this.Pattern, Pattern.Length);
            this.MapSize = MapSize;
            this.Health = Health;
        }

        public void ChangePattern(char[] Pattern)
        {
            this.Pattern = Pattern;
        }

        public void Print()
        {
            for (int i = 0; i < Pattern.Length; i += 5)
            {
                Console.SetCursorPosition(X, Y + i / 5);
                Console.WriteLine("" + Pattern[i] + Pattern[i + 1] + Pattern[i + 2] + Pattern[i + 3] + Pattern[i + 4]);
            }
        }

        public void Print(ConsoleColor clr)
        {
            Console.ForegroundColor = clr;
            for (int i = 0; i < Pattern.Length; i += 5)
            {
                Console.SetCursorPosition(X, Y + i / 5);
                Console.WriteLine("" + Pattern[i] + Pattern[i + 1] + Pattern[i + 2] + Pattern[i + 3] + Pattern[i + 4]);
            }
            Console.ResetColor();
        }
    }


    [Guid("5594556B-7D95-44ED-BD05-2C276C803504"),
    ComVisible(true),
    ComSourceInterfaces(typeof(IMyEvents))]
    public class Player : Creature, IPlayer
    {
        public delegate void RocketHandler(int[] cords, int speed);
        public event RocketHandler Shooter;
        public delegate void HealthHandler();
        public event HealthHandler HealthChange;
        public delegate void ActionHandler(int act);
        public event ActionHandler KeyAction;

        public bool ShootEnablePlayer = true;
        TimerCallback tmplayercallback;
        Timer timerplayer;
        TimerCallback tmplayercallbackheal;
        Timer timerplayersheal;
        KeyEvent kevt;
        public Player(char[] pattern, int[] StartPoint, int[] MapSize, int Health, int CoolDown)
            : base(pattern, StartPoint, MapSize, Health)
        {
            Print();
            tmplayercallback = new TimerCallback(ShootEnableMeth);
            timerplayer = new Timer(tmplayercallback, 2, 0, CoolDown);
            tmplayercallbackheal = new TimerCallback(HealPleyer);
            timerplayersheal = new Timer(tmplayercallbackheal, 2, 0, 500);
            kevt = new KeyEvent();
            kevt.KeyPress += KeyPress;
        }

        public void ExternalKeyPress(ConsoleKeyInfo key) => kevt.OnKeyPress(key);

        private void KeyPress(object sender, myKeyEventArgs e)
        {
            char ch = e.key.KeyChar;
            if (char.ToLower(ch) != 'a' && char.ToLower(ch) != 'd' && char.ToLower(ch) != 'r' && char.ToLower(ch) != 'l' && char.ToLower(ch) != 'x' && ch != ' ')
            {
                e.Handled = true;
            }
            else
            {
                //действия
                switch (char.ToLower(ch))
                {
                    case 'a': Move(-1); break;
                    case 'd': Move(1); break;
                    case ' ': Shoot(); break;
                    case 'r': KeyAction?.Invoke(1); break;
                    case 'x': KeyAction?.Invoke(2); break;
                    case 'l': KeyAction?.Invoke(3); break;
                }

            }
        }

        public void Move(int dx)
        {
            Semaphore.WaitOne();
            if (X + dx > 0 && X + dx < MapSize[0] && !win)
            {
                X += dx;
                Print();
            }
            Semaphore.Release();
        }
        public void Tick()
        {
            while (true)
            {
                if (Health <= 0)
                    break;
                Semaphore.WaitOne();
                Print(ConsoleColor.Blue);
                Semaphore.Release();

            }
        }
        public void HealPleyer(object obj)
        {
            if (Health < 100)
            {
                Health += 1;
                Semaphore.WaitOne();
                HealthChange?.Invoke();
                Semaphore.Release();
            }
        }
        public void ShootEnableMeth(object obj)
        {
            ShootEnablePlayer = true;
        }

        public void Shoot()
        {
            if (ShootEnablePlayer)
            {
                ShootEnablePlayer = false;

                Mutex mutex = Mutex.OpenExisting("CollisionMutex");
                mutex.WaitOne();

                using (MemoryMappedFile mmf = MemoryMappedFile.OpenExisting("RocketsMap"))
                {
                    using (MemoryMappedViewStream stream = mmf.CreateViewStream((X + 2) + MapSize[0] * (Y - 4), 0))
                    {
                        BinaryWriter writer = new BinaryWriter(stream);
                        writer.Write(1);
                    }
                }
                mutex.ReleaseMutex();

                Shooter?.Invoke(new int[] { X + 2, Y - 4 }, -1);

            }
        }
    }

    [Guid("FA3A0E69-307E-4BC8-8AC3-14A78100D5F5"),
    ComVisible(true),
    ComSourceInterfaces(typeof(IMyEvents))]
    public class Enemy : Creature, IEnemy
    {
        public delegate void EnemysHandler();
        public event EnemysHandler DefeatPattern;
        public delegate void RocketHandler(int[] cords, int speed);
        public event RocketHandler Shooter;

        private int x1, x2;
        private bool isRight;
        private int Speed;
        private ConsoleColor Enclr = ConsoleColor.White;

        public int Chance { get; set; }

        public Enemy(char[] pattern, int[] StartPoint, int[] MapSize, int Health, int Speed, int Chance)
            : base(pattern, StartPoint, MapSize, Health)
        {
            x1 = 10;
            x2 = 50;
            isRight = true;
            this.Speed = Speed;
            this.Chance = Chance;
            if (this.Health > 20) Enclr = ConsoleColor.Red;
        }

        public void Move(int dx, int dy)
        {
            if (x1 > 0 && x2 > 0)
            {
                X += dx * (isRight ? 1 : -1);
                x1 += dx * (isRight ? 1 : -1);
                x2 += dx * (!isRight ? 1 : -1);
            }
            else
            {
                if (!isRight) { x1 = 1; x2--; }
                if (isRight) { x2 = 1; x1--; }
                isRight = !isRight;
                Y += dy;
                X += dx * (isRight ? 1 : -1);
                if (Y >= MapSize[1])
                {
                    DefeatPattern?.Invoke();
                }
            }
        }

        public void Tick()
        {
            while (true)
            {
                Semaphore.WaitOne();
                if (Health > 0 && !win)
                {
                    this.Move(1, 1);
                    this.Print(Enclr);
                    Random rnd = new Random();
                    if (rnd.Next(Chance) == 1)
                        Shooter?.Invoke(new int[] { X, Y + 4 }, 1);
                    Semaphore.Release();
                    Thread.Sleep(this.Speed);
                }
                else
                {
                    if (!win)
                        Print(Enclr);
                    Semaphore.Release();
                    break;
                }
            }

        }
    }


    [Guid("7BD4CD6E-E006-456E-913C-724785641F5B"),
    ComVisible(true),
    ComSourceInterfaces(typeof(IMyEvents))]
    public class Item : Creature, IItem
    {
        public delegate void ItemDestroyHandler(Creature obj);
        public event ItemDestroyHandler ItemDestroy;

        public int Speed { get; set; }

        public Item(char[] pattern, int[] StartPoint, int[] MapSize, int Speed, int Health)
            : base(pattern, StartPoint, MapSize, Health)
        {
            this.Speed = Speed;
        }

        public bool Move(int dy)
        {
            if (Y + dy > 0 && Y + dy < MapSize[1] && Health > 0)
            {

                Mutex mutex = Mutex.OpenExisting("CollisionMutex");
                mutex.WaitOne();
                using (MemoryMappedFile mmf = MemoryMappedFile.OpenExisting("RocketsMap"))
                {
                    byte type = 1;
                    using (MemoryMappedViewStream stream = mmf.CreateViewStream(X + MapSize[0] * Y, 0))
                    {
                        BinaryReader writer = new BinaryReader(stream);
                        type = writer.ReadByte();
                    }
                    using (MemoryMappedViewStream stream = mmf.CreateViewStream(X + MapSize[0] * Y, 0))
                    {
                        BinaryWriter writer = new BinaryWriter(stream);
                        writer.Write(0);
                    }
                    using (MemoryMappedViewStream stream = mmf.CreateViewStream(X + MapSize[0] * (Y + dy), 0))
                    {
                        BinaryWriter writer = new BinaryWriter(stream);
                        writer.Write(type);
                    }
                }
                mutex.ReleaseMutex();
                Y += dy;
                return true;
            }
            else return false;
        }

        public void Tick()
        {
            while (true)
            {
                Semaphore.WaitOne();
                if (this.Move(Speed) && !win)
                {
                    this.Print(ConsoleColor.Yellow);
                    Semaphore.Release();
                    Thread.Sleep(40);
                }
                else
                {
                    Console.SetCursorPosition(X, Y + 2);
                    Console.Write(' ');
                    Mutex mutex = Mutex.OpenExisting("CollisionMutex");
                    mutex.WaitOne();
                    using (MemoryMappedFile mmf = MemoryMappedFile.OpenExisting("RocketsMap"))
                    {
                        using (MemoryMappedViewStream stream = mmf.CreateViewStream(X + MapSize[0] * Y, 0))
                        {
                            BinaryWriter writer = new BinaryWriter(stream);
                            writer.Write(0);
                        }
                    }
                    
                    
                    mutex.ReleaseMutex();
                    ItemDestroy?.Invoke(this);
                    if (Pattern != BlowPattern)
                    {
                        Item blow = new Item(BlowPattern, new int[] { X, Y }, MapSize, 0, 1);
                        blow.Print();
                        Thread.Sleep(40);
                        blow.ChangePattern(EmptyPattern);
                        blow.Print();
                    }
                    Semaphore.Release();
                    
                    break;
                }
            }
        }

        public new void Print(ConsoleColor clr)
        {
            Console.ForegroundColor = clr;
            for (int i = 0; i < Pattern.Length; i += 1)
            {
                Console.SetCursorPosition(X, Y + i);
                Console.WriteLine("" + Pattern[i]);
            }
            Console.ResetColor();
        }
    }

    [Guid("91F1509B-E2B4-4541-BB6A-D3C2CED8A011"),
    ComVisible(true),
    ComSourceInterfaces(typeof(IMyEvents))]
    public class Shield : Creature
    {
        public Shield(char[] pattern, int[] StartPoint, int[] MapSize, int Health)
            : base(pattern, StartPoint, MapSize, Health)
        {
        }
    }
}
