using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Threading;
using OpenCvSharp;
using OpenCvSharp.Extensions;
using WindowsInput;
using WindowsInput.Native;
using Point = OpenCvSharp.Point;
using Size = System.Drawing.Size;

namespace SimpleScript
{
    public class Execute
    {

        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport("user32.dll", SetLastError = true)]
        static extern bool GetWindowInfo(IntPtr hwnd, ref WINDOWINFO pwi);

        [DllImport("user32.dll")]
        static extern IntPtr GetWindowRect(IntPtr hWnd, ref Rect rect);

        [DllImport("user32.dll")]
        public static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        [DllImport("user32.dll")]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        static extern int GetSystemMetrics(int smIndex);

        [DllImport("user32.dll", SetLastError = true)]
        static extern uint SendInput(uint nInputs, ref INPUT pInputs, int cbSize);

        [DllImport("user32.dll")]
        static extern void mouse_event(uint dwFlags, int dx, int dy, uint dwdata, IntPtr dwExtraInfo);

        const uint LEFTDOWN = 0x0002;
        const uint LEFTUP = 0x0004;

        public WINDOWINFO windowInfo { get; set; }

        public Dimensions dimensions { get; set; }

        InputSimulator sim = new InputSimulator();

        public bool _IsCheck = false;

        public struct Dimensions
        {
            public int width;
            public int height;
        }

        public void MainLoop()
        {

            for (int i = 7; i > 0; i--)
            {
                Console.WriteLine(i.ToString());
                Thread.Sleep(999);
            }

            Console.WriteLine("ПУСК!!!");


            while (true)
            {

                GetAllWindowInfo();


                Mat FainPhotoZero = Cv2.ImRead("E:\\Photo\\Faza44.png");
                IsCheck(in FainPhotoZero);
                if (_IsCheck)
                {
                    Console.WriteLine("Закидываем удочку");
                    TheBeginningOfFishing();
                }

                Mat FaindPthotoOne = Cv2.ImRead("E:\\Photo\\Faza5555.png");
                IsCheck(in FaindPthotoOne);
                if (_IsCheck)
                {
                    Console.WriteLine("В ожидание улова...");
                    Fishing();
                    Console.WriteLine("Успех!");
                }

                Mat FaindPthotoTwo = Cv2.ImRead("E:\\Photo\\Faza66.png");
                IsCheck(in FaindPthotoTwo);
                if (_IsCheck)
                {
                    Console.WriteLine("Идет процесс моимки рыбы!");

                    
                    for (int i = 0; i <= 10; i++)
                    {
                        
                        TighteningTheFishingLine();
                        
                    }

                    Console.WriteLine("Ура рыба пойманно!");
                    int animationSleepTime = new Random().Next(4000, 8000);
                    Thread.Sleep(animationSleepTime);

                    if (new Random().Next(1, 4) == 2)
                    {
                        Console.WriteLine("АнтиАФК");
                        ExitFromAFK();
                    }
                }


            }


        }

        void TheBeginningOfFishing()
        {
            int castingBaseTime = 1000;
            int castingRandom = 300;
            Random random = new Random();
            sim.Keyboard.KeyDown(VirtualKeyCode.MENU);
            int castingTime = (castingBaseTime + (castingRandom + random.Next(30, 400)));
            mouse_event(LEFTDOWN, 0, 0, 0, IntPtr.Zero);
            Thread.Sleep(castingTime);
            mouse_event(LEFTUP, 0, 0, 0, IntPtr.Zero);
            sim.Keyboard.KeyUp(VirtualKeyCode.MENU);
        }

        void Fishing()
        { 
            sim.Keyboard.KeyDown(VirtualKeyCode.MENU);
            mouse_event(LEFTDOWN, 0, 0, 0, IntPtr.Zero);
            Thread.Sleep(200);
            mouse_event(LEFTUP, 0, 0, 0, IntPtr.Zero);
            Thread.Sleep(100);
            sim.Keyboard.KeyUp(VirtualKeyCode.MENU);
        }

        void TighteningTheFishingLine()
        {
            int lineSlackTime = 1100;
            Random random = new Random();
            int castingTimeRnd = lineSlackTime + random.Next(100, 300);
            sim.Keyboard.KeyDown(VirtualKeyCode.MENU);
            mouse_event(LEFTDOWN, 0, 0, 0, IntPtr.Zero);
            Thread.Sleep(castingTimeRnd);
            mouse_event(LEFTUP, 0, 0, 0, IntPtr.Zero);
            sim.Keyboard.KeyUp(VirtualKeyCode.MENU);
            Thread.Sleep(castingTimeRnd - 200);
        }

        void ExitFromAFK()
        {
            sim.Keyboard.KeyDown(VirtualKeyCode.MENU);
            sim.Keyboard.KeyPress(VirtualKeyCode.VK_A);
            Thread.Sleep(100);
            sim.Keyboard.KeyPress(VirtualKeyCode.VK_A);
            Thread.Sleep(400);
            sim.Keyboard.KeyPress(VirtualKeyCode.VK_D);
            Thread.Sleep(100);
            sim.Keyboard.KeyPress(VirtualKeyCode.VK_D);
            sim.Keyboard.KeyUp(VirtualKeyCode.MENU);
        }

        public void IsCheck(in Mat FaindPthoto)
        {
            Mat objectPhoto = BitmapConverter.ToMat(makeScreenshot());
            objectPhoto.ConvertTo(objectPhoto, MatType.CV_32FC1);
            Mat gray1 = new Mat();
            Cv2.CvtColor(objectPhoto, gray1, ColorConversionCodes.BGR2GRAY);


            FaindPthoto.ConvertTo(FaindPthoto, MatType.CV_32FC1);
            Mat gray2 = new Mat();
            Cv2.CvtColor(FaindPthoto, gray2, ColorConversionCodes.BGR2GRAY);
            Mat result = new Mat();

            Cv2.MatchTemplate(gray1, gray2, result, TemplateMatchModes.CCoeffNormed);//The best match is 1, the smaller the value, the worse the match

            Double minVul = 0.0d;
            Double maxVul = 0.0d;
            Point minLoc = new Point(0, 0);
            Point maxLoc = new Point(0, 0);
            Point matchLoc = new Point(0, 0);

            Cv2.MinMaxLoc(result, out minVul, out maxVul, out minLoc, out maxLoc);
            Cv2.Threshold(result, result, 0.8, 1, ThresholdTypes.Tozero);


            Cv2.WaitKey();
            result.Release();
            GC.Collect();

            
            double threshold = 0.8;

            if (maxVul > threshold)
            {
                _IsCheck = true;
            }
            else
            {
                _IsCheck = false;
            }
        }

        public Bitmap makeScreenshot()
        {
            Bitmap screenshot = new Bitmap(dimensions.width, dimensions.height, PixelFormat.Format24bppRgb);

            Graphics gfxScreenshot = Graphics.FromImage(screenshot);

            gfxScreenshot.CopyFromScreen(windowInfo.rcClient.Left, windowInfo.rcClient.Top, 0, 0, new Size(dimensions.width, dimensions.height), CopyPixelOperation.SourceCopy);

            return screenshot;
        }


        public void GetAllWindowInfo()
        {
            //Это тот процесс, который мы ищем. Вы можете найти это, наведя курсор на процесс в нижней панели и используя это имя.
            //Вы также можете выполнить поиск по всем активным окнам и найти его таким образом: https://stackoverflow.com/questions/7268302/get-the-titles-of-all-open-windows
            string processName = "New World";
            IntPtr hwndWindow = FindWindow(null, processName);

            //Если указатель равен 0, это означает, что мы ничего не нашли
            if (hwndWindow == (IntPtr)0)
                throw new Exception($"Unable to find {processName} as an active process.");

            //Otherwise we happily use the information :^D
            WINDOWINFO windowInfo = new WINDOWINFO();
            GetWindowInfo(hwndWindow, ref windowInfo);

            //Store the window information for later use
            this.windowInfo = windowInfo;

            //Same for the dimensions of the screen -> this means it can be updated by resizing the client/moving the client
            this.dimensions = new Dimensions
            {
                width = windowInfo.rcClient.Right - windowInfo.rcClient.Left,
                height = windowInfo.rcClient.Bottom - windowInfo.rcClient.Top,
            };
        }

    }
}
