using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Trending_Toward_Writeoff
{
    public static class SpinAnimation
    {
        private static System.ComponentModel.BackgroundWorker m_spinner = initialiseBackgroundWorker();
        private static int m_spinnerPosition = 25;
        private static int m_spinWait = 25;
        private static bool m_isRunning;

        public static bool IsRunning { get { return m_isRunning; } }


        private static System.ComponentModel.BackgroundWorker initialiseBackgroundWorker()
        {

            System.ComponentModel.BackgroundWorker obj = new System.ComponentModel.BackgroundWorker();
            obj.WorkerSupportsCancellation = true;
            obj.DoWork += delegate
            {
                m_spinnerPosition = Console.CursorLeft;
                while (!obj.CancellationPending)
                {
                    char[] spinChars = new char[] { '|', '/', '-', '\\' };
                    foreach (char spinChar in spinChars)
                    {
                        Console.CursorLeft = m_spinnerPosition;
                        Console.Write(spinChar);
                        System.Threading.Thread.Sleep(m_spinWait);
                    }
                }
            };
            return obj;
        }

        /// <summary>
        /// Start the animation
        /// </summary>
        /// <param name="spinWait">wait time between spin steps in milliseconds</param>
        public static void Start(int spinWait)
       {
            m_isRunning = true;
            SpinAnimation.m_spinWait = spinWait;
            if (!m_spinner.IsBusy)
            {
                m_spinner.RunWorkerAsync();
            }
            else 
            {
                throw new InvalidOperationException("Cannot start spinner whilst spinner is already running");
            }
        }

        public static void Start() { Start(25); }

        public static void Stop()
       {
            m_spinner.CancelAsync();
            while (m_spinner.IsBusy) System.Threading.Thread.Sleep(100);
            Console.CursorLeft = m_spinnerPosition;
            m_isRunning = false;
        }
    }
    }

