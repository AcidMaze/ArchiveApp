﻿using System;
using System.Text;
using System.Windows.Forms;

namespace ArchiveApp
{
    internal static class Program
    {
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new HelloForm());//Форма приветствия при запуске программы
            Application.Run(new Form1());

        }

    }

}
