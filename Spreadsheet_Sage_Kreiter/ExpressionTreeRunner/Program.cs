// <copyright file="Program.cs" company="Sage Kreiter - 11659212">
// Copyright (c) Sage Kreiter. All rights reserved.
// </copyright>

namespace ExpressionTreeRunner
{
    using System;
    using SpreadsheetEngine;

    /// <summary>
    /// Runs a console app for the expression tree.
    /// </summary>
    internal static class Program
    {
        /// <summary>
        /// Main method to run app.
        /// </summary>
        /// <param name="args">PC input.</param>
        [STAThread]
        public static void Main(string[] args)
        {
            int choice = 0;
            string expression = "A1+B1+C1";
            ExpressionTree tree = new ExpressionTree(expression);

            while (choice != 4)
            {
                Console.WriteLine("Menu (Current Expression: " + expression + ")\n   1 = Enter a new expression\n   2 = Set a variable value\n   3 = Evaluate tree\n   4 = Quit");
                int.TryParse(Console.ReadLine(), out choice);

                switch (choice)
                {
                    case 1:
                        Console.WriteLine("Enter a new expression: ");
                        expression = Console.ReadLine();
                        tree = new ExpressionTree(expression);
                        break;

                    case 2:
                        Console.WriteLine("Enter a varibale name: ");
                        string name = Console.ReadLine();
                        Console.WriteLine("Enter variable value: ");
                        double num = double.Parse(Console.ReadLine());
                        tree.SetVariable(name, num);
                        break;

                    case 3:
                        Console.WriteLine(tree.Evaluate());
                        break;

                    default:
                        break;
                }
            }
        }
    }
}
