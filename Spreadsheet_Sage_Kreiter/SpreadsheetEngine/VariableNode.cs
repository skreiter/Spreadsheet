// <copyright file="Node.cs" company="Sage Kreiter - 11659212">
// Copyright (c) Sage Kreiter. All rights reserved.
// </copyright>

namespace SpreadsheetEngine
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Class that holds a variable leaf node.
    /// </summary>
    public class VariableNode : Node
    {
        /// <summary>
        /// Gets or Sets value of Name.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Helps to evaluate expression tree by returning the variables value from the dictionary.
        /// </summary>
        /// <param name="dic">Dictionary with all variables.</param>
        /// <returns>Variable value.</returns>
        public override double Evaluate(Dictionary<string, double> dic)
        {
            double num = 0;
            num = dic[this.Name];

            return num;
        }
    }
}
