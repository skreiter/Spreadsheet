// <copyright file="Node.cs" company="Sage Kreiter - 11659212">
// Copyright (c) Sage Kreiter. All rights reserved.
// </copyright>

namespace SpreadsheetEngine
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Holds operator and its children.
    /// </summary>
    public class OperatorNode : Node
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="OperatorNode"/> class.
        /// </summary>
        /// <param name="c">Operator to be set.</param>
        public OperatorNode(char c)
        {
            this.Operator = c;
            this.Left = this.Right = null;
        }

        /// <summary>
        /// Gets or sets the operator.
        /// </summary>
        public char Operator { get; set; }

        /// <summary>
        /// Gets or Sets the Left node.
        /// </summary>
        public Node Left { get; set; }

        /// <summary>
        /// Gets or Sets the Right node.
        /// </summary>
        public Node Right { get; set; }

        /// <summary>
        /// Helps evaluate the tree by recursively going through nodes.
        /// </summary>
        /// <param name="dic">Dictionary of variables.</param>
        /// <returns>Value of expression at current point.</returns>
        public override double Evaluate(Dictionary<string, double> dic)
        {
            double evalLeft = this.Left.Evaluate(dic);
            double evalRight = this.Right.Evaluate(dic);
            double val = 0;

            switch (this.Operator)
            {
                case '+':
                    val = evalLeft + evalRight;
                    break;

                case '*':
                    val = evalRight * evalLeft;
                    break;

                case '-':
                    val = evalLeft - evalRight;
                    break;

                case '/':
                    val = evalLeft / evalRight;
                    break;
            }

            return val;
        }
    }
}
