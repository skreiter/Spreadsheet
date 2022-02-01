// <copyright file="ExpressionTree.cs" company="Sage Kreiter - 11659212">
// Copyright (c) Sage Kreiter. All rights reserved.
// </copyright>

namespace SpreadsheetEngine
{
    using System.Collections.Generic;

    /// <summary>
    /// Class to create and evaluate expressions.
    /// </summary>
    public class ExpressionTree
    {
        private Node root;

        private Dictionary<string, double> dic = new Dictionary<string, double>();

        /// <summary>
        /// Initializes a new instance of the <see cref="ExpressionTree"/> class.
        /// </summary>
        /// <param name="expression">String expression to be put into the tree.</param>
        public ExpressionTree(string expression)
        {
            this.root = this.TreeCreator(expression);
        }

        /// <summary>
        /// Gets the dictionary.
        /// </summary>
        public Dictionary<string, double> Dict
        {
            get { return this.dic; }
        }

        /// <summary>
        /// Evaluates the expression to a value.
        /// </summary>
        /// <returns>Expression value.</returns>
        public double Evaluate()
        {
            if (this.root != null)
            {
                return this.root.Evaluate(this.dic);
            }

            return 0;
        }

        /// <summary>
        /// Sets a variable of a given name to a given value.
        /// </summary>
        /// <param name="variableName">Requested name.</param>
        /// <param name="variableValue">Requested value.</param>
        public void SetVariable(string variableName, double variableValue)
        {
            this.dic[variableName] = variableValue;
        }

        private Node TreeCreator(string expression)
        {
            expression = expression.Replace(" ", string.Empty); // Getting rid of white space.

            if (expression == null || expression == string.Empty)
            {
                ConstantNode constant = new ConstantNode { Value = 0 };
                return constant;
            }

            // Solving expression inside of ( recursively.
            if (expression[0] == '(')
            {
                int parenthesis = 1; // Keeps track whether we have closed all ().

                // Since first value is a ( we start at 2nd value.
                for (int i = 1; expression.Length > i; i++)
                {
                    if (expression[i] == ')')
                    {
                        parenthesis--;
                        if (parenthesis == 0)
                        {
                            if (i != expression.Length - 1)
                            {
                                break;
                            }
                            else
                            {
                                return this.TreeCreator(expression.Substring(1, expression.Length - 2));
                            }
                        }
                    }

                    if (expression[i] == '(')
                    {
                        parenthesis++; // Add 1 because there is another ( we need to make sure gets closed.
                    }
                }
            }

            int index = this.GetIndex(expression);

            // If the index is -1 then we know it is either a var or num so we return that as a node.
            if (index == -1)
            {
                double value;
                if (double.TryParse(expression, out value))
                {
                    ConstantNode constant = new ConstantNode { Value = value };
                    return constant;
                }
                else if (!this.dic.ContainsKey(expression))
                {
                    this.dic[expression] = 0;
                }

                VariableNode variable = new VariableNode { Name = expression };
                return variable;
            }

            // If program gets here then we must create an operator node.
            OperatorNode newOperator = new OperatorNode(expression[index])
            {
                Left = this.TreeCreator(expression.Substring(0, index)),
                Right = this.TreeCreator(expression.Substring(index + 1)),
            };

            return newOperator;
        }

        private int GetIndex(string expression)
        {
            int index = -1;
            int parenthesis = 0;

            for (int i = 0; i < expression.Length; i++)
            {
                switch (expression[i])
                {
                    case '*':
                    case '/': // If the expression is * or / and index has not been set then set index
                        if (index == -1 && parenthesis == 0)
                        {
                            index = i;
                        }

                        break;

                    case '+':
                    case '-':
                        if (parenthesis == 0)
                        {
                            return i; // If it finds - or + then it returns because those should be lower in the tree than * or /.
                        }

                        break;

                    case '(':
                        parenthesis++;
                        break;

                    case ')':
                        parenthesis--;
                        break;
                }
            }

            return index;
        }
    }
}
