using System;
using System.Collections.Generic;
using SpreadsheetApp.Models;
using Cell = SpreadsheetApp.Models.Cell;

namespace SpreadsheetApp.Parsers
{
    public class ParserException : Exception
    {
        public ParserException(string message) : base(message) { }
    }

    public class ExpressionParser
    {
        private readonly Spreadsheet _spreadsheet;
        private readonly HashSet<string> _callStack;
        private Cell? _currentCell;
        private string _expression = string.Empty;
        private int _pos;
        private string _token = string.Empty;
        private TokenType _tokenType;

        private enum TokenType
        {
            None, Number, Identifier, Plus, Minus, Multiply, Divide,
            LessEqual, GreaterEqual, NotEqual, Not,
            LParen, RParen, Comma, EOF
        }

        public ExpressionParser(Spreadsheet spreadsheet, HashSet<string>? callStack = null)
        {
            _spreadsheet = spreadsheet;
            _callStack = callStack != null 
                ? new HashSet<string>(callStack, StringComparer.OrdinalIgnoreCase) 
                : new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        }

        public object Evaluate(string expression, Cell? currentCell = null)
        {
            if (string.IsNullOrWhiteSpace(expression))
                throw new ParserException("Порожній вираз");

            // Strip leading '=' if user used Excel notation
            expression = expression.Trim();
            if (expression.StartsWith("="))
                expression = expression.Substring(1).Trim();

            _currentCell = currentCell;
            if (_currentCell != null)
            {
                _callStack.Add(_currentCell.Name);
            }

            _expression = expression;
            _pos = 0;
            GetToken();

            if (_tokenType == TokenType.EOF)
                throw new ParserException("Порожній вираз");

            object result = ParseLogical();

            if (_tokenType != TokenType.EOF)
                throw new ParserException($"Неочікуваний токен: {_token}");

            return result;
        }

        private object ParseLogical()
        {
            if (_token.Equals("not", StringComparison.OrdinalIgnoreCase))
            {
                GetToken();
                object val = ParseLogical();
                if (val is bool b) return !b;
                throw new ParserException("Оператор 'not' очікує логічний операнд");
            }

            return ParseComparison();
        }

        private object ParseComparison()
        {
            object left = ParseArithmetic();

            while (_tokenType == TokenType.LessEqual || _tokenType == TokenType.GreaterEqual || _tokenType == TokenType.NotEqual)
            {
                TokenType op = _tokenType;
                GetToken();
                object right = ParseArithmetic();

                if (left is double l && right is double r)
                {
                    if (op == TokenType.LessEqual) left = l <= r;
                    else if (op == TokenType.GreaterEqual) left = l >= r;
                    else if (op == TokenType.NotEqual) left = l != r;
                }
                else if (left is bool lb && right is bool rb && op == TokenType.NotEqual)
                {
                    left = lb != rb;
                }
                else
                {
                    throw new ParserException("Неправильні типи операндів для порівняння");
                }
            }

            return left;
        }

        private object ParseArithmetic()
        {
            object left = ParseTerm();

            while (_tokenType == TokenType.Plus || _tokenType == TokenType.Minus)
            {
                TokenType op = _tokenType;
                GetToken();
                object right = ParseTerm();

                if (left is double l && right is double r)
                {
                    if (op == TokenType.Plus) left = l + r;
                    else left = l - r;
                }
                else
                {
                    throw new ParserException("Арифметичні операції підтримують лише числа");
                }
            }

            return left;
        }

        private object ParseTerm()
        {
            object left = ParseFactor();

            while (_tokenType == TokenType.Multiply || _tokenType == TokenType.Divide)
            {
                TokenType op = _tokenType;
                GetToken();
                object right = ParseFactor();

                if (left is double l && right is double r)
                {
                    if (op == TokenType.Multiply) left = l * r;
                    else
                    {
                        if (Math.Abs(r) < 1e-15) throw new ParserException("Ділення на нуль");
                        left = l / r;
                    }
                }
                else
                {
                    throw new ParserException("Арифметичні операції підтримують лише числа");
                }
            }

            return left;
        }

        private object ParseFactor()
        {
            object result = 0.0;

            if (_tokenType == TokenType.Plus)
            {
                GetToken();
                result = ParseFactor();
                if (result is not double) throw new ParserException("Унарний плюс очікує число");
            }
            else if (_tokenType == TokenType.Minus)
            {
                GetToken();
                result = ParseFactor();
                if (result is double d) result = -d;
                else throw new ParserException("Унарний мінус очікує число");
            }
            else if (_tokenType == TokenType.LParen)
            {
                GetToken();
                result = ParseLogical();
                if (_tokenType != TokenType.RParen) throw new ParserException("Очікується закриваюча дужка ')'");
                GetToken();
            }
            else if (_tokenType == TokenType.Number)
            {
                result = double.Parse(_token, System.Globalization.CultureInfo.InvariantCulture);
                GetToken();
            }
            else if (_tokenType == TokenType.Identifier)
            {
                string id = _token.ToUpper();
                GetToken();

                if (id == "TRUE")
                {
                    result = true;
                }
                else if (id == "FALSE")
                {
                    result = false;
                }
                else if (id == "MAX" || id == "MIN")
                {
                    if (_tokenType != TokenType.LParen) throw new ParserException($"Очікується '(' після {id}");
                    GetToken();
                    object arg1 = ParseLogical();
                    if (_tokenType != TokenType.Comma) throw new ParserException("Очікується ','");
                    GetToken();
                    object arg2 = ParseLogical();
                    if (_tokenType != TokenType.RParen) throw new ParserException("Очікується ')'");
                    GetToken();

                    if (arg1 is double d1 && arg2 is double d2)
                    {
                        result = id == "MAX" ? Math.Max(d1, d2) : Math.Min(d1, d2);
                    }
                    else
                    {
                        throw new ParserException($"Функція {id} очікує числові аргументи");
                    }
                }
                else
                {
                    // Cell reference
                    string normalizedId = Spreadsheet.NormalizeName(id);
                    if (_currentCell != null && !_currentCell.References.Contains(normalizedId))
                    {
                        _currentCell.References.Add(normalizedId);
                    }

                    var cell = _spreadsheet.GetCell(normalizedId);
                    if (cell == null) throw new ParserException($"Невідома клітинка: {id}");

                    if (cell.Value == null)
                    {
                        if (string.IsNullOrWhiteSpace(cell.Expression))
                        {
                            result = 0.0; // Порожня клітинка за замовчуванням 0
                        }
                        else
                        {
                            // Точкова перевірка циклів (Targeted cycle check)
                            if (_callStack.Contains(cell.Name))
                                throw new ParserException($"Циклічне посилання на клітинку {cell.Name}");

                            _callStack.Add(cell.Name);
                            try
                            {
                                var subParser = new ExpressionParser(_spreadsheet, _callStack);
                                cell.Value = subParser.Evaluate(cell.Expression, cell);
                                cell.Error = string.Empty;
                                result = cell.Value;
                            }
                            finally
                            {
                                _callStack.Remove(cell.Name);
                            }
                        }
                    }
                    else
                    {
                        result = cell.Value;
                    }
                }
            }
            else
            {
                throw new ParserException($"Синтаксична помилка, неочікуваний токен: {_token}");
            }

            return result;
        }

        private void GetToken()
        {
            while (_pos < _expression.Length && char.IsWhiteSpace(_expression[_pos]))
                _pos++;

            if (_pos >= _expression.Length)
            {
                _token = "";
                _tokenType = TokenType.EOF;
                return;
            }

            char c = _expression[_pos];

            if (char.IsDigit(c) || c == '.')
            {
                int start = _pos;
                while (_pos < _expression.Length && (char.IsDigit(_expression[_pos]) || _expression[_pos] == '.'))
                    _pos++;
                _token = _expression.Substring(start, _pos - start);
                _tokenType = TokenType.Number;
                return;
            }

            if (char.IsLetter(c))
            {
                int start = _pos;
                while (_pos < _expression.Length && char.IsLetterOrDigit(_expression[_pos]))
                    _pos++;
                _token = _expression.Substring(start, _pos - start);
                
                if (_token.Equals("not", StringComparison.OrdinalIgnoreCase))
                    _tokenType = TokenType.Not;
                else
                {
                    _token = Spreadsheet.NormalizeName(_token);
                    _tokenType = TokenType.Identifier;
                }
                
                return;
            }

            _token = c.ToString();
            _pos++;

            switch (c)
            {
                case '+': _tokenType = TokenType.Plus; break;
                case '-': _tokenType = TokenType.Minus; break;
                case '*': _tokenType = TokenType.Multiply; break;
                case '/': _tokenType = TokenType.Divide; break;
                case '(': _tokenType = TokenType.LParen; break;
                case ')': _tokenType = TokenType.RParen; break;
                case ',': _tokenType = TokenType.Comma; break;
                case '<':
                    if (_pos < _expression.Length && _expression[_pos] == '=')
                    {
                        _token = "<=";
                        _pos++;
                        _tokenType = TokenType.LessEqual;
                    }
                    else if (_pos < _expression.Length && _expression[_pos] == '>')
                    {
                        _token = "<>";
                        _pos++;
                        _tokenType = TokenType.NotEqual;
                    }
                    else
                    {
                        throw new ParserException("Очікується <= або <>");
                    }
                    break;
                case '>':
                    if (_pos < _expression.Length && _expression[_pos] == '=')
                    {
                        _token = ">=";
                        _pos++;
                        _tokenType = TokenType.GreaterEqual;
                    }
                    else
                    {
                        throw new ParserException("Очікується >=");
                    }
                    break;
                default:
                    throw new ParserException($"Невідомий символ: {c}");
            }
        }
    }
}

