import Cell from "./Cell";
import SheetMemory from "./SheetMemory";
import { ErrorMessages } from "./GlobalDefinitions";

export class FormulaEvaluator {
  private _errorOccured: boolean = false;
  private _errorMessage: string = "";
  private _currentFormula: FormulaType = [];
  private _lastResult: number = 0;
  private _result: number = 0;
  private _sheetMemory: SheetMemory;

  constructor(memory: SheetMemory) {
    this._sheetMemory = memory;
  }

  public evaluate(formula: FormulaType): void {
    this.resetForNewFormula(formula);

    if (this._currentFormula.length === 0) {
      this._errorMessage = ErrorMessages.emptyFormula;
      this._result = 0;
      return;
    }

    this._result = this.calculateFormula();

    if (this._currentFormula.length > 0 && !this._errorOccured) {
      this._errorMessage = ErrorMessages.invalidFormula;
      this._errorOccured = true;
    }

    if (this._errorOccured) {
      this._result = this._lastResult;
    }
  }

  public get error(): string {
    return this._errorMessage;
  }

  public get result(): number {
    return this._result;
  }

  private resetForNewFormula(formula: FormulaType): void {
    this._currentFormula = [...formula];
    this._lastResult = 0;
    this._errorMessage = "";
    this._errorOccured = false;
  }

  private calculateFormula(): number {
    if (this._errorOccured) return this._lastResult;

    let result = this.parseToken();
    while (this._currentFormula.length > 0 && this.isAddOrSubtractOperator()) {
      const operator = this._currentFormula.shift();
      const token = this.parseToken();
      result += operator === "+" ? token : -token;
    }
    this._lastResult = result;
    return result;
  }

  private isAddOrSubtractOperator(): boolean {
    const firstToken = this._currentFormula[0];
    return firstToken === "+" || firstToken === "-";
  }

  private parseToken(): number {
    if (this._errorOccured) return this._lastResult;

    let result = this.getValue();
    while (this._currentFormula.length > 0 && this.isMultiplyOrDivideOperator()) {
      const operator = this._currentFormula.shift();

      if (operator === "+/-") {
        result = result === 0 ? 0 : -result;
        continue;
      }

      const number = this.getValue();
      result = operator === "*" ? result * number : result / number;

      if (number === 0) {
        this.handleError(ErrorMessages.divideByZero, Infinity);
        return Infinity;
      }
    }
    this._lastResult = result;
    return result;
  }

  private isMultiplyOrDivideOperator(): boolean {
    const token = this._currentFormula[0];
    return token === "*" || token === "/" || token === "+/-";
  }

  private getValue(): number {
    if (this._errorOccured) return this._lastResult;

    const token = this._currentFormula.shift();
    
    // If the token is a number
    if (this.isNumber(token)) {
      const value = Number(token);
      this._lastResult = value;
      return value;
    }

    // If the token is an opening parenthesis
    if (token === "(") {
      const value = this.calculateFormula();
      if (!this._currentFormula.length || this._currentFormula.shift() !== ")") {
        this.handleError(ErrorMessages.missingParentheses, value);
        return value;
      }
      this._lastResult = value;
      return value;
    }

    // If the token is a cell reference
    if (this.isCellReference(token)) {
      const [value, error] = this.getCellValue(token);
      if (error) {
        this.handleError(error, value);
        return value;
      }
      this._lastResult = value;
      return value;
    }

    // Default error handling for an unrecognized token
    this.handleError(ErrorMessages.invalidFormula);
    return 0;
}


  private isNumber(token: TokenType): boolean {
    return !isNaN(Number(token));
  }

  private handleError(message: string, result = 0): void {
    this._errorMessage = message;
    this._errorOccured = true;
    this._lastResult = result;
  }

  /**
   *
   * @param token
   * @returns true if the token is a cell reference
   *
   */
  isCellReference(token: TokenType): boolean {
    return Cell.isValidCellLabel(token);
  }

  /**
   *
   * @param token
   * @returns [value, ""] if the cell formula is not empty and has no error
   * @returns [0, error] if the cell has an error
   * @returns [0, ErrorMessages.invalidCell] if the cell formula is empty
   *
   */
  getCellValue(token: TokenType): [number, string] {
    let cell = this._sheetMemory.getCellByLabel(token);
    let formula = cell.getFormula();
    let error = cell.getError();

    // if the cell has an error return 0
    if (error !== "" && error !== ErrorMessages.emptyFormula) {
      return [0, error];
    }

    // if the cell formula is empty return 0
    if (formula.length === 0) {
      return [0, ErrorMessages.invalidCell];
    }

    let value = cell.getValue();
    return [value, ""];
  }
}

export default FormulaEvaluator;
