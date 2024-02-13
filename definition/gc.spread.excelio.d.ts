declare module GC{
    module Spread{
        module Excel{

            export class IO{
                /**
                 * Represents an excel import and export class.
                 * @class
                 */
                constructor();
                /**
                 * Imports an excel file.
                 * @param {Blob} file The excel file.
                 * @param {function} successCallBack Call this function after successfully loading the file. `function (json) { }`.
                 * @param {function} errorCallBack Call this function if an error occurs. The exception parameter object structure `{ errorCode: GC.Spread.Excel.IO.ErrorCode, errorMessage: string}`.
                 * @param {GC.Spread.Excel.IO.OpenOptions} options The options for import excel.
                 * @returns {void}
                 * @example
                 * ```
                 * var workbook = new GC.Spread.Sheets.Workbook(document.getElementById("ss"));
                 * var excelIO = new GC.Spread.Excel.IO();
                 * var excelFile = document.getElementById("fileDemo").files[0];
                 * excelIO.open(excelFile, function (json) {
                 *    workbook.fromJSON(json);
                 * }, function (e) {
                 *    console.log(e);
                 * }, {
                 *    password: "password",
                 *    importPictureAsFloatingObject: false
                 * });
                 * ```
                 */
                open(file: Blob,  successCallBack: Function,  errorCallBack?: Function,  options?: GC.Spread.Excel.IO.OpenOptions): void;
                /**
                 * Register a unknown max digit width info to ExcelIO.
                 * @param {string} fontFamily The font family of default style's font.
                 * @param {number} fontSize The font size of default style's font(in point).
                 * @param {number} maxDigitWidth The  max digit width of default style's font.
                 * @returns {void}
                 */
                registerMaxDigitWidth(fontFamily: string,  fontSize: number,  maxDigitWidth: number): void;
                /**
                 * Creates and saves an excel file with the SpreadJS json.
                 * @param {object} json The spread sheets json object, or string.
                 * @param {function} successCallBack Call this function after successfully exporting the file. `function (blob) { }`.
                 * @param {function} errorCallBack Call this function if an error occurs. The exception parameter object structure `{ errorCode: GC.Spread.Excel.IO.ErrorCode, errorMessage: string}`.
                 * @param {GC.Spread.Excel.IO.SaveOptions} options The options for export excel.
                 * @returns {void}
                 * @example
                 * ```
                 * var workbook = new GC.Spread.Sheets.Workbook(document.getElementById("ss"));
                 * var excelIO = new GC.Spread.Excel.IO();
                 * var json = JSON.stringify(workbook.toJSON());
                 * excelIO.save(json, function (blob) {
                 *    saveAs(blob, fileName); //saveAs is from FileSaver.
                 * }, function (e) {
                 *    console.log(e);
                 * }, {
                 *    password: "password",
                 *    xlsxStrictMode: false
                 * });
                 * ```
                 */
                save(json: string | object,  successCallBack: Function,  errorCallBack?: Function,  options?: GC.Spread.Excel.IO.SaveOptions): void;
            }
            module IO{

                /**
                 * @typedef GC.Spread.Excel.IO.OpenOptions - The options for import excel.
                 * @property {string} password the excel file's password.
                 * @property {boolean} importPictureAsFloatingObject import picture as floatingObject instead of shape.
                 */
                export type OpenOptions = 
                    {
                        password?: string;
                        importPictureAsFloatingObject?: boolean;
                    }
                

                /**
                 * @typedef GC.Spread.Excel.IO.SaveOptions - The options for export excel.
                 * @property {string} password the excel file's password.
                 * @property {boolean} xlsxStrictMode the mode of exporting process, the non-strict mode may reduce the export size. Default is true.
                 */
                export type SaveOptions = 
                    {
                        password?: string;
                        xlsxStrictMode?: boolean;
                    }
                
                /**
                 * Specifies the excel io error code.
                 * @enum {number}
                 */
                export enum ErrorCode{
                    /**
                     *  File read and write exception.
                     */
                    fileIOError= 0,
                    /**
                     *  Incorrect file format.
                     */
                    fileFormatError= 1,
                    /**
                     *  The Excel file cannot be opened because the workbook/worksheet is password protected.
                     */
                    noPassword= 2,
                    /**
                     *  The specified password is incorrect.
                     */
                    invalidPassword= 3
                }

            }

        }

    }

}
