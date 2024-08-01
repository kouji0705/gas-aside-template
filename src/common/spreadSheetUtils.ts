/**
 * Copyright 2024 k.k.Factory
 */
import { isNullish } from "./common";
import {
	createInvalidPositionError,
	createInvalidPositionsError,
	createSheetNotFoundError,
} from "./errorUtils";

export class SpreadsheetUtil {
	private spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet;

	constructor() {
		this.spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
	}

	/**
	 * 指定したシート名のシートを取得する。
	 *
	 * @param sheetName - シートの名称
	 * @returns 指定したシート
	 */
	getSheetByName = (sheetName: string): GoogleAppsScript.Spreadsheet.Sheet => {
		const sheet = this.spreadsheet.getSheetByName(sheetName);
		if (isNullish(sheet)) {
			throw createSheetNotFoundError(sheetName);
		}
		return sheet;
	};

	/**
	 * 指定した範囲を取得する。
	 *
	 * @param sheetName - シートの名称
	 * @param startRowIndex - 開始行
	 * @param startColumnIndex - 開始列
	 * @param rowCount - 取得する行数
	 * @param columnCount - 取得する列数
	 * @returns 指定したセル範囲
	 */
	getRange = (
		sheetName: string,
		startRowIndex: number,
		startColumnIndex: number,
		rowCount: number,
		columnCount: number,
	): GoogleAppsScript.Spreadsheet.Range => {
		const sheet = this.spreadsheet.getSheetByName(sheetName);
		if (isNullish(sheet)) {
			throw createSheetNotFoundError(sheetName);
		}
		return sheet.getRange(
			startRowIndex,
			startColumnIndex,
			rowCount,
			columnCount,
		);
	};

	/**
	 * 指定した範囲を取得する。
	 *
	 * @param sheetName - シートの名称
	 * @param positions - セルの位置 'B2:B15'
	 * @returns 指定したセル範囲
	 */
	getRangeByPositions = (
		sheetName: string,
		positions: string,
	): GoogleAppsScript.Spreadsheet.Range => {
		const sheet = this.spreadsheet.getSheetByName(sheetName);
		if (isNullish(sheet)) {
			throw createSheetNotFoundError(sheetName);
		}
		const positionsPattern = /^[A-Z]+[0-9]+:[A-Z]+[0-9]+$/;
		if (!positionsPattern.test(positions)) {
			throw createInvalidPositionError(positions);
		}
		return sheet.getRange(positions);
	};

	/**
	 * 指定したセルの値を取得する。
	 *
	 * @param sheetName - シートの名称
	 * @param position - セルの位置 'B2'
	 * @returns 取得したセルの値
	 */
	getCellValue = (sheetName: string, position: string): string => {
		const sheet = this.spreadsheet.getSheetByName(sheetName);
		if (isNullish(sheet)) {
			throw createSheetNotFoundError(sheetName);
		}
		const positionPattern = /^[A-Z]+[0-9]+$/;
		if (!positionPattern.test(position)) {
			throw createInvalidPositionError(position);
		}
		return sheet.getRange(position).getValue();
	};

	/**
	 * 指定したセル範囲の情報を取得する。
	 *
	 * @param sheetName - シートの名称
	 * @param positions - セルの範囲 'A6:X385'
	 * @returns 取得した行の配列
	 */
	getRowsByPositions = (sheetName: string, positions: string): string[][] => {
		const sheet = this.spreadsheet.getSheetByName(sheetName);
		if (isNullish(sheet)) {
			throw createSheetNotFoundError(sheetName);
		}
		const positionsPattern = /^[A-Z]+[0-9]+:[A-Z]+[0-9]+$/;
		if (!positionsPattern.test(positions)) {
			throw createInvalidPositionsError(positions);
		}
		return sheet.getRange(positions).getValues();
	};

	/**
	 * 指定したセル範囲の情報を取得する。
	 *
	 * @param sheetName - シートの名称
	 * @param startRowIndex - 開始行
	 * @param startColumnIndex - 開始列
	 * @param rowCount - 取得する行数
	 * @param columnCount - 取得する列数
	 * @returns 取得した行の配列
	 */
	getRows = (
		sheetName: string,
		startRowIndex: number,
		startColumnIndex: number,
		rowCount: number,
		columnCount: number,
	): string[][] => {
		const sheet = this.spreadsheet.getSheetByName(sheetName);
		if (isNullish(sheet)) {
			throw createSheetNotFoundError(sheetName);
		}
		return sheet
			.getRange(startRowIndex, startColumnIndex, rowCount, columnCount)
			.getValues();
	};

	/**
	 * 指定したセル範囲に値を設定する。
	 *
	 * @param sheetName - シートの名称
	 * @param rowNo - 設定開始セルの行番号
	 * @param columnNo - 設定開始セルの列番号
	 * @param rowCount - 設定する行数
	 * @param columnCount - 設定する列数
	 * @param values - 設定する値の二次元配列
	 * @returns void
	 */
	setValues = (
		sheetName: string,
		rowNo: number,
		columnNo: number,
		rowCount: number,
		columnCount: number,
		values: string[][],
	): void => {
		const sheet = this.spreadsheet.getSheetByName(sheetName);
		if (isNullish(sheet)) {
			throw createSheetNotFoundError(sheetName);
		}
		sheet.getRange(rowNo, columnNo, rowCount, columnCount).setValues(values);
	};
	/**
	 * 指定したセルに値を設定する。
	 *
	 * @param sheetName - シートの名称
	 * @param rowNo - 設定セルの行番号
	 * @param columnNo - 設定セルの列番号
	 * @param value - 設定する値
	 * @returns void
	 */
	setValue = (
		sheetName: string,
		rowNo: number,
		columnNo: number,
		value: string,
	): void => {
		const sheet = this.spreadsheet.getSheetByName(sheetName);
		if (isNullish(sheet)) {
			throw createSheetNotFoundError(sheetName);
		}
		sheet.getRange(rowNo, columnNo).setValue(value);
	};

	/**
	 * 指定したシートの最終行数を取得する。
	 *
	 * @param sheetName - シートの名称
	 * @returns num
	 */
	getLastRow = (sheetName: string) => {
		const sheet = this.spreadsheet.getSheetByName(sheetName);
		if (isNullish(sheet)) {
			throw createSheetNotFoundError(sheetName);
		}
		const lastRow = sheet.getLastRow();
		return lastRow;
	};

	/**
	 * 指定したシートの最終列数を取得する。
	 *
	 * @param sheetName - シートの名称
	 * @returns num
	 */
	getLastColumn = (sheetName: string) => {
		const sheet = this.spreadsheet.getSheetByName(sheetName);
		if (isNullish(sheet)) {
			throw createSheetNotFoundError(sheetName);
		}
		const lastColumn = sheet.getLastColumn();
		return lastColumn;
	};

	/**
	 * 指定したシートの重複した行を削除する。
	 *
	 * @param sheetName - シートの名称
	 * @param indexes - 重複をチェックする列番号の配列
	 * @param startRowIndex - 重複チェック開始セルの行番号
	 * @param startColumnIndex - 重複チェック開始セルの列番号
	 * @param rowCount - 重複チェックする行数
	 * @param columnCount - 重複チェックする列数
	 * @returns void
	 */
	removeDuplicates = (
		sheetName: string,
		indexes: number[],
		startRowIndex: number,
		startColumnIndex: number,
		rowCount: number,
		columnCount: number,
	): void => {
		const sheet = this.spreadsheet.getSheetByName(sheetName);
		if (isNullish(sheet)) {
			throw createSheetNotFoundError(sheetName);
		}
		sheet
			.getRange(startRowIndex, startColumnIndex, rowCount, columnCount)
			.removeDuplicates(indexes);
	};

	/**
	 * 指定したシートのセル範囲にチェックボックスを設置する。
	 *
	 * @param sheetName - シートの名称
	 * @param startRowIndex - 開始セルの行番号
	 * @param startColumnIndex - 開始セルの列番号
	 * @param rowCount - 設置する行数
	 * @param columnCount - 設置する列数
	 * @returns void
	 */
	setCheckbox = (
		sheetName: string,
		startRowIndex: number,
		startColumnIndex: number,
		rowCount: number,
		columnCount: number,
	): void => {
		const sheet = this.spreadsheet.getSheetByName(sheetName);
		if (isNullish(sheet)) {
			throw createSheetNotFoundError(sheetName);
		}
		sheet
			.getRange(startRowIndex, startColumnIndex, rowCount, columnCount)
			.insertCheckboxes();
	};

	/**
	 * 指定したシートのセル範囲を指定の背景色にする。
	 *
	 * @param sheetName - シートの名称
	 * @param startRowIndex - 開始セルの行番号
	 * @param startColumnIndex - 開始セルの列番号
	 * @param rowCount - 設定する行数
	 * @param columnCount - 設定する列数
	 * @param colorCode - 設定する色コード #ffffff
	 * @returns void
	 */
	setBackgroundColor = (
		sheetName: string,
		startRowIndex: number,
		startColumnIndex: number,
		rowCount: number,
		columnCount: number,
		colorCode: string | null,
	): void => {
		const sheet = this.spreadsheet.getSheetByName(sheetName);
		if (isNullish(sheet)) {
			throw createSheetNotFoundError(sheetName);
		}
		sheet
			.getRange(startRowIndex, startColumnIndex, rowCount, columnCount)
			.setBackground(colorCode);
	};

	/**
	 * 指定したシートのセル範囲にプルダウン設定をする。
	 *
	 * @param sheetName - シートの名称
	 * @param startRowIndex - 開始セルの行番号
	 * @param startColumnIndex - 開始セルの列番号
	 * @param rowCount - 設定する行数
	 * @param columnCount - 設定する列数
	 * @param itemsRange - 設定する選択肢の範囲
	 * @returns void
	 */
	setPulldownRule = (
		sheetName: string,
		startRowIndex: number,
		startColumnIndex: number,
		rowCount: number,
		columnCount: number,
		itemsRange: GoogleAppsScript.Spreadsheet.Range,
	): void => {
		const sheet = this.spreadsheet.getSheetByName(sheetName);
		if (isNullish(sheet)) {
			throw createSheetNotFoundError(sheetName);
		}
		const rule = SpreadsheetApp.newDataValidation()
			.requireValueInRange(itemsRange)
			.build();
		sheet
			.getRange(startRowIndex, startColumnIndex, rowCount, columnCount)
			.setDataValidation(rule);
	};
}
