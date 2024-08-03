/**
 * Copyright 2024 k.k.Factory
 */
import { SpreadsheetUtil } from "@/common/spreadSheetUtils";

function doPost(
	e: GoogleAppsScript.Events.DoPost,
): GoogleAppsScript.Content.TextOutput {
	// POSTリクエストのペイロードを取得
	const payload = e.postData.contents;

	// ペイロードを解析（必要に応じて）
	const data = JSON.parse(payload);
	console.log("=======data: ", data);
	const spreadsheetUtil = new SpreadsheetUtil();
	const a1cell = spreadsheetUtil.getCellValue("Sheet1", "A1");
	console.log("a1cell", a1cell);
	return ContentService.createTextOutput(`Received payload: ${payload}`);
}

function doGet(
	e: GoogleAppsScript.Events.DoPost,
): GoogleAppsScript.Content.TextOutput {
	// POSTリクエストのペイロードを取得
	const payload = e.postData.contents;

	// ペイロードを解析（必要に応じて）
	const data = JSON.parse(payload);
	console.log("=======data: ", data);
	const spreadsheetUtil = new SpreadsheetUtil();
	const a1cell = spreadsheetUtil.getCellValue("Sheet1", "A1");
	console.log("a1cell", a1cell);
	return ContentService.createTextOutput(`Received payload: ${a1cell}`);
}
