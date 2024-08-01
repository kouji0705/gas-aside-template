import { SpreadsheetUtil } from "@/common/spreadSheetUtils";

function doPost(
	e: GoogleAppsScript.Events.DoPost,
): GoogleAppsScript.Content.TextOutput {
	// POSTリクエストのペイロードを取得
	const payload = e.postData.contents;

	// ペイロードを解析（必要に応じて）
	const data = JSON.parse(payload);

	const spreadsheetUtil = new SpreadsheetUtil();
	const a1cell = spreadsheetUtil.getCellValue("Sheet1", "A1");
	console.log('=======HIT8 ', a1cell);
	return ContentService.createTextOutput(`Received payload: ${payload}`);
}
