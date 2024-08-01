import { hello } from "./example-module";

console.log(hello());

function doPost(
	e: GoogleAppsScript.Events.DoPost,
): GoogleAppsScript.Content.TextOutput {
	return ContentService.createTextOutput("Hello, World!");
}
