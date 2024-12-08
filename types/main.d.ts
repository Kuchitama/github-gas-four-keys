// types/gas.d.ts
// declare namespace GoogleAppsScript {
//   // export interface Properties {
//   //   getProperty(key: string): string | null;
//   // }

//   // export interface PropertiesService {
//   //   getScriptProperties(): Properties;
//   // }

//   export interface Range {
//     setValues(values: any[][]): Range;
//     setBackgroundRGB(red: number, green: number, blue: number): Range;
//     setNumberFormats(formats: string[][]): Range;
//     copyTo(destination: Range): void;
//     setValue(value: any): Range;
//   }

//   export interface Sheet {
//     getRange(row: number, column: number, numRows?: number, numColumns?: number): Range;
//     setConditionalFormatRules(rules: any[]): void;
//     getConditionalFormatRules(): any[];
//     newChart(): ChartBuilder;
//     insertChart(chart: EmbeddedChart): void;
//   }

//   export interface Spreadsheet {
//     getSheetByName(name: string): Sheet | null;
//     insertSheet(sheetName?: string): Sheet;
//   }

//   export interface SpreadsheetApp {
//     getActiveSpreadsheet(): Spreadsheet;
//     newConditionalFormatRule(): ConditionalFormatRuleBuilder;
//     WeekDay: {
//       MONDAY: number;
//     };
//   }

//   export interface ChartBuilder {
//     addRange(range: Range): ChartBuilder;
//     setChartType(type: string): ChartBuilder;
//     setNumHeaders(headers: number): ChartBuilder;
//     setOption(option: string, value: any): ChartBuilder;
//     setPosition(row: number, column: number, offsetX: number, offsetY: number): ChartBuilder;
//     build(): EmbeddedChart;
//   }

//   export interface EmbeddedChart {}

//   export interface ConditionalFormatRuleBuilder {
//     whenTextEqualTo(text: string): ConditionalFormatRuleBuilder;
//     setBackground(color: string): ConditionalFormatRuleBuilder;
//     setRanges(ranges: Range[]): ConditionalFormatRuleBuilder;
//     build(): any;
//   }

//   export interface URLFetchApp {
//     fetch(url: string, params?: URLFetchRequestOptions): HTTPResponse;
//   }

//   export interface HTTPResponse {
//     getContentText(): string;
//   }

//   export interface URLFetchRequestOptions {
//     method?: string;
//     contentType?: string;
//     headers?: { [key: string]: string };
//     payload?: string;
//   }

//   export interface ScriptApp {
//     getProjectTriggers(): Trigger[];
//     newTrigger(functionName: string): TriggerBuilder;
//     WeekDay: {
//       MONDAY: number;
//     };
//   }

//   export interface Trigger {
//     getHandlerFunction(): string;
//   }

//   export interface TriggerBuilder {
//     timeBased(): TriggerBuilder;
//     onWeekDay(day: number): TriggerBuilder;
//     atHour(hour: number): TriggerBuilder;
//     create(): Trigger;
//   }
// }

// types/main.d.ts
interface PullRequest {
  author: {
    login: string;
  };
  headRefName: string;
  bodyText: string;
  merged: boolean;
  mergedAt: string | null;
  commits: {
    nodes: Array<{
      commit: {
        committedDate: string;
      };
    }>;
  };
}

export { PullRequest };
