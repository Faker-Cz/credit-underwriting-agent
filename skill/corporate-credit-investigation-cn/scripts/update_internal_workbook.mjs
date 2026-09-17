import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = process.argv[2];
const previewPath = process.argv[3];
if (!inputPath || !previewPath) {
  throw new Error("usage: node update_internal_workbook.mjs <workbook.xlsx> <preview.png>");
}

const input = await FileBlob.load(inputPath);
const workbook = await SpreadsheetFile.importXlsx(input);
const sheet = workbook.worksheets.getItemAt(0);

sheet.getRange("F16").values = [["其他银行可使用的授信额度、发行债券、融资租赁、股东借款、吸收直接投资、发行股票、商业信用、财政拨款、非银行金融机构等可用于日常经营生产的其他融资，无则填0"]];
sheet.getRange("B23").values = [["1、本表黄色区域为需录入数据或由调查报告中财务数据引入（未锁定）；蓝色区域、红色区域分别是中间和最后结果（已锁定）；2、本表仅适用于存续两年以上的生产或销售型单位；对不足两年的客户及其他类型的客户，应根据订单情况、其他资金来源、行业平均营运资金周转次数等合理确定贷款需求及期限；3、对集团关联客户，可采用合并报表和单独的子公司分别估算流动资金贷款额度，原则上纳入合并报表范围内的成员企业流动资金贷款总和不能超过估算值；4、本表是基于真实且理想化的财务报表的测算结果，供参考。"]];
sheet.getRange("E17").dataValidation = {
  rule: { type: "decimal", operator: "between", formula1: 1, formula2: 2 },
  prompt: { showPrompt: true, title: "调节系数", message: "请输入1至2之间的数值，一般取1。" },
  errorAlert: { showAlert: true, style: "stop", title: "输入超出范围", message: "调节系数必须在1至2之间。" },
  ignoreBlanks: true,
};

const inspection = await workbook.inspect({ kind: "region,formula", sheetId: sheet.name, range: "A1:H23", maxChars: 5000 });
const errors = await workbook.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A", options: { useRegex: true, maxResults: 50 }, maxChars: 3000 });
console.log(inspection.ndjson);
console.log(errors.ndjson);

await fs.mkdir(path.dirname(previewPath), { recursive: true });
const preview = await workbook.render({ sheetName: sheet.name, range: "A1:H23", scale: 1.5, format: "png" });
await fs.writeFile(previewPath, new Uint8Array(await preview.arrayBuffer()));
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(inputPath);
