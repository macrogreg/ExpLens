/// <reference types="office-js" />

import { errorTypeMessageString } from "src/util/format_util";
import { ensureSheetActive, findTableByName } from "./excel-util";
import { authorizedFetch } from "./fetch-tools";
import type { Category } from "./lunchtime-types";

const SheetNameCategories = "LM.Categories";
const TableNameCategories = "LM.CategoriesTable";

export interface ExpenseCategory {
    labelId: string;
    description: string;
    isGroup: boolean;
    isIncome: boolean;
    isExcludeFromBudget: boolean;
    isExcludeFromTotals: boolean;
    isArchived: boolean;
    createdAtExcelUtc: number;
    updatedAtExcelUtc: number;
    archivedOnExcelUtc: number | "";
    labelL1: string;
    labelL2: string;
    lunchId: number;
    displayOrder: number;
    name: string;
}

const categoryHeaders = [
    "labelId",
    "description",
    "isGroup",
    "isIncome",
    "isExcludeFromBudget",
    "isExcludeFromTotals",
    "isArchived",
    "createdAtExcelUtc",
    "updatedAtExcelUtc",
    "archivedOnExcelUtc",
    "labelL1",
    "labelL2",
    "lunchId",
    "displayOrder",
    "name",
];

function categoryToArray(cat: ExpenseCategory): (string | boolean | number)[] {
    const arr: (string | boolean | number)[] = [];
    for (const h of categoryHeaders) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        arr.push((cat as any)[h]);
    }
    return arr;
}

function timeStrToExcel(datetimeStr: string): number {
    const dt = new Date(datetimeStr);

    if (isNaN(dt.getTime())) {
        throw new Error(`Invalid date string '${datetimeStr}'.`);
    }

    // Excel epoch (1900 system): 1899-12-30
    const excelEpoch = Math.abs(Date.UTC(1899, 11, 30));

    const msecPerDay = 24 * 60 * 60 * 1000;

    // Convert JS ms → Excel days
    return (dt.getTime() + excelEpoch) / msecPerDay;
}

function extractLevelLabel(labelId: string, level: number): string {
    labelId = labelId.trim();
    let d = 0;
    for (let p = 0; p < labelId.length; p++) {
        if (labelId[p] === "/") {
            d++;
            if (d === level) {
                return labelId.substring(0, p);
            }
        }
    }

    return labelId;
}

export async function downloadCategories() {
    await Excel.run(async (context: Excel.RequestContext) => {
        // Find and activate the Categories sheet:
        const categSheet = await ensureSheetActive(SheetNameCategories, context);

        const errorMsgCell = categSheet.getRange("B4");
        errorMsgCell.clear();
        await context.sync();

        try {
            // Fetch Tags from the Cloud:
            const fetchedResponseText = await authorizedFetch("GET", "categories?format=flattened", "get all categories");

            // Parse fetched Tags:
            const fetchedResponseData: { categories: Category[] } = JSON.parse(fetchedResponseText);

            // Build category data:
            const fetchedCats = new Map<number, Category>();
            for (const lmCat of fetchedResponseData.categories) {
                fetchedCats.set(lmCat.id, lmCat);
            }

            const categories: ExpenseCategory[] = [];
            for (const lmCat of fetchedCats.values()) {
                const lookupCat = (id: number | null) => (id === null ? undefined : fetchedCats.get(id));

                let labelId = lmCat.name.trim();
                let parent = lookupCat(lmCat.group_id);
                while (parent) {
                    labelId = `${parent.name.trim()}/${labelId}`;
                    parent = lookupCat(parent.group_id);
                }

                categories.push({
                    labelId,
                    isGroup: lmCat.is_group,
                    isIncome: lmCat.is_income,
                    isExcludeFromBudget: lmCat.exclude_from_budget,
                    isExcludeFromTotals: lmCat.exclude_from_totals,
                    description: lmCat.description ?? "",
                    updatedAtExcelUtc: timeStrToExcel(lmCat.updated_at),
                    isArchived: lmCat.archived,
                    createdAtExcelUtc: timeStrToExcel(lmCat.created_at),
                    archivedOnExcelUtc: lmCat.archived_on === null ? "" : timeStrToExcel(lmCat.archived_on),
                    labelL1: extractLevelLabel(labelId, 1),
                    labelL2: extractLevelLabel(labelId, 2),
                    lunchId: lmCat.id,
                    displayOrder: lmCat.order,
                    name: lmCat.name,
                });
            }

            categories.sort((c1, c2) => c1.displayOrder - c2.displayOrder);

            // Is there an existing Categories table?
            const prevCats = await findTableByName(TableNameCategories, context);

            // There is an existing tags table, but it is not on this sheet:
            if (prevCats && prevCats.sheet.id !== categSheet.id) {
                throw new Error(
                    `Table '${TableNameCategories}' exists on the wrong sheet: ${prevCats.range.address}.` +
                        `\nDon't edit this Categories-sheet. Don't name any objects using prefix 'LM.'`
                );
            }

            // Location of Categ tables:
            const catTableOffs = { row: 7, col: 3 };
            const catLeafsTableOffs = { row: 7, col: 1 };

            // !! Must NOT sync until tables are rebuilt to avoid breaking references {

            context.workbook.application.suspendApiCalculationUntilNextSync();
            context.workbook.application.suspendScreenUpdatingUntilNextSync();

            // Delete everything:
            const prevUsedRange = categSheet.getUsedRange();
            prevUsedRange.clear();
            prevUsedRange.conditionalFormats.clearAll();

            // Print Categ headers:
            categSheet.getRangeByIndexes(catTableOffs.row, catTableOffs.col, 1, categoryHeaders.length).values = [
                categoryHeaders,
            ];

            // Print Categ data:
            const categTableData = categories.map((c) => categoryToArray(c));

            categSheet.getRangeByIndexes(
                catTableOffs.row + 1,
                catTableOffs.col,
                categories.length,
                categoryHeaders.length
            ).values = categTableData;

            // Frame Categ table:
            const categTable = categSheet.tables.add(
                categSheet.getRangeByIndexes(
                    catTableOffs.row,
                    catTableOffs.col,
                    categories.length + 1,
                    categoryHeaders.length
                ),
                true
            );
            categTable.name = TableNameCategories;
            categTable.style = "TableStyleMedium12"; // e.g."TableStyleMedium2", "TableStyleDark1", "TableStyleLight9" ...

            // Set time columns format:
            for (let h = 0; h < categoryHeaders.length; h++) {
                const catName = categoryHeaders[h]!;
                if (catName.endsWith("ExcelUtc")) {
                    categSheet.getRangeByIndexes(
                        catTableOffs.row + 1,
                        catTableOffs.col + h,
                        categories.length,
                        1
                    ).numberFormat = [["yyyy-mm-dd hh:mm"]];
                }
            }

            // } Cleared tables rebuilt. Can sync again.
            await context.sync();

            // Leaf Categories:
            const leafCatsHeadRange = categSheet.getRangeByIndexes(catLeafsTableOffs.row, catLeafsTableOffs.col, 1, 1);
            leafCatsHeadRange.values = [["Assignable Categories"]];
            leafCatsHeadRange.format.fill.color = "#0f9ed5";
            leafCatsHeadRange.format.font.color = "#FFFFFF";

            const leafCatsFormulaRange = categSheet.getRangeByIndexes(
                catLeafsTableOffs.row + 1,
                catLeafsTableOffs.col,
                1,
                1
            );
            leafCatsFormulaRange.load("address");
            leafCatsFormulaRange.formulas = [
                [
                    "=FILTER(\n" +
                        "    LM.CategoriesTable[labelId],\n" +
                        "    (LM.CategoriesTable[isGroup]=FALSE) * (LM.CategoriesTable[isArchived]=FALSE)\n" +
                        ")",
                ],
            ];

            const countLeafs = categories.reduce((n, cat) => (cat.isGroup || cat.isArchived ? n : n + 1), 0);
            const leafCatsBodyRange = categSheet.getRangeByIndexes(
                catLeafsTableOffs.row,
                catLeafsTableOffs.col,
                countLeafs + 1,
                1
            );

            const leafCatsBorderTop = leafCatsBodyRange.format.borders.getItem("EdgeTop");
            leafCatsBorderTop.color = "#0f9ed5";
            leafCatsBorderTop.weight = "Thick";
            leafCatsBorderTop.style = "Continuous";
            const leafCatsBorderLeft = leafCatsBodyRange.format.borders.getItem("EdgeLeft");
            leafCatsBorderLeft.color = "#0f9ed5";
            leafCatsBorderLeft.weight = "Thick";
            leafCatsBorderLeft.style = "Continuous";
            const leafCatsBorderRight = leafCatsBodyRange.format.borders.getItem("EdgeRight");
            leafCatsBorderRight.color = "#0f9ed5";
            leafCatsBorderRight.weight = "Thick";
            leafCatsBorderRight.style = "Continuous";
            const leafCatsBorderBottom = leafCatsBodyRange.format.borders.getItem("EdgeBottom");
            leafCatsBorderBottom.color = "#0f9ed5";
            leafCatsBorderBottom.weight = "Thick";
            leafCatsBorderBottom.style = "Continuous";

            leafCatsBodyRange.format.font.bold = true;
            await context.sync();

            // Counts:
            const countsLabelRange = categSheet.getRangeByIndexes(catTableOffs.row - 1, catTableOffs.col - 1, 1, 1);
            countsLabelRange.values = [["← counts →"]];
            countsLabelRange.format.horizontalAlignment = "Center";
            countsLabelRange.format.font.color = "#074f69";
            countsLabelRange.format.font.bold = true;

            const allCatsCountRange = categSheet.getRangeByIndexes(catTableOffs.row - 1, catTableOffs.col, 1, 1);
            allCatsCountRange.format.horizontalAlignment = "Left";
            allCatsCountRange.format.fill.color = "#f2f2f2";
            allCatsCountRange.format.font.color = "#074f69";
            allCatsCountRange.format.font.bold = true;

            allCatsCountRange.formulas = [[`="  " & COUNTA(${TableNameCategories}[labelId])`]];

            const leafCatsCountRange = categSheet.getRangeByIndexes(catLeafsTableOffs.row - 1, catLeafsTableOffs.col, 1, 1);
            leafCatsCountRange.format.horizontalAlignment = "Left";
            leafCatsCountRange.format.fill.color = "#f2f2f2";
            leafCatsCountRange.format.font.color = "#074f69";
            leafCatsCountRange.format.font.bold = true;

            const leafCatsFormulaLocation = leafCatsFormulaRange.address.split("!")[1];
            leafCatsCountRange.formulas = [[`="  " & COUNTA(${leafCatsFormulaLocation}#)`]];

            await context.sync();

            // Headings:
            categSheet.getRange("B3").values = [["Categories"]];
            categSheet.getRange("B3:C3").style = "Heading 1";

            categSheet.getRange("B5").values = [["Leaf Categories"]];
            categSheet.getRange("B5:B5").style = "Heading 2";

            categSheet.getRange("D5").values = [["All Categories"]];
            categSheet.getRangeByIndexes(4, 3, 1, categoryHeaders.length).style = "Heading 2";
            await context.sync();
        } catch (err) {
            console.error(err);
            errorMsgCell.values = [[`Error: ${errorTypeMessageString(err)}`]];
            errorMsgCell.format.font.color = "#FF0000";
            await context.sync();
            throw err;
        }
    });
}
