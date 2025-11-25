/// <reference types="office-js" />

import { errorTypeMessageString } from "src/util/format_util";
import { findTableByName } from "./excel-util";
import { authorizedFetch } from "./fetch-tools";
import type { Tag } from "./lunchmoney-types";
import { strcmp } from "src/util/string_util";

const SheetNameTags = "LM.Tags";
const TableNameTags = "LM.TagsTable";
const TableNameTagGroups = "LM.TagGroupsTable";

export const TagGroupSeparator = ":";
export const UngroupedTagMoniker = ":Ungrouped";

//
// ToDo: This file needs some refactoring!
//

// export function getTagGroup(tag: string): string | null {
//     const i = tag.indexOf(TagGroupSeparator);
//     return i === -1 ? null : tag.substring(0, i);
// }

export function parseTag(tag: string): {
    group: string;
    value: string;
    isGrp: boolean;
} {
    const sepPos = tag.indexOf(TagGroupSeparator);
    return sepPos < 0
        ? {
              group: UngroupedTagMoniker,
              value: tag,
              isGrp: false,
          }
        : {
              group: tag.substring(0, sepPos),
              value: tag.substring(sepPos + 1),
              isGrp: true,
          };
}

export function createTagGroupHeader(groupName: string | null) {
    return `Tag:${groupName ?? UngroupedTagMoniker}`;
}

export function downloadTags(): Promise<void> {
    return Excel.run(_downloadTags);
}

async function _downloadTags(context: Excel.RequestContext) {
    // Find and activate the Tags sheet:
    let tagsSheet = null;
    try {
        tagsSheet = context.workbook.worksheets.getItem(SheetNameTags);
        tagsSheet.load("name, id");
        await context.sync();
    } catch {
        // Sheet does not exist, create it
        tagsSheet = context.workbook.worksheets.add(SheetNameTags);
        await context.sync();
    }

    // Activate the sheet:
    tagsSheet.activate();
    await context.sync();

    const errorMsgCell = tagsSheet.getRange("B4");
    errorMsgCell.clear();
    await context.sync();

    // Headings:
    tagsSheet.getRange("B3").values = [["Tags"]];
    tagsSheet.getRange("B3:C3").style = "Heading 1";

    tagsSheet.getRange("B5").values = [["LunchMoney Tags"]];
    tagsSheet.getRange("B5:E5").style = "Heading 2";

    tagsSheet.getRange("G5").values = [["Tag Groups"]];
    tagsSheet.getRange("G5:I5").style = "Heading 2";

    await context.sync();

    try {
        // Fetch Tags from the Cloud:
        const responseText = await authorizedFetch("GET", "tags", "get all tags");

        // Parse fetched Tags:
        const tags: Tag[] = JSON.parse(responseText);

        // Is there an existing tags table?
        const existingTagsTable = await findTableByName(TableNameTags, context);
        const existingTagGroupsTable = await findTableByName(TableNameTagGroups, context);

        // There is an existing tags table, but it is not on this sheet:
        if (existingTagsTable && existingTagsTable.sheet.id !== tagsSheet.id) {
            throw new Error(
                `Table '${TableNameTags}' exists on the wrong sheet: ${existingTagsTable.range.address}.` +
                    `\nDon't edit this Tags-sheet. Don't name any objects using prefix 'LM.'`
            );
        }

        if (existingTagGroupsTable && existingTagGroupsTable.sheet.id !== tagsSheet.id) {
            throw new Error(
                `Table '${TableNameTagGroups}' exists on the wrong sheet: ${existingTagGroupsTable.range.address}.` +
                    `\nDon't edit this Tags-sheet. Don't name any objects using prefix 'LM.'`
            );
        }

        // Select or create the table:
        const tagsTable = existingTagsTable
            ? await (async () => {
                  console.log(`Tags table '${TableNameTags}' found at '${existingTagsTable.range.address}'.`);

                  existingTagsTable.table.columns.load("count");
                  existingTagsTable.table.rows.load("count");
                  await context.sync();
                  if (existingTagsTable.table.columns.count !== 4) {
                      throw new Error(
                          `Table '${TableNameTags}' must have 4 columns,` +
                              ` but it actually has ${existingTagsTable.table.columns.count}.`
                      );
                  }
                  return existingTagsTable.table;
              })()
            : await (async () => {
                  const tbl = tagsSheet.tables.add("B8:E9", true);
                  tbl.name = TableNameTags;
                  tbl.style = "TableStyleMedium10"; // e.g."TableStyleMedium2", "TableStyleDark1", "TableStyleLight9" ...
                  tbl.columns.load("count");
                  tbl.rows.load("count");
                  await context.sync();
                  console.log(`New Tags table '${TableNameTags}' created.`);
                  return tbl;
              })();

        // Set header names:
        const headerNames = ["id", "name", "description", "archived"];
        const tagsTableHeaderRange = tagsTable.getHeaderRowRange();
        tagsTableHeaderRange.values = [headerNames];

        // Delete existing rows only if there are any
        console.log(`Tags table had ${tagsTable.rows.count} rows before the refresh.`);
        if (tagsTable.rows.count > 0) {
            tagsTable.rows.deleteRowsAt(0, tagsTable.rows.count - 1);
            tagsTable.getDataBodyRange().values = [["", "", "", ""]];
        }

        await context.sync();

        // Output tags:
        if (tags.length > 0) {
            const rowsToAdd = tags
                .sort((t1, t2) => strcmp(t1.name, t2.name))
                .map((t) => [t.id, t.name, t.description, t.archived]);

            tagsTable.rows.add(0, rowsToAdd);
            await context.sync();
            tagsTable.rows.load("count");
            console.log(`Tags table has ${tagsTable.rows.count} rows after the refresh.`);
        }

        console.log(`Tags table had ${tagsTable.rows.count} rows before the refresh.`);

        // Check if the last row of the tagsTable is empty and delete it if so
        const tagsTableBodyRange = tagsTable.getDataBodyRange();
        tagsTableBodyRange.load(["rowCount", "values"]);
        await context.sync();

        if (tagsTableBodyRange.rowCount > 0) {
            const lastRowIndex = tagsTableBodyRange.rowCount - 1;
            const lastRowValues = tagsTableBodyRange.values && tagsTableBodyRange.values[lastRowIndex];

            if (Array.isArray(lastRowValues)) {
                if (!lastRowValues[0]) {
                    tagsTable.rows.deleteRowsAt(lastRowIndex, 1);
                    await context.sync();
                }
            }
        }

        // Formula to count tags:

        const countTagsPreclearRange = tagsSheet.getRange("B7:K7");
        countTagsPreclearRange.values = [["", "", "", "", "← counts →", "", "", "", "", ""]];
        countTagsPreclearRange.format.fill.clear();
        countTagsPreclearRange.format.font.color = "black";

        const countTagsLabelRange = tagsSheet.getRange("F7");
        countTagsLabelRange.format.horizontalAlignment = "Center";
        countTagsLabelRange.format.font.color = "#7e350e";
        countTagsLabelRange.format.font.bold = true;

        const tagNameCountRange = tagsSheet.getRange("C7");
        tagNameCountRange.format.horizontalAlignment = "Left";
        tagNameCountRange.format.fill.color = "#f2f2f2";
        tagNameCountRange.format.font.color = "#7e350e";
        tagNameCountRange.format.font.bold = true;
        tagNameCountRange.formulas = [[`="  " & COUNTA(${TableNameTags}[name])`]];
        await context.sync();

        // Compute tag groups:
        const tagGroups: Record<string, Set<string>> = {};
        for (const tag of tags) {
            const parsedTag = parseTag(tag.name);
            if (!(parsedTag.group in tagGroups)) {
                tagGroups[parsedTag.group] = new Set<string>();
            }
            tagGroups[parsedTag.group]!.add(parsedTag.value);
        }

        // Select or create the table:
        const tagGroupsTable = existingTagGroupsTable
            ? await (async () => {
                  console.log(
                      `Tag groups table '${TableNameTagGroups}' found at '${existingTagGroupsTable.range.address}'.`
                  );

                  existingTagGroupsTable.table.columns.load("count");
                  existingTagGroupsTable.table.rows.load("count");
                  await context.sync();

                  return existingTagGroupsTable.table;
              })()
            : await (async () => {
                  const tbl = tagsSheet.tables.add("G8:G9", true);
                  tbl.name = TableNameTagGroups;
                  tbl.style = "TableStyleMedium10";
                  tbl.columns.load("count");
                  tbl.rows.load("count");
                  await context.sync();
                  console.log(`New Tags table '${TableNameTagGroups}' created.`);
                  return tbl;
              })();

        console.log(
            `Tag Groups table had ${tagGroupsTable.columns.count} columns and ${tagGroupsTable.rows.count} rows before the refresh.`
        );

        if (tagGroupsTable.rows.count > 0) {
            tagGroupsTable.rows.deleteRowsAt(0, tagGroupsTable.rows.count - 1);
            await context.sync();
        }
        while (tagGroupsTable.columns.count > 1) {
            tagGroupsTable.columns.getItemAt(0).delete();
            tagGroupsTable.columns.load("count");
            await context.sync();
        }

        // Output a row starting at G7 with the sorted keys of tagGroups
        const groupHeads = Object.keys(tagGroups).sort();

        if (groupHeads.length > 0) {
            // Headers:
            tagsSheet.getRangeByIndexes(7, 6, 1, groupHeads.length).values = [groupHeads];
            await context.sync();

            // Output each group's values as a column under its key
            let maxGrTags = 0;
            groupHeads.forEach((gHd, i) => {
                const gTags = Array.from(tagGroups[gHd] ?? []).sort();
                if (gTags.length > 0) {
                    maxGrTags = Math.max(maxGrTags, gTags.length);
                    tagsSheet.getRangeByIndexes(8, 6 + i, gTags.length, 1).values = gTags.map((v) => [v]);
                }
            });

            tagGroupsTable.rows.load("count");
            console.log(
                `Tag Groups table had ${tagGroupsTable.columns.count} columns and ${tagGroupsTable.rows.count} rows before the refresh.`
            );

            // Counts:
            const tagGroupCountRange = tagsSheet.getRangeByIndexes(6, 6, 1, groupHeads.length);
            tagGroupCountRange.format.horizontalAlignment = "Left";
            tagGroupCountRange.format.fill.color = "#f2f2f2";
            tagGroupCountRange.format.font.color = "#7e350e";
            tagGroupCountRange.format.font.bold = true;

            const tagGroupCountRangeInsideBorder = tagGroupCountRange.format.borders.getItem("InsideVertical");
            tagGroupCountRangeInsideBorder.color = "white";
            tagGroupCountRangeInsideBorder.style = "Continuous";
            tagGroupCountRangeInsideBorder.weight = "Thin";

            tagGroupCountRange.formulas = [groupHeads.map((grH) => `="  " & COUNTA(${TableNameTagGroups}[${grH}])`)];
            await context.sync();
        }
    } catch (err) {
        console.error(err);
        errorMsgCell.values = [[`ERR: ${errorTypeMessageString(err)}`]];
        errorMsgCell.format.font.color = "#FF0000";
        await context.sync();
        throw err;
    }
}
