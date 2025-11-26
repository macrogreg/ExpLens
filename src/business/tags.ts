/// <reference types="office-js" />

import { errorTypeMessageString } from "src/util/format_util";
import { findTableByNameOnSheet, getRangeBasedOn, parseOnSheetAddress } from "./excel-util";
import { authorizedFetch } from "./fetch-tools";
import type { Tag } from "./lunchmoney-types";
import { strcmp } from "src/util/string_util";
import type { SyncContext } from "./sync-driver";

export const SheetNameTags = "LM.Tags";
const TableNameTags = "LM.TagsTable";
const TableNameTagGroups = "LM.TagGroupsTable";

export const TagGroupSeparator = ":";
const UngroupedTagMoniker = ":Ungrouped";

export type TagValuesCollection = Map<string, Set<string>>;

export function getTagGroups(tagCollection: TagValuesCollection): string[] {
    return [...tagCollection.keys()].sort();
}

export function getTagValues(tagCollection: TagValuesCollection, tagGroupName: string): string[] {
    const groupTagsSet = tagCollection.get(tagGroupName);
    const groupTagsList = groupTagsSet === undefined ? [] : [...groupTagsSet].sort();
    return groupTagsList;
}

export function addTagValue(tagCollection: TagValuesCollection, tag: TagInfo): TagValuesCollection {
    let groupTagsSet = tagCollection.get(tag.group);

    if (groupTagsSet === undefined) {
        groupTagsSet = new Set<string>();
        tagCollection.set(tag.group, groupTagsSet);
    }

    groupTagsSet.add(tag.value);
    return tagCollection;
}

export type TagInfo = {
    group: string;
    value: string;
    isGrp: boolean;
    name: string;
};

export function parseTag(tagName: string): TagInfo {
    const sepPos = tagName.indexOf(TagGroupSeparator);
    return sepPos < 0
        ? {
              group: UngroupedTagMoniker,
              value: tagName,
              isGrp: false,
              name: tagName.trim(),
          }
        : {
              group: tagName.substring(0, sepPos),
              value: tagName.substring(sepPos + 1),
              isGrp: true,
              name: tagName.trim(),
          };
}

export function createTagGroupHeader(groupName: string | null) {
    return `Tag:${groupName ?? UngroupedTagMoniker}`;
}

export async function downloadTags(context: SyncContext) {
    // Activate the sheet:
    context.sheets.tags.activate();
    await context.excel.sync();

    // Clear area for the error messages:
    const errorMsgCell = context.sheets.tags.getRange("B4");
    errorMsgCell.clear();
    await context.excel.sync();

    try {
        // Tags table spec:
        const tagsTableOffs = { row: 7, col: 1 };
        const tagsTableHead_Grp = "Structured:Group";
        const tagsTableHead_Val = "Structured:Value";
        const tagsTableHeaderNames = ["id", "name", "description", "archived", tagsTableHead_Grp, tagsTableHead_Val];

        // Sheet header:
        context.sheets.tags.getRange("B3").values = [["Tags"]];
        context.sheets.tags.getRange("B3:C3").style = "Heading 1";
        await context.excel.sync();

        // Fetch Tags from the Cloud:
        const fetchedResponseText = await authorizedFetch("GET", "tags", "get all tags");

        // Parse fetched Tags:
        const parsedTags: Tag[] = JSON.parse(fetchedResponseText);

        // Preliminary table area header (final version after autofit, so that the header does not stretch its column):
        getRangeBasedOn(context.sheets.tags, tagsTableOffs, -3, 0, 1, 1).values = [["Tags"]];
        getRangeBasedOn(context.sheets.tags, tagsTableOffs, -3, 0, 1, tagsTableHeaderNames.length).style = "Heading 2";
        await context.excel.sync();

        // Is there an existing tags table? (this will throw if table is not on the Tags sheet)
        const existingTagsTableInfo = await findTableByNameOnSheet(TableNameTags, context.sheets.tags, context.excel);

        // Make sure the table exists:
        const tagsTable =
            existingTagsTableInfo !== null
                ? existingTagsTableInfo.table
                : await (async () => {
                      const tbl = context.sheets.tags.tables.add(
                          getRangeBasedOn(context.sheets.tags, tagsTableOffs, 0, 0, 2, 1),
                          true
                      );
                      tbl.name = TableNameTags;
                      tbl.style = "TableStyleMedium10"; // e.g."TableStyleMedium2", "TableStyleDark1", "TableStyleLight9" ...
                      await context.excel.sync();
                      console.debug(`New Tags table '${TableNameTags}' created.`);
                      return tbl;
                  })();

        tagsTable.columns.load("count");
        tagsTable.rows.load("count");
        await context.excel.sync();

        // Make sure the table has the right columns:
        while (tagsTable.columns.count < tagsTableHeaderNames.length) {
            tagsTable.columns.add();
            tagsTable.columns.load("count");
            await context.excel.sync();
        }
        while (tagsTable.columns.count > tagsTableHeaderNames.length) {
            tagsTable.columns.getItemAt(0).delete();
            tagsTable.columns.load("count");
            await context.excel.sync();
        }

        tagsTable.getHeaderRowRange().values = [tagsTableHeaderNames];
        tagsTable.columns.load("count");
        await context.excel.sync();

        // Clear all rows:
        if (tagsTable.rows.count > 1) {
            tagsTable.rows.deleteRowsAt(0, tagsTable.rows.count - 1);
        }

        // Top row cannot be deleted (tables must not have zero data rows), so fill with "":
        tagsTable.getDataBodyRange().values = [Array(tagsTableHeaderNames.length).fill("")];

        tagsTable.rows.load("count");
        await context.excel.sync();

        {
            const tagsTableRange = tagsTable.getRange();
            tagsTableRange.load("address");
            await context.excel.sync();
            console.debug(
                `Tags table "${TableNameTags}" is located at "${tagsTableRange.address}",` +
                    ` it has ${tagsTable.columns.count} columns and ${tagsTable.rows.count} row(s).`
            );
        }

        // Add tag count field:
        const countTagsLabelRange = getRangeBasedOn(context.sheets.tags, tagsTableOffs, -1, 0, 1, 1);
        countTagsLabelRange.format.fill.clear();
        countTagsLabelRange.format.font.color = "#7e350e";
        countTagsLabelRange.format.font.bold = true;
        countTagsLabelRange.format.horizontalAlignment = "Right";
        countTagsLabelRange.values = [["Count:"]];

        const countTagsFormulaRange = getRangeBasedOn(context.sheets.tags, tagsTableOffs, -1, 1, 1, 1);
        countTagsFormulaRange.format.fill.color = "#f2f2f2";
        countTagsFormulaRange.format.font.color = "#7e350e";
        countTagsFormulaRange.format.font.bold = true;
        countTagsFormulaRange.format.horizontalAlignment = "Left";
        countTagsFormulaRange.formulas = [[`="  " & COUNTA(${TableNameTags}[name])`]];
        await context.excel.sync();

        // Process fetched tags and populate table:
        context.tags.assignable.clear();
        context.tags.groupListFormulaLocations.clear();
        context.tags.byId.clear();

        const sortedParsedTags = parsedTags.sort((t1, t2) => strcmp(t1.name, t2.name));
        for (let t = 0; t < sortedParsedTags.length; t++) {
            const parsedTag = sortedParsedTags[t]!;
            const tagInfo = parseTag(parsedTag.name);

            addTagValue(context.tags.assignable, tagInfo);
            context.tags.byId.set(parsedTag.id, tagInfo);

            const rowToAdd = [
                parsedTag.id,
                parsedTag.name,
                parsedTag.description,
                parsedTag.archived,
                tagInfo.group,
                tagInfo.value,
            ];

            if (t === 0) {
                tagsTable.rows.getItemAt(0).getRange().values = [rowToAdd];
            } else {
                tagsTable.rows.add(-1, [rowToAdd]);
            }

            tagsTable.rows.load("count");
            await context.excel.sync();
        }

        tagsTable.getRange().format.autofitColumns();
        await context.excel.sync();

        // Reprint table area header after the autofit, so that the header does not stretch its column:
        getRangeBasedOn(context.sheets.tags, tagsTableOffs, -3, 0, 1, 1).values = [["LunchMoney Tags"]];
        await context.excel.sync();

        console.log(`Tags table has ${tagsTable.rows.count} rows after the refresh.`);

        // Now, we build the Tag Groups Table.

        const tagGroupsTableOffs = {
            row: tagsTableOffs.row,
            col: tagsTableOffs.col + tagsTableHeaderNames.length + 1,
        };

        // Print Tag Groups table area header:
        getRangeBasedOn(context.sheets.tags, tagGroupsTableOffs, -3, 0, 1, 1).values = [["Tag Groups"]];
        getRangeBasedOn(context.sheets.tags, tagGroupsTableOffs, -3, 0, 1, context.tags.assignable.size).style =
            "Heading 2";
        await context.excel.sync();

        // Find and delete Tag Groups table:
        const existingTagGroupsTable = await findTableByNameOnSheet(
            TableNameTagGroups,
            context.sheets.tags,
            context.excel
        );

        if (existingTagGroupsTable !== null) {
            existingTagGroupsTable.table.delete();
            await context.excel.sync();
        }

        // Print data about the Tag Groups:
        const tagGroupNames = getTagGroups(context.tags.assignable);

        for (let g = 0; g < tagGroupNames.length; g++) {
            const groupName = tagGroupNames[g]!;

            const groupNameRange = getRangeBasedOn(context.sheets.tags, tagGroupsTableOffs, 0, g, 1, 1);
            const groupCountRange = getRangeBasedOn(context.sheets.tags, tagGroupsTableOffs, 1, g, 1, 1);
            const groupGapRange = getRangeBasedOn(context.sheets.tags, tagGroupsTableOffs, 2, g, 1, 1);
            const groupListRange = getRangeBasedOn(context.sheets.tags, tagGroupsTableOffs, 3, g, 1, 1);

            const grCnt = getTagValues(context.tags.assignable, groupName).length;
            const groupListBackRange = getRangeBasedOn(context.sheets.tags, tagGroupsTableOffs, 3, g, grCnt, 1);

            groupNameRange.load(["address"]);
            groupCountRange.load(["address"]);
            groupListRange.load(["address"]);
            await context.excel.sync();

            const groupNameRangeAddr = parseOnSheetAddress(groupNameRange.address);
            context.tags.groupListFormulaLocations.set(groupName, groupListRange.address);
            // console.debug(
            //     `Tag Group #${g}: groupNameRangeAddr='${groupNameRangeAddr}';` +
            //         ` groupListRange.address='${groupListRange.address}'.`
            // );

            groupNameRange.formulas = [[""]];
            groupNameRange.values = [[groupName]];

            groupGapRange.values = [[""]];
            groupGapRange.formulas = [[""]];
            groupGapRange.format.fill.clear();

            groupListBackRange.values = Array(grCnt).fill([""]);
            groupListBackRange.formulas = Array(grCnt).fill([""]);
            groupListBackRange.format.fill.color = "#f2f2f2";
            groupListBackRange.format.font.color = "#7e350e";
            groupListBackRange.format.font.bold = false;
            groupListBackRange.format.horizontalAlignment = "Left";
            const groupListBackRangeLeftBorder = groupListBackRange.format.borders.getItem("EdgeLeft");
            groupListBackRangeLeftBorder.color = "white";
            groupListBackRangeLeftBorder.style = "Continuous";
            groupListBackRangeLeftBorder.weight = "Medium";

            groupListRange.values = [[""]];
            groupListRange.formulas = [
                [
                    `= FILTER(${TableNameTags}[${tagsTableHead_Val}], ${TableNameTags}[${tagsTableHead_Grp}]=${groupNameRangeAddr})`,
                ],
            ];

            groupCountRange.values = [[""]];
            groupCountRange.formulas = [[`= COUNTA(${groupListRange.address}#)`]];

            groupNameRange.format.autofitColumns();
            groupListRange.format.autofitColumns();
            await context.excel.sync();
        }

        // Frame the printed Tag Group Item Count data with a table:
        const tagGroupsTable = context.sheets.tags.tables.add(
            getRangeBasedOn(context.sheets.tags, tagGroupsTableOffs, 0, 0, 2, tagGroupNames.length),
            true
        );
        tagGroupsTable.name = TableNameTagGroups;
        tagGroupsTable.style = "TableStyleMedium10"; // e.g."TableStyleMedium2", "TableStyleDark1", "TableStyleLight9" ...
        await context.excel.sync();
        console.debug(`New Tags Groups table '${TableNameTagGroups}' created.`);
    } catch (err) {
        console.error(err);
        errorMsgCell.values = [[`ERR: ${errorTypeMessageString(err)}`]];
        errorMsgCell.format.font.color = "#FF0000";
        await context.excel.sync();
        throw err;
    }
}
