import type * as Lunch from "./lunchmoney-types";
import { timeStrToExcel } from "./excel-util";
import { isNullOrWhitespace } from "src/util/string_util";
import { getTagGroups, getTagValues, TagGroupSeparator, type TagValuesCollection } from "./tags";
import type { SyncContext } from "./sync-driver";
import { IndexedMap } from "./IndexedMap";

export interface Transaction {
    trn: Lunch.Transaction;
    pld: Lunch.PlaidMetadata | null;
    tag: TagValuesCollection;
    grpMoniker: string;
    id: number;
}

export const TagColumnsPlaceholder = "<Tag Groups Columns>";
const TagGroupColumnNamePrefix = `Tags${TagGroupSeparator}`;

export const LunchIdColumnName = "LunchId";

const AccountingWithMinusFormatStr = `_($* #,##0.00_);_($* -#,##0.00_);_($* "-"??_);_(@_)`;

export interface TransactionColumnSpec {
    name: string;
    valueFn: ValueExtractor;
    format: string | null;
}

const transactionColumnsSpecs: TransactionColumnSpec[] = [
    transColumn(LunchIdColumnName, (t) => t.trn.id),

    transColumn("date", (t) => timeStrToExcel(t.trn.date), "yyyy-mm-dd"),
    transColumn("Account", (t) => JJ(t.trn.account_display_name, t.pld?.account_owner)),
    transColumn("payee", (t) => t.trn.payee),
    transColumn("amount", (t) => t.trn.to_base, AccountingWithMinusFormatStr),

    transColumn("Category", (t) => JJ(t.trn.category_group_name, t.trn.category_name)),

    transColumn("Plaid:MerchantCategory", (t) => J(t.pld?.category)),
    transColumn("Plaid:TransactionCategory", (t) => {
        const p = t.pld?.personal_finance_category?.primary?.trim() ?? "";
        const d = t.pld?.personal_finance_category?.detailed?.trim() ?? "";

        let s = d.startsWith(p) ? d.slice(p.length) : d;
        while (s.startsWith("_")) {
            s = s.slice(1);
        }
        return JJ(p, s);
    }),

    transColumn(TagColumnsPlaceholder, (_) => null),

    transColumn("status", (t) => t.trn.status),

    transColumn("recurring_description", (t) => t.trn.recurring_description),

    transColumn("GroupMoniker", (t) => t.grpMoniker),

    transColumn("is_income", (t) => t.trn.is_income),
    transColumn("exclude_from_budget", (t) => t.trn.exclude_from_budget),
    transColumn("exclude_from_totals", (t) => t.trn.exclude_from_totals),

    transColumn("has_children", (t) => t.trn.has_children),
    transColumn("is_group", (t) => t.trn.is_group),
    transColumn("is_pending", (t) => t.trn.is_pending),

    transColumn("display_notes", (t) => t.trn.display_notes),

    transColumn("currency", (t) => t.trn.currency),
    transColumn("to_base", (t) => t.trn.to_base, AccountingWithMinusFormatStr),

    transColumn("category_id", (t) => t.trn.category_id),
    transColumn("category_name", (t) => t.trn.category_name),
    transColumn("category_group_id", (t) => t.trn.category_group_id),
    transColumn("category_group_name", (t) => t.trn.category_group_name),

    transColumn("created_at", (t) => t.trn.created_at),
    transColumn("updated_at", (t) => t.trn.updated_at),

    transColumn("notes", (t) => t.trn.notes),
    transColumn("original_name", (t) => t.trn.original_name),
    transColumn("recurring_id", (t) => t.trn.recurring_id),
    transColumn("recurring_payee", (t) => t.trn.recurring_payee),

    transColumn("recurring_cadence", (t) => t.trn.recurring_cadence),
    transColumn("recurring_granularity", (t) => t.trn.recurring_granularity),
    transColumn("recurring_quantity", (t) => t.trn.recurring_quantity),
    transColumn("recurring_type", (t) => t.trn.recurring_type),
    transColumn("recurring_amount", (t) => t.trn.recurring_amount),
    transColumn("recurring_currency", (t) => t.trn.recurring_currency),
    transColumn("parent_id", (t) => t.trn.parent_id),

    transColumn("group_id", (t) => t.trn.group_id),

    transColumn("asset_id", (t) => t.trn.asset_id),
    transColumn("asset_institution_name", (t) => t.trn.asset_institution_name),
    transColumn("asset_name", (t) => t.trn.asset_name),
    transColumn("asset_display_name", (t) => t.trn.asset_display_name),
    transColumn("asset_status", (t) => t.trn.asset_status),
    transColumn("plaid_account_id", (t) => t.trn.plaid_account_id),
    transColumn("plaid_account_name", (t) => t.trn.plaid_account_name),
    transColumn("plaid_account_mask", (t) => t.trn.plaid_account_mask),
    transColumn("institution_name", (t) => t.trn.institution_name),
    transColumn("plaid_account_display_name", (t) => t.trn.plaid_account_display_name),
    //plaid_metadata,
    transColumn("plaid_category", (t) => t.trn.plaid_category ?? ""),
    transColumn("source", (t) => t.trn.source),
    transColumn("display_name", (t) => t.trn.display_name),

    transColumn("account_display_name", (t) => t.trn.account_display_name),
    transColumn("original_tags", (t) =>
        J(
            t.trn.tags?.map((t) => t.name),
            TagListSeparator
        )
    ),
    transColumn("external_id", (t) => t.trn.external_id),

    transColumn("plaid:account_id", (t) => t.pld?.account_id),
    transColumn("plaid:account_owner", (t) => t.pld?.account_owner),
    transColumn("plaid:amount", (t) => t.pld?.amount),
    transColumn("plaid:authorized_date", (t) => t.pld?.authorized_date),
    transColumn("plaid:authorized_datetime", (t) => t.pld?.authorized_datetime),
    transColumn("plaid:category,l1", (t) => t.pld?.category?.[0]),
    transColumn("plaid:category,l2", (t) => J([t.pld?.category?.[0], t.pld?.category?.[1]])),
    transColumn("plaid:category,l3", (t) => J([t.pld?.category?.[0], t.pld?.category?.[1], t.pld?.category?.[2]])),
    transColumn("plaid:category_id", (t) => t.pld?.category_id),
    transColumn("plaid:check_number", (t) => t.pld?.check_number),
    transColumn("plaid:counterparties.count", (t) => t.pld?.counterparties?.length),
    transColumn("plaid:counterparty#01.confidence_level", (t) => t.pld?.counterparties?.[0]?.confidence_level),
    transColumn("plaid:counterparty#01.entity_id", (t) => t.pld?.counterparties?.[0]?.entity_id),
    transColumn("plaid:counterparty#01.logo_url", (t) => t.pld?.counterparties?.[0]?.logo_url),
    transColumn("plaid:counterparty#01.name", (t) => t.pld?.counterparties?.[0]?.name),
    transColumn("plaid:counterparty#01.phone_number", (t) => t.pld?.counterparties?.[0]?.phone_number),
    transColumn("plaid:counterparty#01.type", (t) => t.pld?.counterparties?.[0]?.type),
    transColumn("plaid:counterparty#01.website", (t) => t.pld?.counterparties?.[0]?.website),
    transColumn("plaid:date", (t) => t.pld?.date),
    transColumn("plaid:datetime", (t) => t.pld?.datetime),
    transColumn("plaid:iso_currency_code", (t) => t.pld?.iso_currency_code),
    transColumn("plaid:location.address", (t) => t.pld?.location?.address),
    transColumn("plaid:location.city", (t) => t.pld?.location?.city),
    transColumn("plaid:location.country", (t) => t.pld?.location?.country),
    transColumn("plaid:location.lat", (t) => t.pld?.location?.lat),
    transColumn("plaid:location.lon", (t) => t.pld?.location?.lon),
    transColumn("plaid:location.postal_code", (t) => t.pld?.location?.postal_code),
    transColumn("plaid:location.region", (t) => t.pld?.location?.region),
    transColumn("plaid:location.store_number", (t) => t.pld?.location?.store_number),
    transColumn("plaid:logo_url", (t) => t.pld?.logo_url),
    transColumn("plaid:merchant_entity_id", (t) => t.pld?.merchant_entity_id),
    transColumn("plaid:merchant_name", (t) => t.pld?.merchant_name),
    transColumn("plaid:name", (t) => t.pld?.name),
    transColumn("plaid:payment_channel", (t) => t.pld?.payment_channel),
    transColumn("plaid:payment_meta.by_order_of", (t) => t.pld?.payment_meta?.by_order_of),
    transColumn("plaid:payment_meta.payee", (t) => t.pld?.payment_meta?.payee),
    transColumn("plaid:payment_meta.payer", (t) => t.pld?.payment_meta?.payer),
    transColumn("plaid:payment_meta.payment_method", (t) => t.pld?.payment_meta?.payment_method),
    transColumn("plaid:payment_meta.payment_processor", (t) => t.pld?.payment_meta?.payment_processor),
    transColumn("plaid:payment_meta.ppd_id", (t) => t.pld?.payment_meta?.ppd_id),
    transColumn("plaid:payment_meta.reason", (t) => t.pld?.payment_meta?.reason),
    transColumn("plaid:payment_meta.reference_number", (t) => t.pld?.payment_meta?.reference_number),
    transColumn("plaid:pending", (t) => t.pld?.pending),
    transColumn("plaid:pending_transaction_id", (t) => t.pld?.pending_transaction_id),
    transColumn(
        "plaid:personal_finance_category.confidence_level",
        (t) => t.pld?.personal_finance_category?.confidence_level
    ),
    transColumn("plaid:personal_finance_category.detailed", (t) => t.pld?.personal_finance_category?.detailed),
    transColumn("plaid:personal_finance_category.primary", (t) => t.pld?.personal_finance_category?.primary),
    transColumn("plaid:personal_finance_category.version", (t) => t.pld?.personal_finance_category?.version),
    transColumn("plaid:personal_finance_category_icon_url", (t) => t.pld?.personal_finance_category_icon_url),
    transColumn("plaid:transaction_code", (t) => t.pld?.transaction_code),
    transColumn("plaid:transaction_id", (t) => t.pld?.transaction_id),
    transColumn("plaid:transaction_type", (t) => t.pld?.transaction_type),
    transColumn("plaid:unofficial_currency_code", (t) => t.pld?.unofficial_currency_code),
    transColumn("plaid:website", (t) => t.pld?.website),
];

type ValueExtractor = (trans: Transaction) => string | boolean | number | null | undefined;

function transColumn(name: string, valueFn: ValueExtractor, format: string | null = null): TransactionColumnSpec {
    return {
        name: name.trim(),
        valueFn,
        format,
    };
}

function transTagColumn(tagGroupName: string): TransactionColumnSpec {
    return {
        name: formatTagGroupColumnHeader(tagGroupName),
        valueFn: (t: Transaction) => getTransactionTagsByGroup(t, tagGroupName),
        format: null,
    };
}

export function createTransactionColumnsSpecs(context: SyncContext): IndexedMap<string, TransactionColumnSpec> {
    const tagColsSpecs = getTagGroups(context.tags.assignable).map((grNm) => transTagColumn(grNm));

    const allColsSpecs = transactionColumnsSpecs.flatMap((col) =>
        col.name === TagColumnsPlaceholder ? tagColsSpecs : col
    );

    const specs = new IndexedMap<string, TransactionColumnSpec>();
    for (const cs of allColsSpecs) {
        specs.tryAdd(cs.name, cs);
    }

    return specs;
}

function getTransactionTagsByGroup(tran: Transaction, tagGroupName: string) {
    const groupTagsList = getTagValues(tran.tag, tagGroupName);
    const tagsStr = J(groupTagsList, TagListSeparator) as string;
    return tagsStr;
}

function formatTagGroupColumnHeader(groupName: string) {
    return `${TagGroupColumnNamePrefix}${groupName}`.trim();
}

export function tryGetTagGroupFromColumnName(columnName: string): string | undefined {
    columnName = columnName.trim();
    if (!columnName.startsWith(TagGroupColumnNamePrefix)) {
        return undefined;
    }

    return columnName.substring(TagGroupColumnNamePrefix.length);
}

export type TransactionRowValue = string | number | boolean | null;

export type TransactionRowData = {
    values: Record<string, TransactionRowValue>;
    range: Excel.Range;
};

export function getTransactionColumnValue(
    tran: Transaction,
    colName: string,
    columnSpecs: IndexedMap<string, TransactionColumnSpec>
): string | boolean | number {
    const colSpec = columnSpecs.getByKey(colName);

    let value;
    if (colSpec !== undefined) {
        value = colSpec.valueFn(tran);
    } else {
        const tagGroupName = tryGetTagGroupFromColumnName(colName);
        if (tagGroupName !== undefined) {
            value = getTransactionTagsByGroup(tran, tagGroupName);
        } else {
            throw new Error(`Cannot find specification for column '${colName}'.`);
        }
    }

    return value === null || value === undefined ? "" : value;
}

const StructureLevelSeparator = " / ";
const TagListSeparator = ", ";

function JJ(v1: string | null | undefined, v2: string | null | undefined, separator: string = StructureLevelSeparator) {
    const r1 = v1 === null || v1 === undefined ? "" : v1;
    const r2 = v2 === null || v2 === undefined ? "" : v2;
    return r1.length > 0 && r2.length > 0 ? r1 + separator + r2 : r1 + r2;
}

function J(vals: (string | null | undefined)[] | null | undefined, separator: string = StructureLevelSeparator) {
    if (vals === null || vals === undefined) {
        return null;
    }
    return vals.map((v) => (isNullOrWhitespace(v) ? "*" : v)).join(separator);
}
