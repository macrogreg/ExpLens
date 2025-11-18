///
/// See https://lunchmoney.dev/
///

export interface Tag {
    id: number;
    name: string;
    description: string;
    archived: boolean;
}

export interface Category {
    id: number;
    name: string;
    description: string | null;
    is_income: boolean;
    exclude_from_budget: boolean;
    exclude_from_totals: boolean;
    updated_at: string;
    created_at: string;
    is_group: boolean;
    group_id: number | null;
    archived: boolean;
    archived_on: string | null;
    order: number;
    children?: Category[];
}
