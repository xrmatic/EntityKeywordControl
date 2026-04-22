import * as React from "react";
import {
    FluentProvider,
    webLightTheme,
    TagGroup,
    makeStyles,
    tokens,
    Spinner,
    MessageBar,
    MessageBarBody,
    Button,
    Text,
    Tooltip,
} from "@fluentui/react-components";
import { AddRegular, TagRegular } from "@fluentui/react-icons";
import { KeywordTag, OverflowBadge } from "./KeywordTag";
import { KeywordPickerPanel } from "./KeywordPickerPanel";
import { IKeywordRecord, IEntityKeywordControlProps } from "./types";

// ─── Styles ───────────────────────────────────────────────────────────────────

const useStyles = makeStyles({
    root: {
        display: "flex",
        flexDirection: "column",
        gap: tokens.spacingVerticalS,
        fontFamily: tokens.fontFamilyBase,
        minHeight: "32px",
    },
    tagRow: {
        display: "flex",
        flexWrap: "wrap",
        alignItems: "center",
        gap: tokens.spacingHorizontalXS,
    },
    actions: {
        display: "flex",
        gap: tokens.spacingHorizontalXS,
        alignItems: "center",
    },
    emptyLabel: {
        color: tokens.colorNeutralForeground4,
        fontStyle: "italic",
    },
    loadingWrapper: {
        display: "flex",
        alignItems: "center",
        gap: tokens.spacingHorizontalS,
    },
});

// ─── Helper – read a string param ────────────────────────────────────────────

function param(p: ComponentFramework.PropertyTypes.StringProperty | undefined): string {
    return p?.raw ?? "";
}

function boolParam(p: ComponentFramework.PropertyTypes.TwoOptionsProperty | undefined, defaultValue = true): boolean {
    if (p?.raw === undefined || p?.raw === null) return defaultValue;
    return Boolean(p.raw);
}

function numParam(p: ComponentFramework.PropertyTypes.WholeNumberProperty | undefined, defaultValue = 0): number {
    if (p?.raw === undefined || p?.raw === null) return defaultValue;
    return Number(p.raw);
}

// ─── WebAPI helpers ───────────────────────────────────────────────────────────

/**
 * Builds the OData $select string from the configured field names.
 */
function buildSelectClause(
    labelField: string,
    colorField: string,
    iconField: string,
    groupField: string,
    descField: string,
    sortField: string
): string {
    const fields = [labelField, colorField, iconField, groupField, descField, sortField]
        .filter(Boolean);
    return fields.length ? `$select=${fields.join(",")}` : "";
}

/**
 * Maps a raw Dataverse entity record to our IKeywordRecord interface.
 */
function mapRecord(
    raw: ComponentFramework.WebApi.Entity,
    labelField: string,
    colorField: string,
    iconField: string,
    groupField: string,
    descField: string,
    sortField: string
): IKeywordRecord {
    const idKey = Object.keys(raw).find((k) => k.endsWith("id") && !k.startsWith("@"));
    return {
        id: (idKey ? raw[idKey] : raw["@odata.id"]) as string,
        label: (raw[labelField] as string) ?? "(unnamed)",
        color: colorField ? (raw[colorField] as string | undefined) : undefined,
        icon: iconField ? (raw[iconField] as string | undefined) : undefined,
        group: groupField ? (raw[groupField] as string | undefined) : undefined,
        description: descField ? (raw[descField] as string | undefined) : undefined,
        sortOrder: sortField ? (raw[sortField] as number | undefined) : undefined,
    };
}

/**
 * Sorts an array of IKeywordRecord by sortOrder (ascending), then by label.
 */
function sortKeywords(kws: IKeywordRecord[]): IKeywordRecord[] {
    return [...kws].sort((a, b) => {
        const oa = a.sortOrder ?? Number.MAX_SAFE_INTEGER;
        const ob = b.sortOrder ?? Number.MAX_SAFE_INTEGER;
        if (oa !== ob) return oa - ob;
        return a.label.localeCompare(b.label);
    });
}

// ─── Component ────────────────────────────────────────────────────────────────

export const EntityKeywordControlComponent: React.FC<IEntityKeywordControlProps> = ({
    context,
}) => {
    const styles = useStyles();

    // ── Read manifest parameters ──────────────────────────────────────────────
    const p = context.parameters;
    const relationshipName     = param(p.relationshipName);
    const relatedEntityName    = param(p.relatedEntityName);
    const relatedEntitySetName = param(p.relatedEntitySetName);
    const navPropertyName      = param(p.navigationPropertyName);
    const labelField           = param(p.labelFieldName);
    const colorField           = param(p.colorFieldName);
    const iconField            = param(p.iconFieldName);
    const groupField           = param(p.groupFieldName);
    const descField            = param(p.descriptionFieldName);
    const sortField            = param(p.sortOrderFieldName);
    const allowAssociate       = boolParam(p.allowAssociate, true);
    const allowDisassociate    = boolParam(p.allowDisassociate, true);
    const allowCreate          = boolParam(p.allowCreate, false);
    const searchPlaceholder    = param(p.searchPlaceholder) || "Search keywords…";
    const maxVisible           = numParam(p.maxVisibleTags, 0);
    const isControlDisabled    = context.mode.isControlDisabled;

    // ── State ─────────────────────────────────────────────────────────────────
    const [selectedKeywords, setSelectedKeywords] = React.useState<IKeywordRecord[]>([]);
    const [allKeywords, setAllKeywords]           = React.useState<IKeywordRecord[]>([]);
    const [isLoading, setIsLoading]               = React.useState(true);
    const [error, setError]                       = React.useState<string | null>(null);
    const [isPanelOpen, setIsPanelOpen]           = React.useState(false);
    const [showAll, setShowAll]                   = React.useState(false);

    // ── Derive current record identity ────────────────────────────────────────
    // context.page is available at runtime in model-driven apps but is not
    // declared in the published @types/powerapps-component-framework typings.
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const pageCtx = (context as any).page as { entityId?: string; entityTypeName?: string } | undefined;
    const entityId   = pageCtx?.entityId   ?? "";
    const entityName = pageCtx?.entityTypeName ?? "";

    // associateRecord / disassociateRecord are available at runtime but absent
    // from the current version of the published @types/powerapps-component-framework
    // typings – hold a typed reference to avoid spreading `any` everywhere.
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const webAPI = context.webAPI as any;

    // ── Load data ─────────────────────────────────────────────────────────────

    const selectClause = React.useMemo(
        () => buildSelectClause(labelField, colorField, iconField, groupField, descField, sortField),
        [labelField, colorField, iconField, groupField, descField, sortField]
    );

    /**
     * Loads all available keyword records from the related entity set.
     */
    const loadAllKeywords = React.useCallback(async () => {
        if (!relatedEntityName || !labelField) return;
        try {
            const query = selectClause ? `?${selectClause}` : undefined;
            const result = await webAPI.retrieveMultipleRecords(
                relatedEntityName,
                query
            );
            const mapped = result.entities.map((e: ComponentFramework.WebApi.Entity) =>
                mapRecord(e, labelField, colorField, iconField, groupField, descField, sortField)
            );
            setAllKeywords(sortKeywords(mapped));
        } catch (e) {
            setError(`Failed to load keywords: ${(e as Error).message ?? String(e)}`);
        }
    }, [webAPI, relatedEntityName, labelField, selectClause, colorField, iconField, groupField, descField, sortField]);

    /**
     * Loads keywords currently associated with the host record by expanding
     * the navigation property.
     */
    const loadSelectedKeywords = React.useCallback(async () => {
        if (!entityId || !entityName || !navPropertyName || !labelField) return;
        try {
            const expandClause = selectClause
                ? `$expand=${navPropertyName}(${selectClause})`
                : `$expand=${navPropertyName}`;
            const record = await webAPI.retrieveRecord(
                entityName,
                entityId,
                `?${expandClause}`
            );
            const relatedArray: ComponentFramework.WebApi.Entity[] =
                (record[navPropertyName] as ComponentFramework.WebApi.Entity[]) ?? [];
            const mapped = relatedArray.map((e) =>
                mapRecord(e, labelField, colorField, iconField, groupField, descField, sortField)
            );
            setSelectedKeywords(sortKeywords(mapped));
        } catch (e) {
            setError(`Failed to load associated keywords: ${(e as Error).message ?? String(e)}`);
        }
    }, [webAPI, entityId, entityName, navPropertyName, labelField, selectClause, colorField, iconField, groupField, descField, sortField]);

    React.useEffect(() => {
        let mounted = true;
        setIsLoading(true);
        setError(null);

        Promise.all([loadAllKeywords(), loadSelectedKeywords()])
            .catch((e) => {
                if (mounted) setError(String(e));
            })
            .finally(() => {
                if (mounted) setIsLoading(false);
            });

        return () => { mounted = false; };
    }, [loadAllKeywords, loadSelectedKeywords]);

    // ── Disassociate handler ──────────────────────────────────────────────────

    const handleDismiss = React.useCallback(
        async (keyword: IKeywordRecord) => {
            if (!entityId || !entityName || !relationshipName) return;
            try {
                await webAPI.disassociateRecord(
                    entityName,
                    entityId,
                    relationshipName,
                    keyword.id
                );
                setSelectedKeywords((prev) => prev.filter((k) => k.id !== keyword.id));
            } catch (e) {
                setError(`Failed to remove keyword: ${(e as Error).message ?? String(e)}`);
            }
        },
        [webAPI, entityId, entityName, relationshipName]
    );

    // ── Associate handler ─────────────────────────────────────────────────────

    const handleApply = React.useCallback(
        async (toAssociate: IKeywordRecord[]) => {
            setIsPanelOpen(false);
            if (!entityId || !entityName || !relationshipName || !relatedEntitySetName) return;
            const errors: string[] = [];
            for (const kw of toAssociate) {
                // Synthetic "new" records from the Create flow
                if (kw.id.startsWith("__new__::")) {
                    // Future: open a quick-create form for kw.label
                    continue;
                }
                try {
                    await webAPI.associateRecord(
                        entityName,
                        entityId,
                        relationshipName,
                        relatedEntitySetName,
                        kw.id
                    );
                } catch (e) {
                    errors.push(`${kw.label}: ${(e as Error).message ?? String(e)}`);
                }
            }
            if (errors.length) {
                setError(`Some keywords could not be added:\n${errors.join("\n")}`);
            }
            // Refresh selected list
            await loadSelectedKeywords();
        },
        [webAPI, entityId, entityName, relationshipName, relatedEntitySetName, loadSelectedKeywords]
    );

    // ── Derived display values ────────────────────────────────────────────────

    const selectedIds = React.useMemo(
        () => new Set(selectedKeywords.map((k) => k.id)),
        [selectedKeywords]
    );

    const visibleKeywords = React.useMemo(() => {
        if (maxVisible <= 0 || showAll) return selectedKeywords;
        return selectedKeywords.slice(0, maxVisible);
    }, [selectedKeywords, maxVisible, showAll]);

    const overflowCount = React.useMemo(() => {
        if (maxVisible <= 0 || showAll) return 0;
        return Math.max(0, selectedKeywords.length - maxVisible);
    }, [selectedKeywords, maxVisible, showAll]);

    const canEdit = !isControlDisabled && !!entityId && !!entityName;

    // ── Render ────────────────────────────────────────────────────────────────

    if (isLoading) {
        return (
            <FluentProvider theme={webLightTheme}>
                <div className={styles.loadingWrapper}>
                    <Spinner size="tiny" />
                    <Text>Loading keywords…</Text>
                </div>
            </FluentProvider>
        );
    }

    return (
        <FluentProvider theme={webLightTheme}>
            <div className={styles.root}>
                {error && (
                    <MessageBar intent="error">
                        <MessageBarBody>{error}</MessageBarBody>
                    </MessageBar>
                )}

                <div className={styles.tagRow}>
                    {visibleKeywords.length === 0 && (!canEdit || !allowAssociate) && (
                        <Text className={styles.emptyLabel}>No keywords assigned.</Text>
                    )}

                    {visibleKeywords.length > 0 && (
                        <TagGroup
                            aria-label="Assigned keywords"
                            onDismiss={(_e, { value }) => {
                                const kw = selectedKeywords.find((k) => k.id === value);
                                if (kw) handleDismiss(kw);
                            }}
                        >
                            {visibleKeywords.map((kw) => (
                                <KeywordTag
                                    key={kw.id}
                                    keyword={kw}
                                    readOnly={!canEdit || !allowDisassociate}
                                    onDismiss={canEdit && allowDisassociate ? handleDismiss : undefined}
                                />
                            ))}
                        </TagGroup>
                    )}

                    {overflowCount > 0 && (
                        <OverflowBadge
                            count={overflowCount}
                            onClick={() => setShowAll(true)}
                        />
                    )}

                    {canEdit && allowAssociate && (
                        <div className={styles.actions}>
                            <Tooltip content="Add keyword" relationship="label">
                                <Button
                                    appearance="subtle"
                                    icon={<AddRegular />}
                                    size="small"
                                    onClick={() => setIsPanelOpen(true)}
                                    aria-label="Add keyword"
                                >
                                    {selectedKeywords.length === 0 ? (
                                        <>
                                            <TagRegular style={{ marginRight: "4px" }} />
                                            Add keyword
                                        </>
                                    ) : null}
                                </Button>
                            </Tooltip>
                        </div>
                    )}
                </div>

                <KeywordPickerPanel
                    isOpen={isPanelOpen}
                    availableKeywords={allKeywords}
                    selectedIds={selectedIds}
                    isLoading={false}
                    searchPlaceholder={searchPlaceholder}
                    allowCreate={allowCreate}
                    onApply={handleApply}
                    onDismiss={() => setIsPanelOpen(false)}
                />
            </div>
        </FluentProvider>
    );
};
