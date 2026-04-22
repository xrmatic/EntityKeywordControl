import * as React from "react";
import {
    Input,
    makeStyles,
    tokens,
    Spinner,
    Text,
    Divider,
    Badge,
    Button,
    Checkbox,
    Dialog,
    DialogSurface,
    DialogTitle,
    DialogBody,
    DialogActions,
    DialogContent,
} from "@fluentui/react-components";
import { SearchRegular, AddRegular } from "@fluentui/react-icons";
import { IKeywordRecord } from "./types";

// ─── Styles ───────────────────────────────────────────────────────────────────

const useStyles = makeStyles({
    container: {
        display: "flex",
        flexDirection: "column",
        gap: tokens.spacingVerticalS,
        padding: `0 ${tokens.spacingHorizontalM}`,
        height: "100%",
        overflowY: "auto",
    },
    searchBox: {
        width: "100%",
    },
    groupLabel: {
        fontWeight: tokens.fontWeightSemibold,
        color: tokens.colorNeutralForeground3,
        fontSize: tokens.fontSizeBase200,
        textTransform: "uppercase",
        letterSpacing: "0.05em",
        marginTop: tokens.spacingVerticalM,
    },
    keywordRow: {
        display: "flex",
        alignItems: "center",
        gap: tokens.spacingHorizontalS,
        padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalS}`,
        borderRadius: tokens.borderRadiusMedium,
        cursor: "pointer",
        ":hover": {
            backgroundColor: tokens.colorNeutralBackground2Hover,
        },
    },
    colorDot: {
        width: "10px",
        height: "10px",
        borderRadius: "50%",
        flexShrink: 0,
    },
    keywordLabel: {
        flexGrow: 1,
        overflow: "hidden",
        textOverflow: "ellipsis",
        whiteSpace: "nowrap",
    },
    keywordDescription: {
        color: tokens.colorNeutralForeground3,
        fontSize: tokens.fontSizeBase100,
    },
    emptyState: {
        textAlign: "center",
        color: tokens.colorNeutralForeground3,
        padding: tokens.spacingVerticalXXL,
    },
    newKeywordRow: {
        display: "flex",
        gap: tokens.spacingHorizontalS,
        marginTop: tokens.spacingVerticalM,
        alignItems: "center",
    },
    footer: {
        display: "flex",
        gap: tokens.spacingHorizontalS,
        justifyContent: "flex-end",
        paddingTop: tokens.spacingVerticalM,
        borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
        marginTop: "auto",
    },
});

// ─── Types ────────────────────────────────────────────────────────────────────

export interface IKeywordPickerPanelProps {
    /** Whether the panel is open. */
    isOpen: boolean;
    /** All available keyword records. */
    availableKeywords: IKeywordRecord[];
    /** IDs of keywords already associated (to exclude from the list). */
    selectedIds: Set<string>;
    /** Whether the list is still loading. */
    isLoading: boolean;
    /** Placeholder text for the search input. */
    searchPlaceholder?: string;
    /** Whether to show a "New keyword" row when search term matches nothing. */
    allowCreate?: boolean;
    /** Called when the user confirms their selection. */
    onApply: (toAssociate: IKeywordRecord[]) => void;
    /** Called when the panel is dismissed without applying. */
    onDismiss: () => void;
}

// ─── Component ────────────────────────────────────────────────────────────────

export const KeywordPickerPanel: React.FC<IKeywordPickerPanelProps> = ({
    isOpen,
    availableKeywords,
    selectedIds,
    isLoading,
    searchPlaceholder = "Search keywords…",
    allowCreate = false,
    onApply,
    onDismiss,
}) => {
    const styles = useStyles();
    const [searchText, setSearchText] = React.useState("");
    const [pendingIds, setPendingIds] = React.useState<Set<string>>(new Set());

    // Reset state whenever the panel opens
    React.useEffect(() => {
        if (isOpen) {
            setSearchText("");
            setPendingIds(new Set());
        }
    }, [isOpen]);

    // ── Filter & group ────────────────────────────────────────────────────────

    const filteredKeywords = React.useMemo(() => {
        const query = searchText.trim().toLowerCase();
        return availableKeywords.filter((kw) => {
            if (selectedIds.has(kw.id)) return false;           // already associated
            if (!query) return true;
            return (
                kw.label.toLowerCase().includes(query) ||
                (kw.description?.toLowerCase().includes(query) ?? false) ||
                (kw.group?.toLowerCase().includes(query) ?? false)
            );
        });
    }, [availableKeywords, selectedIds, searchText]);

    const groupedKeywords = React.useMemo(() => {
        const groups = new Map<string, IKeywordRecord[]>();
        for (const kw of filteredKeywords) {
            const key = kw.group ?? "";
            const list = groups.get(key) ?? [];
            list.push(kw);
            groups.set(key, list);
        }
        return groups;
    }, [filteredKeywords]);

    // ── Selection helpers ─────────────────────────────────────────────────────

    const togglePending = (kw: IKeywordRecord) => {
        setPendingIds((prev) => {
            const next = new Set(prev);
            if (next.has(kw.id)) {
                next.delete(kw.id);
            } else {
                next.add(kw.id);
            }
            return next;
        });
    };

    const handleApply = () => {
        const toAssociate = availableKeywords.filter((kw) => pendingIds.has(kw.id));
        onApply(toAssociate);
    };

    // ── Render helpers ────────────────────────────────────────────────────────

    const renderKeywordRow = (kw: IKeywordRecord) => {
        const checked = pendingIds.has(kw.id);
        return (
            <div
                key={kw.id}
                className={styles.keywordRow}
                onClick={() => togglePending(kw)}
                role="checkbox"
                aria-checked={checked}
                tabIndex={0}
                onKeyDown={(e) => {
                    if (e.key === "Enter" || e.key === " ") togglePending(kw);
                }}
            >
                <Checkbox
                    checked={checked}
                    onChange={() => togglePending(kw)}
                    tabIndex={-1}
                />
                {kw.color && (
                    <span
                        className={styles.colorDot}
                        style={{ backgroundColor: kw.color }}
                        aria-hidden="true"
                    />
                )}
                <div className={styles.keywordLabel}>
                    <Text>{kw.label}</Text>
                    {kw.description && (
                        <Text block className={styles.keywordDescription}>
                            {kw.description}
                        </Text>
                    )}
                </div>
                {kw.group && (
                    <Badge appearance="outline" size="small" color="informative">
                        {kw.group}
                    </Badge>
                )}
            </div>
        );
    };

    const renderBody = () => {
        if (isLoading) {
            return <Spinner label="Loading keywords…" />;
        }

        if (filteredKeywords.length === 0) {
            return (
                <div className={styles.emptyState}>
                    <Text>No keywords found{searchText ? ` for "${searchText}"` : ""}.</Text>
                    {allowCreate && searchText && (
                        <div className={styles.newKeywordRow}>
                            <Button
                                icon={<AddRegular />}
                                appearance="subtle"
                                onClick={() => {
                                    // Caller can intercept via a dedicated onCreateRequested prop
                                    // (future extension). For now we close with the search text
                                    // as a signal in the apply callback via a synthetic record.
                                    const syntheticRecord: IKeywordRecord = {
                                        id: `__new__::${searchText}`,
                                        label: searchText,
                                    };
                                    onApply([syntheticRecord]);
                                }}
                            >
                                Create &ldquo;{searchText}&rdquo;
                            </Button>
                        </div>
                    )}
                </div>
            );
        }

        const sections: React.ReactNode[] = [];
        for (const [group, keywords] of groupedKeywords) {
            if (group) {
                sections.push(
                    <React.Fragment key={`group-${group}`}>
                        <Text className={styles.groupLabel}>{group}</Text>
                        <Divider />
                    </React.Fragment>
                );
            }
            keywords.forEach((kw) => sections.push(renderKeywordRow(kw)));
        }
        return sections;
    };

    // ── Dialog render ─────────────────────────────────────────────────────────

    return (
        <Dialog open={isOpen} onOpenChange={(_e, data) => { if (!data.open) onDismiss(); }}>
            <DialogSurface style={{ minWidth: "420px", maxWidth: "560px" }}>
                <DialogTitle>Add Keywords</DialogTitle>
                <DialogBody>
                    <DialogContent>
                        <div className={styles.container}>
                            <Input
                                className={styles.searchBox}
                                contentBefore={<SearchRegular />}
                                placeholder={searchPlaceholder}
                                value={searchText}
                                onChange={(_, data) => setSearchText(data.value)}
                                aria-label="Search keywords"
                            />
                            {renderBody()}
                        </div>
                    </DialogContent>
                </DialogBody>
                <DialogActions>
                    <Button appearance="secondary" onClick={onDismiss}>
                        Cancel
                    </Button>
                    <Button
                        appearance="primary"
                        disabled={pendingIds.size === 0}
                        onClick={handleApply}
                    >
                        Add{pendingIds.size > 0 ? ` (${pendingIds.size})` : ""}
                    </Button>
                </DialogActions>
            </DialogSurface>
        </Dialog>
    );
};
