import * as React from "react";
import {
    Tag,
    TagGroup,
    TagGroupProps,
    TagProps,
    makeStyles,
    tokens,
    Tooltip,
    Image,
} from "@fluentui/react-components";
import { IKeywordRecord } from "./types";

// ─── Styles ──────────────────────────────────────────────────────────────────

const useStyles = makeStyles({
    tagRoot: {
        cursor: "default",
        borderRadius: tokens.borderRadiusMedium,
        transition: "opacity 150ms ease",
    },
    tagIcon: {
        width: "14px",
        height: "14px",
        objectFit: "contain",
        borderRadius: tokens.borderRadiusSmall,
    },
    overflowBadge: {
        display: "inline-flex",
        alignItems: "center",
        justifyContent: "center",
        height: "28px",
        minWidth: "40px",
        padding: `0 ${tokens.spacingHorizontalS}`,
        borderRadius: tokens.borderRadiusMedium,
        backgroundColor: tokens.colorNeutralBackground3,
        color: tokens.colorNeutralForeground2,
        fontSize: tokens.fontSizeBase200,
        fontWeight: tokens.fontWeightSemibold,
        cursor: "pointer",
        userSelect: "none",
        border: `1px solid ${tokens.colorNeutralStroke2}`,
    },
});

// ─── Helper – derive inline colour overrides ─────────────────────────────────

function buildColorStyle(color: string | undefined): React.CSSProperties | undefined {
    if (!color) return undefined;
    // Lighten the provided hex colour for the background (30 % opacity via alpha)
    return {
        backgroundColor: `${color}33`,   // hex + ~20% alpha (0x33 = 51/255 ≈ 20%)
        borderColor: `${color}99`,        // hex + 60 % alpha
        color: color,
    };
}

// ─── TagGroupWrapper (for standalone dismiss support) ─────────────────────────

/**
 * A zero-layout wrapper around a single Tag that provides the TagGroup context
 * needed for dismiss events when KeywordTag is used outside an explicit TagGroup.
 */
const TagGroupWrapper: React.FC<{ onDismiss: TagGroupProps["onDismiss"]; children: React.ReactNode }> = ({
    onDismiss,
    children,
}) => (
    <TagGroup
        onDismiss={onDismiss}
        style={{ display: "contents" }}
    >
        {children}
    </TagGroup>
);

// ─── KeywordTag ───────────────────────────────────────────────────────────────

export interface IKeywordTagProps {
    keyword: IKeywordRecord;
    /**
     * Called when the dismiss (×) button is clicked.
     *
     * NOTE: In Fluent UI v9 the dismiss event bubbles through the enclosing
     * `TagGroup.onDismiss` handler.  When `KeywordTag` is used inside a
     * `TagGroup`, wire `TagGroup.onDismiss` and look up the keyword by `value`.
     * This prop is provided for standalone usage outside a `TagGroup`.
     */
    onDismiss?: (keyword: IKeywordRecord) => void;
    /** When true the dismiss button is hidden. */
    readOnly?: boolean;
    /** Forwarded to the Fluent UI Tag appearance prop. */
    appearance?: TagProps["appearance"];
    /** Forwarded to the Fluent UI Tag size prop. */
    size?: TagProps["size"];
    /** Forwarded to the Fluent UI Tag shape prop. */
    shape?: TagProps["shape"];
}

export const KeywordTag: React.FC<IKeywordTagProps> = ({
    keyword,
    onDismiss,
    readOnly = false,
    appearance,
    size,
    shape,
}) => {
    const styles = useStyles();
    const colorStyle = buildColorStyle(keyword.color);

    // Resolve icon: treat as an image URL if it contains "/" or ".", otherwise
    // treat as a Fluent icon name stub and render a plain emoji-style text.
    const iconElement = React.useMemo<React.ReactElement | undefined>(() => {
        if (!keyword.icon) return undefined;
        if (keyword.icon.startsWith("http") || keyword.icon.includes("/") || keyword.icon.includes(".")) {
            return (
                <Image
                    src={keyword.icon}
                    alt={keyword.label}
                    className={styles.tagIcon}
                />
            );
        }
        // Fallback: first character of the icon name as a visual badge
        return <span aria-hidden="true">{keyword.icon.charAt(0)}</span>;
    }, [keyword.icon, keyword.label, styles.tagIcon]);

    // When the Tag is used outside a TagGroup (standalone) and onDismiss is
    // provided, wrap in a lightweight TagGroup so the dismiss event is caught.
    const tagElement = (
        <Tag
            appearance={appearance}
            size={size}
            shape={shape}
            className={styles.tagRoot}
            style={colorStyle}
            dismissible={!readOnly}
            dismissIcon={{
                "aria-label": `Remove ${keyword.label}`,
            }}
            {...(iconElement ? { icon: iconElement } : {})}
            value={keyword.id}
        >
            {keyword.label}
        </Tag>
    );

    const tag = onDismiss && !readOnly ? (
        // Minimal TagGroup wrapper so the built-in dismiss event is captured
        // without requiring callers to always supply a TagGroup.
        <TagGroupWrapper onDismiss={() => onDismiss(keyword)}>
            {tagElement}
        </TagGroupWrapper>
    ) : tagElement;

    if (keyword.description) {
        return (
            <Tooltip content={keyword.description} relationship="description">
                {tag}
            </Tooltip>
        );
    }
    return tag;
};

// ─── OverflowBadge ────────────────────────────────────────────────────────────

export interface IOverflowBadgeProps {
    count: number;
    onClick?: () => void;
}

export const OverflowBadge: React.FC<IOverflowBadgeProps> = ({ count, onClick }) => {
    const styles = useStyles();
    return (
        <span
            className={styles.overflowBadge}
            role="button"
            tabIndex={0}
            aria-label={`Show ${count} more keywords`}
            onClick={onClick}
            onKeyDown={(e) => {
                if (e.key === "Enter" || e.key === " ") onClick?.();
            }}
        >
            +{count} more
        </span>
    );
};
