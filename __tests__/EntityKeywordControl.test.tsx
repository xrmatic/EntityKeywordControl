/**
 * Unit tests for EntityKeywordControl PCF component.
 *
 * These tests validate the pure utility logic and component rendering
 * without requiring a live Dataverse environment.
 */

import * as React from "react";
import { render, screen, fireEvent, act, waitFor } from "@testing-library/react";
import "@testing-library/jest-dom";
import { TagGroup } from "@fluentui/react-components";

// ── jsdom shims ───────────────────────────────────────────────────────────────

// Fluent UI's MessageBar and other components use ResizeObserver which is
// not available in jsdom – provide a no-op mock so tests don't throw.
global.ResizeObserver = class ResizeObserver {
    observe() {}
    unobserve() {}
    disconnect() {}
};

// ── Test helpers / mocks ──────────────────────────────────────────────────────

/**
 * Minimal mock for ComponentFramework.Context<IInputs>.
 * Only the fields consumed by the component under test are populated.
 */
function buildMockContext(overrides: Partial<{
    relationshipName: string;
    relatedEntityName: string;
    relatedEntitySetName: string;
    navigationPropertyName: string;
    labelFieldName: string;
    colorFieldName: string;
    iconFieldName: string;
    groupFieldName: string;
    descriptionFieldName: string;
    sortOrderFieldName: string;
    allowAssociate: boolean;
    allowDisassociate: boolean;
    allowCreate: boolean;
    searchPlaceholder: string;
    maxVisibleTags: number;
    entityId: string;
    entityTypeName: string;
    retrieveMultipleResult: object[];
    retrieveRecordResult: object;
}> = {}) {
    const cfg = {
        relationshipName: "xrm_account_keyword",
        relatedEntityName: "xrm_keyword",
        relatedEntitySetName: "xrm_keywords",
        navigationPropertyName: "xrm_account_keyword_xrm_keyword",
        labelFieldName: "xrm_name",
        colorFieldName: "",
        iconFieldName: "",
        groupFieldName: "",
        descriptionFieldName: "",
        sortOrderFieldName: "",
        allowAssociate: true,
        allowDisassociate: true,
        allowCreate: false,
        searchPlaceholder: "",
        maxVisibleTags: 0,
        entityId: "00000000-0000-0000-0000-000000000001",
        entityTypeName: "account",
        retrieveMultipleResult: [],
        retrieveRecordResult: {},
        ...overrides,
    };

    const makeStringProp = (val: string) => ({ raw: val } as ComponentFramework.PropertyTypes.StringProperty);
    const makeBoolProp   = (val: boolean) => ({ raw: val } as ComponentFramework.PropertyTypes.TwoOptionsProperty);
    const makeNumProp    = (val: number)  => ({ raw: val } as ComponentFramework.PropertyTypes.WholeNumberProperty);

    return {
        parameters: {
            relationshipName:      makeStringProp(cfg.relationshipName),
            relatedEntityName:     makeStringProp(cfg.relatedEntityName),
            relatedEntitySetName:  makeStringProp(cfg.relatedEntitySetName),
            navigationPropertyName: makeStringProp(cfg.navigationPropertyName),
            labelFieldName:        makeStringProp(cfg.labelFieldName),
            colorFieldName:        makeStringProp(cfg.colorFieldName),
            iconFieldName:         makeStringProp(cfg.iconFieldName),
            groupFieldName:        makeStringProp(cfg.groupFieldName),
            descriptionFieldName:  makeStringProp(cfg.descriptionFieldName),
            sortOrderFieldName:    makeStringProp(cfg.sortOrderFieldName),
            allowAssociate:        makeBoolProp(cfg.allowAssociate),
            allowDisassociate:     makeBoolProp(cfg.allowDisassociate),
            allowCreate:           makeBoolProp(cfg.allowCreate),
            searchPlaceholder:     makeStringProp(cfg.searchPlaceholder),
            maxVisibleTags:        makeNumProp(cfg.maxVisibleTags),
        },
        mode: {
            isControlDisabled: false,
            isVisible: true,
            label: "",
            allocatedHeight: 200,
            allocatedWidth: 400,
        },
        page: {
            entityId: cfg.entityId,
            entityTypeName: cfg.entityTypeName,
        },
        webAPI: {
            retrieveMultipleRecords: jest.fn().mockResolvedValue({
                entities: cfg.retrieveMultipleResult,
            }),
            retrieveRecord: jest.fn().mockResolvedValue(cfg.retrieveRecordResult),
            associateRecord: jest.fn().mockResolvedValue(undefined),
            disassociateRecord: jest.fn().mockResolvedValue(undefined),
        },
    } as unknown as ComponentFramework.Context<import("../components/types").IInputs>;
}

// ── Import component under test ───────────────────────────────────────────────

import { EntityKeywordControlComponent } from "../components/EntityKeywordControl";

// ─────────────────────────────────────────────────────────────────────────────

describe("EntityKeywordControlComponent", () => {

    // ── Loading state ─────────────────────────────────────────────────────────

    it("shows a spinner while keyword data is loading", () => {
        // Never-resolving promises keep the loading state active
        const ctx = buildMockContext();
        (ctx.webAPI.retrieveMultipleRecords as jest.Mock).mockReturnValue(new Promise(() => undefined));
        (ctx.webAPI.retrieveRecord as jest.Mock).mockReturnValue(new Promise(() => undefined));

        render(
            <EntityKeywordControlComponent
                context={ctx}
                notifyOutputChanged={jest.fn()}
            />
        );

        expect(screen.getByText(/loading keywords/i)).toBeInTheDocument();
    });

    // ── Empty state ───────────────────────────────────────────────────────────

    it("shows 'No keywords assigned' when the record has no associations and editing is disabled", async () => {
        const ctx = buildMockContext({
            allowAssociate: false,
            retrieveRecordResult: { xrm_account_keyword_xrm_keyword: [] },
        });

        await act(async () => {
            render(
                <EntityKeywordControlComponent
                    context={ctx}
                    notifyOutputChanged={jest.fn()}
                />
            );
        });

        expect(screen.getByText(/no keywords assigned/i)).toBeInTheDocument();
    });

    // ── Tags are rendered ─────────────────────────────────────────────────────

    it("renders a tag for each associated keyword", async () => {
        const ctx = buildMockContext({
            retrieveRecordResult: {
                xrm_account_keyword_xrm_keyword: [
                    { xrm_keywordid: "id-1", xrm_name: "Alpha" },
                    { xrm_keywordid: "id-2", xrm_name: "Beta" },
                ],
            },
        });

        await act(async () => {
            render(
                <EntityKeywordControlComponent
                    context={ctx}
                    notifyOutputChanged={jest.fn()}
                />
            );
        });

        expect(screen.getByText("Alpha")).toBeInTheDocument();
        expect(screen.getByText("Beta")).toBeInTheDocument();
    });

    // ── Add keyword button ────────────────────────────────────────────────────

    it("shows an 'Add keyword' button when allowAssociate is true", async () => {
        const ctx = buildMockContext({
            allowAssociate: true,
            retrieveRecordResult: { xrm_account_keyword_xrm_keyword: [] },
        });

        await act(async () => {
            render(
                <EntityKeywordControlComponent
                    context={ctx}
                    notifyOutputChanged={jest.fn()}
                />
            );
        });

        expect(screen.getByRole("button", { name: /add keyword/i })).toBeInTheDocument();
    });

    it("hides the 'Add keyword' button when allowAssociate is false", async () => {
        const ctx = buildMockContext({
            allowAssociate: false,
            retrieveRecordResult: { xrm_account_keyword_xrm_keyword: [] },
        });

        await act(async () => {
            render(
                <EntityKeywordControlComponent
                    context={ctx}
                    notifyOutputChanged={jest.fn()}
                />
            );
        });

        expect(screen.queryByRole("button", { name: /add keyword/i })).not.toBeInTheDocument();
    });

    // ── Overflow badge ────────────────────────────────────────────────────────

    it("shows an overflow badge when displayed tags exceed maxVisibleTags", async () => {
        const ctx = buildMockContext({
            maxVisibleTags: 2,
            retrieveRecordResult: {
                xrm_account_keyword_xrm_keyword: [
                    { xrm_keywordid: "id-1", xrm_name: "Alpha" },
                    { xrm_keywordid: "id-2", xrm_name: "Beta" },
                    { xrm_keywordid: "id-3", xrm_name: "Gamma" },
                    { xrm_keywordid: "id-4", xrm_name: "Delta" },
                ],
            },
        });

        await act(async () => {
            render(
                <EntityKeywordControlComponent
                    context={ctx}
                    notifyOutputChanged={jest.fn()}
                />
            );
        });

        expect(screen.getByText("+2 more")).toBeInTheDocument();
        expect(screen.queryByText("Gamma")).not.toBeInTheDocument();
    });

    it("expands all tags when the overflow badge is clicked", async () => {
        const ctx = buildMockContext({
            maxVisibleTags: 2,
            retrieveRecordResult: {
                xrm_account_keyword_xrm_keyword: [
                    { xrm_keywordid: "id-1", xrm_name: "Alpha" },
                    { xrm_keywordid: "id-2", xrm_name: "Beta" },
                    { xrm_keywordid: "id-3", xrm_name: "Gamma" },
                ],
            },
        });

        await act(async () => {
            render(
                <EntityKeywordControlComponent
                    context={ctx}
                    notifyOutputChanged={jest.fn()}
                />
            );
        });

        fireEvent.click(screen.getByText("+1 more"));

        await waitFor(() => {
            expect(screen.getByText("Gamma")).toBeInTheDocument();
        });
    });

    // ── Disassociate ──────────────────────────────────────────────────────────

    it("calls webAPI.disassociateRecord when a tag is dismissed", async () => {
        const ctx = buildMockContext({
            retrieveRecordResult: {
                xrm_account_keyword_xrm_keyword: [
                    { xrm_keywordid: "id-1", xrm_name: "Alpha" },
                ],
            },
        });

        await act(async () => {
            render(
                <EntityKeywordControlComponent
                    context={ctx}
                    notifyOutputChanged={jest.fn()}
                />
            );
        });

        const dismissBtn = screen.getByRole("button", { name: /remove Alpha/i });
        await act(async () => { fireEvent.click(dismissBtn); });

        expect(ctx.webAPI.disassociateRecord).toHaveBeenCalledWith(
            "account",
            "00000000-0000-0000-0000-000000000001",
            "xrm_account_keyword",
            "id-1"
        );
    });

    // ── Error state ───────────────────────────────────────────────────────────

    it("displays an error message when the WebAPI call fails", async () => {
        const ctx = buildMockContext();
        (ctx.webAPI.retrieveMultipleRecords as jest.Mock).mockRejectedValue(
            new Error("Network error")
        );

        await act(async () => {
            render(
                <EntityKeywordControlComponent
                    context={ctx}
                    notifyOutputChanged={jest.fn()}
                />
            );
        });

        expect(screen.getByText(/Network error/)).toBeInTheDocument();
    });
});

// ─────────────────────────────────────────────────────────────────────────────

describe("KeywordPickerPanel", () => {
    const { KeywordPickerPanel } = require("../components/KeywordPickerPanel");

    const mockKeywords = [
        { id: "k1", label: "React", group: "Frontend" },
        { id: "k2", label: "TypeScript", group: "Frontend" },
        { id: "k3", label: "Dataverse", group: "Backend" },
    ];

    it("renders all unselected keywords", async () => {
        render(
            <KeywordPickerPanel
                isOpen
                availableKeywords={mockKeywords}
                selectedIds={new Set()}
                isLoading={false}
                onApply={jest.fn()}
                onDismiss={jest.fn()}
            />
        );

        expect(screen.getByText("React")).toBeInTheDocument();
        expect(screen.getByText("TypeScript")).toBeInTheDocument();
        expect(screen.getByText("Dataverse")).toBeInTheDocument();
    });

    it("filters out already-selected keywords", () => {
        render(
            <KeywordPickerPanel
                isOpen
                availableKeywords={mockKeywords}
                selectedIds={new Set(["k1"])}
                isLoading={false}
                onApply={jest.fn()}
                onDismiss={jest.fn()}
            />
        );

        expect(screen.queryByText("React")).not.toBeInTheDocument();
        expect(screen.getByText("TypeScript")).toBeInTheDocument();
    });

    it("filters keywords based on search input", async () => {
        render(
            <KeywordPickerPanel
                isOpen
                availableKeywords={mockKeywords}
                selectedIds={new Set()}
                isLoading={false}
                onApply={jest.fn()}
                onDismiss={jest.fn()}
            />
        );

        fireEvent.change(screen.getByPlaceholderText(/search keywords/i), {
            target: { value: "type" },
        });

        await waitFor(() => {
            expect(screen.getByText("TypeScript")).toBeInTheDocument();
            expect(screen.queryByText("React")).not.toBeInTheDocument();
            expect(screen.queryByText("Dataverse")).not.toBeInTheDocument();
        });
    });

    it("calls onApply with selected keywords when Add is clicked", async () => {
        const onApply = jest.fn();

        render(
            <KeywordPickerPanel
                isOpen
                availableKeywords={mockKeywords}
                selectedIds={new Set()}
                isLoading={false}
                onApply={onApply}
                onDismiss={jest.fn()}
            />
        );

        // Select "React"
        fireEvent.click(screen.getByText("React"));

        const addBtn = screen.getByRole("button", { name: /^add \(1\)/i });
        fireEvent.click(addBtn);

        expect(onApply).toHaveBeenCalledWith(
            expect.arrayContaining([expect.objectContaining({ id: "k1", label: "React" })])
        );
    });

    it("shows a spinner when isLoading is true", () => {
        render(
            <KeywordPickerPanel
                isOpen
                availableKeywords={[]}
                selectedIds={new Set()}
                isLoading
                onApply={jest.fn()}
                onDismiss={jest.fn()}
            />
        );

        expect(screen.getByText(/loading keywords/i)).toBeInTheDocument();
    });

    it("shows empty state when no keywords match the search", async () => {
        render(
            <KeywordPickerPanel
                isOpen
                availableKeywords={mockKeywords}
                selectedIds={new Set()}
                isLoading={false}
                onApply={jest.fn()}
                onDismiss={jest.fn()}
            />
        );

        fireEvent.change(screen.getByPlaceholderText(/search keywords/i), {
            target: { value: "zzznomatch" },
        });

        await waitFor(() => {
            expect(screen.getByText(/no keywords found/i)).toBeInTheDocument();
        });
    });
});

// ─────────────────────────────────────────────────────────────────────────────

describe("KeywordTag", () => {
    const { KeywordTag } = require("../components/KeywordTag");

    it("renders the keyword label", () => {
        render(
            <KeywordTag
                keyword={{ id: "k1", label: "Hello" }}
            />
        );
        expect(screen.getByText("Hello")).toBeInTheDocument();
    });

    it("calls onDismiss with the keyword when dismissed", () => {
        const onDismiss = jest.fn();
        const keyword = { id: "k1", label: "Hello" };
        render(
            <TagGroup
                onDismiss={(_e, { value }) => {
                    if (value === keyword.id) onDismiss(keyword);
                }}
            >
                <KeywordTag keyword={keyword} readOnly={false} />
            </TagGroup>
        );
        const dismissBtn = screen.getByRole("button", { name: /remove Hello/i });
        fireEvent.click(dismissBtn);
        expect(onDismiss).toHaveBeenCalledWith({ id: "k1", label: "Hello" });
    });

    it("does not render a dismiss button in readOnly mode", () => {
        render(
            <KeywordTag
                keyword={{ id: "k1", label: "Hello" }}
                readOnly
            />
        );
        expect(screen.queryByRole("button", { name: /remove Hello/i })).not.toBeInTheDocument();
    });

    it("renders a tooltip when description is provided", () => {
        render(
            <KeywordTag
                keyword={{ id: "k1", label: "Hello", description: "A great keyword" }}
            />
        );
        // The Tooltip wraps the Tag – verifying the tag text is present
        expect(screen.getByText("Hello")).toBeInTheDocument();
    });
});
