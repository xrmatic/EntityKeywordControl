import * as React from "react";
import { EntityKeywordControlComponent } from "./components/EntityKeywordControl";
import { IInputs, IOutputs } from "./components/types";

/**
 * EntityKeywordControl – PCF virtual (React) control
 *
 * Presents a Many-to-Many related list of keyword records as interactive
 * Fluent UI tags.  Keywords support per-record colour, icon, group and
 * description metadata to give richer visual feedback than a standard
 * multi-select lookup.
 */
export class EntityKeywordControl
    implements ComponentFramework.ReactControl<IInputs, IOutputs> {

    private notifyOutputChanged!: () => void;

    /**
     * Called once when the control is first placed on the form.
     */
    public init(
        _context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        _state: ComponentFramework.Dictionary
    ): void {
        this.notifyOutputChanged = notifyOutputChanged;
    }

    /**
     * Called every time the host form triggers a re-render.
     * Returns the React element that the PCF framework will mount.
     */
    public updateView(
        context: ComponentFramework.Context<IInputs>
    ): React.ReactElement {
        return React.createElement(EntityKeywordControlComponent, {
            context,
            notifyOutputChanged: this.notifyOutputChanged,
        });
    }

    /**
     * Returns any output property values.
     * Currently the control drives all mutations through WebAPI calls
     * and has no output bindings.
     */
    public getOutputs(): IOutputs {
        return {};
    }

    /**
     * Cleanup – called when the control is removed from the form.
     */
    public destroy(): void {
        // No manual cleanup needed; React handles its own lifecycle.
    }
}
