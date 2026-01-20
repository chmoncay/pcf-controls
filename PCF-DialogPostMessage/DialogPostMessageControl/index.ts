import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from "react";
import * as ReactDOM from "react-dom";
import { FluentProvider, webLightTheme, Button, Dialog, DialogTrigger, DialogSurface, DialogBody } from "@fluentui/react-components";

interface IDialogControlState {
    dialogUrl: string;
    receivedMessage: string;
}

export class DialogPostMessageControl implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private container: HTMLDivElement;
    private notifyOutputChanged: () => void;
    private context: ComponentFramework.Context<IInputs>;
    private state: IDialogControlState;
    private messageHandler: (event: MessageEvent) => void;

    /**
     * Empty constructor.
     */
    constructor() {
        this.state = {
            dialogUrl: "",
            receivedMessage: ""
        };
    }

    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
     */
    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {
        this.container = container;
        this.notifyOutputChanged = notifyOutputChanged;
        this.context = context;

        // Setup message handler for postMessage
        this.messageHandler = this.handlePostMessage.bind(this);
        window.addEventListener("message", this.messageHandler);

        // Initial render
        this.renderControl();
    }

    /**
     * Handle incoming postMessage events
     */
    private handlePostMessage(event: MessageEvent): void {
        // You might want to add origin validation here for security
        // if (event.origin !== expectedOrigin) return;
        
        this.state.receivedMessage = typeof event.data === 'string' ? event.data : JSON.stringify(event.data);
        this.renderControl();
    }

    /**
     * Render the React control
     */
    private renderControl(): void {
        const dialogUrl = this.context.parameters.dialogUrl.raw || "";
        this.state.dialogUrl = dialogUrl;

        ReactDOM.render(
            React.createElement(FluentProvider, { theme: webLightTheme },
                React.createElement("div", { style: { padding: "20px" } },
                    React.createElement(Dialog, { modalType: "modal", children: [
                        React.createElement(DialogTrigger, { disableButtonEnhancement: true, key: "trigger" },
                            React.createElement(Button, { appearance: "primary" }, "Open Dialog")
                        ),
                        React.createElement(DialogSurface, { style: { maxWidth: "90vw", maxHeight: "90vh" }, key: "surface" },
                            React.createElement(DialogBody, {},
                                React.createElement("iframe", {
                                    src: dialogUrl,
                                    style: {
                                        width: "800px",
                                        height: "600px",
                                        border: "none"
                                    },
                                    title: "Dialog Content"
                                })
                            )
                        )
                    ] }),
                    this.state.receivedMessage && React.createElement("div", {
                        style: {
                            marginTop: "20px",
                            padding: "15px",
                            backgroundColor: "#f3f2f1",
                            borderRadius: "4px",
                            border: "1px solid #d1d1d1"
                        }
                    },
                        React.createElement("strong", null, "Received Message:"),
                        React.createElement("div", { style: { marginTop: "8px", wordBreak: "break-word" } }, this.state.receivedMessage)
                    )
                )
            ),
            this.container
        );
    }


    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    public updateView(context: ComponentFramework.Context<IInputs>): void {
        this.context = context;
        this.renderControl();
    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as "bound" or "output"
     */
    public getOutputs(): IOutputs {
        return {};
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void {
        window.removeEventListener("message", this.messageHandler);
        ReactDOM.unmountComponentAtNode(this.container);
    }
}
