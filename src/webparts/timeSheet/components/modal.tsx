import * as React from 'react';
import * as ReactDOM from 'react-dom';
import styles from './TimeSheet.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';

import * as $ from 'jquery';
// declare let jQuery: any;

// console.log(jQuery.fn.extend)

export interface IModalButton {
    type: "primary" | "secondary" | "default" | "success" | "warning";
    text: string;
    closeModal: boolean;
    onClick?: (e: any) => void;
}

export interface IModalEventHandler {
    type: "show" | "shown" | "hide" | "hidden",
    handler: () => void;
}

export interface IModalProps {
    className?: string;
    modalId?: string;
    titleContent?: any;
    bodyContent?: any;
    buttons?: IModalButton[],
    onMount?: (jquery: any) => void;
    eventHandlers?: IModalEventHandler[];
    size?: "large" | "small" | "";
}

export default class Modal extends React.Component<IModalProps, any> {

    private $parent: any = null;
    private $node: any = null;
    public state: any = {
        className: "",
        modalId: ""
    };

    public constructor(public props: IModalProps, context?: any) {
        super(props, context);
        this.state.className = this.props.className || "fade";
        this.state.modalId = this.props.modalId || "modal-" + Math.round(Math.random() * 10000000000);
        this.props.buttons = this.props.buttons || [{ type: "secondary", text: "Close", closeModal: true }];
    }

    componentWillUnmount() {
        if (this.$parent !== null) {
            this.$parent.append(this.$node);
        }
    }

    componentDidMount() {
        if (this.$node !== null) {
            
        }

        this.$node = $(ReactDOM.findDOMNode(this));
        this.$parent = this.$node.parent();
        $("body").append(this.$node);

        (this.props.onMount || ((v) => {}))(this.$node);

        if (this.props.eventHandlers) {
            for (let i = 0; i < this.props.eventHandlers.length; i++) {
                let handler = this.props.eventHandlers[i];
                this.$node.on(`${handler.type}.bs.modal`, () => { handler.handler(); });
            }
        }
    }

    getModalSizeClass() : string {
        switch (this.props.size) {
            case "large": return "modal-lg";
            case "small": return "modal-sm";
            case "": return "";
            default: return "";
        }
    }

  public render(): React.ReactElement<IModalProps> {
    return (<div className={"modal " + (this.state.className)} role="dialog" id={this.state.modalId}>
        <div className={`modal-dialog ${this.getModalSizeClass()}`} role="document">
            <div className="modal-content">
            <div className="modal-header">
                <h5 className="modal-title">{this.props.titleContent}</h5>
                <button type="button" className="close" data-dismiss="modal" aria-label="Close">
                    <span aria-hidden="true">&times;</span>
                </button>
            </div>
            <div className="modal-body">
                <p>{this.props.bodyContent}</p>
            </div>
            <div className="modal-footer">
                {(this.props.buttons || []).map(b => b.closeModal ?
                    <button type="button" className={`btn btn-${b.type}`} onClick={(e) => (b.onClick || ((v) => {}))(e)} data-dismiss="modal">{b.text}</button> :
                    <button type="button" className={`btn btn-${b.type}`} onClick={(e) => (b.onClick || ((v) => {}))(e)}>{b.text}</button>
                )}
            </div>
            </div>
        </div>
    </div>);
  }
}
