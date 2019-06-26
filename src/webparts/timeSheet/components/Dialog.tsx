// import { BaseDialog, IDialogConfiguration } from "@microsoft/sp-dialog";
// import {
//     autobind,
//     ColorPicker,
//     PrimaryButton,
//     Button,
//     DialogFooter,
//     DialogContent
// } from 'office-ui-fabric-react';
// import * as ReactDom from 'react-dom';

// export interface IDialogProps {
//     html: string;
//     close: () => void;
//     submit: (color: string) => void;
// }

// export default class Dialog extends BaseDialog {
//     public message: string;
//     public colorCode: string;

//     public render(): void {
//         ReactDOM.render((<DialogContent
//             title='Color Picker'
//             subText={this.props.message}
//             onDismiss={this.props.close}
//             showCloseButton={true}
//         >
//             <DialogFooter>
//                 <Button text='Cancel' title='Cancel' onClick={this.props.close} />
//                 <PrimaryButton text='OK' title='OK' onClick={() => { this.props.submit(this._pickedColor); }} />
//             </DialogFooter>
//         </DialogContent>), this.domElement);
//     }

//     public getConfig(): IDialogConfiguration {
//         return {
//             isBlocking: false
//         };
//     }

//     protected onAfterClose(): void {
//         super.onAfterClose();

//         // Clean up the element for the next dialog
//         ReactDOM.unmountComponentAtNode(this.domElement);
//     }
// }