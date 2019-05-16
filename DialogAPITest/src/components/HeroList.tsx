import * as React from "react";

export interface HeroListItem {
  icon: string;
  primaryText: string;
}

export interface HeroListProps {
  message: string;
  items: HeroListItem[];
}
let dialog;
let inProgress = false;
let btnEvent;
export default class HeroList extends React.Component<HeroListProps> {
  render() {
    return (
      <main className="ms-welcome__main">
        <button onClick={this.btnClick}>Open Dialog</button>
        <button onClick={this.btnCloseClick}>Close Dialog</button>
      </main>
    );
  }
  componentDidMount() {
    // if (window.location.search.includes("code")) {
    //   if (dialog != null) {
    //     dialog.close();
    //     inProgress = false;
    //   }
    // }
  }
  btnClick = event => {
    btnEvent = event;
    if (inProgress == false) {
      Office.context.ui.displayDialogAsync(
        "https://localhost:3000/assets/auth.html",
        { height: 300, width: 200, displayInIframe: false },
        result => {
          dialog = result.value;
          dialog.addEventHandler(
            Office.EventType.DialogMessageReceived,
            this.openArchivePage
          );
          inProgress = true;
          //   setTimeout(function() {
          //     if (dialog != null) {
          //       dialog.close();
          //       inProgress = false;
          //     }

          //     if (event != null) {
          //       event.completed();
          //     }
          //   }, 1000);
        }
      );
    }
  };
  btnCloseClick() {
    // if (dialog !== null) {
    //   dialog.close();
    // }
    var messageObject = { messageType: "dialogClosed" };
    var jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
  }
  openArchivePage = arg => {
    // if (dialog != null) {
    //   dialog.close();
    //   dialog = null;
    // }

    // if (btnEvent != null) {
    //   // btnEvent.completed();
    //   btnEvent = null;
    // }
    var messageFromDialog = JSON.parse(arg.message);
    if (messageFromDialog.messageType === "dialogClosed") {
      dialog.close();
    }
  };
  dialogCallback2 = asyncResult => {
    if (asyncResult.status == "failed") {
      switch (asyncResult.error.code) {
        case 12004:
          console.log("Domain is not trusted");
          break;
        case 12005:
          console.log("HTTPS is required");
          break;
        case 12007:
          console.log("A dialog is already opened.");
          break;
        default:
          console.log(asyncResult.error.message);
          break;
      }
    } else {
      dialog = asyncResult.value;
      console.log(dialog);
      /*Messages are sent by developers programatically from the dialog using office.context.ui.messageParent(...)*/
      //dialog.addEventHandler(Office.EventType.DialogMessageReceived, messageHandler);

      /*Events are sent by the platform in response to user actions or errors. For example, the dialog is closed via the 'x' button*/
      //dialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived);
      setTimeout(function() {
        if (dialog != null) {
          dialog.close();
          //inProgress = false;
        }

        if (btnEvent != null) {
          btnEvent.completed();
        }
      }, 1000);
    }
  };
}
