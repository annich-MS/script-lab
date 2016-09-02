import {Utilities} from './utilities';

export class ExpectedError {
    // Placeholder class just to indicate that the error was in fact an expected rejection.
}

/**
 * A class for signifying that an error is a "handleable" error that comes from the playground,
 * as opposed to an error that comes from some internal operation or runtime error.
 */
export class PlaygroundError {
    constructor(public message: string | string[]) { }
}

export class UxUtil {
    static showErrorNotification(title, messageOrMessageArray: string | string[], e: any): Promise<string> {
        if (e instanceof ExpectedError) {
            return;
        }

        if (!messageOrMessageArray) {
            messageOrMessageArray = [];
        }

        var messages: string[] = UxUtil.getArrayOfMessages(messageOrMessageArray);

        if (e) {
            var errorDataExtracted = UxUtil.extractErrorMessage(e);
            if (_.isArray(errorDataExtracted)) {
                errorDataExtracted.forEach((msg) => {
                    messages.push(msg);
                })
            } else {
                messages.push(errorDataExtracted);
            }
        }

        console.log(Utilities.stringifyPlusPlus(messages));
        console.log(Utilities.stringifyPlusPlus(e));

        return UxUtil.showDialog(title, messages, "OK");
    }

    static catchError(title, messageOrMessageArray: string | string[]): (e: Error) => Promise<string> {
        return (e: Error) => UxUtil.showErrorNotification(title, messageOrMessageArray, e);
    }

    static extractErrorMessage(e: any): string | string[] {
        if (e instanceof Error || e instanceof PlaygroundError) {
            return e.message;
        } else {
            return e;
        }
    }

    static showDialog(title: string,
        messageOrMessageArray: string | string[],
        buttons: string[] | string
    ): Promise<string> {
        return new Promise(function (resolve) {
            $(document).ready(function () {
                console.log(title, messageOrMessageArray, buttons);
                var $dialogRoot = $('.ms-Dialog--lgHeader');
                $dialogRoot.find('.ms-Dialog-title').text(title);

                var $dialogContent = $dialogRoot.find('.ms-DialogContent').empty();

                UxUtil.getArrayOfMessages(messageOrMessageArray).forEach((message) => {
                    var p = $('p').text(message).addClass('ms-Dialog-subText');
                    $dialogContent.append(p);
                })

                var buttonsArray: string[];
                if (_.isArray(buttons)) {
                    buttonsArray = buttons;
                } else {
                    buttonsArray = [buttons];
                }

                var $buttonPane = $('.ms-Dialog-actionsRight', $dialogRoot);
                $buttonPane.empty();
                buttonsArray.forEach(function (buttonLabel, i) {
                    var $button = $('<button class="ms-Dialog-action ms-Button"><span class="ms-Button-label">' + buttonLabel + '</span></button>');

                    if (i == 0) {
                        $button.addClass('ms-Button--primary');
                    }

                    $button.click(function () {
                        $dialogRoot.hide();
                        resolve(buttonLabel);
                    });
                    $buttonPane.append($button);
                });

                $dialogRoot.show();
            })
        });
    }

    static getArrayOfMessages(messageOrMessageArray: string | string[]) {
        var messages: string[];
        if (_.isArray(messageOrMessageArray)) {
            messages = messageOrMessageArray;
        } else {
            messages = [messageOrMessageArray];
        }

        var result = [];
        messages.map((message) => Utilities.stringOrEmpty(message).split('\n')).forEach((messageArray) => {
            messageArray.forEach((message) => {
                result.push(message);
            })
        })

        return result;
    }

    static debugAlert(msg) {
        return UxUtil.showDialog("Alert!", msg, "OK");
    }
}