/// <reference path="../App.js" />

(function() {
    "use strict";

    // The Office initialize function must be run each time a new page is loaded
    Office.initialize = function(reason) {
        $(document).ready(function() {
            app.initialize();

            displayItemDetails();


        });
    };

    // Displays the "Subject" and "From" fields, based on the current mail item
    function displayItemDetails() {
        function extractDateTime(headers) {
            if (headers.length <= 1) {
                return null;
            }

            var dateTimeRegEx = /<b>(EnvoyÃ©&nbsp;|Sent):<\/b> ([\s\S]+)/;
            var dateTimeMatch = dateTimeRegEx.exec(headers[1].trim());

            return (dateTimeMatch === null) ? null : dateTimeMatch[2];
        };

        function extractFrom(headers) {
            if (headers.length === 0) {
                return null;
            }

            var fromRegEx = /([a-zA-Z0-9\s]*) \[<a href=\"mailto:([\s\S]+@[\s\S]+)\">mailto:[\s\S]+@[\s\S]+<\/a>\]/;
            var fromMatch = fromRegEx.exec(headers[0].trim());

            return (fromMatch === null) ? null : {
                displayName: fromMatch[1],
                emailAddress: fromMatch[2]
            };
        }

        function messagesToHtml(messages) {
            var msgHtml = messages.map(function(msg, i) {
                return '<div class="bubble ' + (i % 2 === 0 ? "you" : "me") + '">'
                        + '<p class="chat-message">' + msg.body + "</p>"
                        + '<p class="chat-datetime">' + msg.dateTime + "</p>"
                        + '</div>';
            });
            
            return msgHtml.reverse().join();
        };

        var item = Office.cast.item.toItemRead(Office.context.mailbox.item);
        Office.context.mailbox.item.body.getAsync("html", {}, function(results) {
            var fakeDom = $(document.createElement("html"));
            fakeDom.html(results.value);
            var conversation = $("div.WordSection1", fakeDom).html();
            var text = conversation.replace(/<p class="MsoNormal">/g, "")
                .replace(/<\/p>/g, "\n")
                .replace(/<span\b[^>]*>/g, "")
                .replace(/<\/span>/g, "")
                .split(/<b>De&nbsp;:<\/b>|<b>From:<\/b>/);

            var messages = text.map(function(el) {
                var msgBegin = el.split("Sincerely,")[0];
                var splitedMsgBegin = msgBegin.split(/Hi\s[a-zA-Z0-9]*,/g);
                var msgHeaders = splitedMsgBegin[0].split("<br>");
                var msgBody = splitedMsgBegin[1].replace(/&nbsp;/g, " ").trim();

                return {
                    from: extractFrom(msgHeaders),
                    dateTime: extractDateTime(msgHeaders),
                    body: msgBody
                };
            });

            messages[0].from = Office.context.mailbox.item.from;
            messages[0].dateTime = Office.context.mailbox.item.dateTimeCreated.toString();

            $("#chat").html(messagesToHtml(messages));
        });

        var from;
        if (item.itemType === Office.MailboxEnums.ItemType.Message) {
            from = Office.cast.item.toMessageRead(item).from;
        } else if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
            from = Office.cast.item.toAppointmentRead(item).organizer;
        }

        if (from) {
            $("#from").text(from.displayName);
            $("#from").click(function() {
                app.showNotification(from.displayName, from.emailAddress);
            });
        }
    }
})
();