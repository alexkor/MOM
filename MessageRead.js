(function () {
    "use strict";

    var messageBanner;

    // Функцию инициализации Office необходимо выполнять при каждой загрузке новой страницы.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            var element = document.querySelector('.MessageBanner');
            messageBanner = new components.MessageBanner(element);
            messageBanner.hideBanner();
            loadProps();
        });
    };

    // Взять массив объектов AttachmentDetails и создать список имен вложений, разделенных разрывом строки.
    function buildAttachmentsString(attachments) {
        if (attachments && attachments.length > 0) {
            var returnString = "";

            for (var i = 0; i < attachments.length; i++) {
                if (i > 0) {
                    returnString = returnString + "<br/>";
                }
                returnString = returnString + attachments[i].name;
            }

            return returnString;
        }

        return "None";
    }

    // Форматировать объект EmailAddressDetails как
    // Имя Фамилия <emailaddress>
    function buildEmailAddressString(address) {
        return "<a href='" + address.emailAddress + "'>" + address.displayName + "</a>";
    }

    // Взять массив объектов AttachmentDetails и
    // создать список форматированных строк, разделенных разрывом строки
    function buildEmailAddressesString(addresses) {
        if (addresses && addresses.length > 0) {
            var returnString = "";

            for (var i = 0; i < addresses.length; i++) {
                if (i > 0) {
                    returnString = returnString + "<br />";
                }
                returnString = returnString + buildEmailAddressString(addresses[i]);
            }

            return returnString;
        }

        return "None";
    }

    // Загрузите свойства из базового объекта Item, затем загрузите
    // свойства конкретного сообщения.
    function loadProps() {
        var mailItem = Office.context.mailbox.item;
        window.item = {};
        if (!mailItem.itemClass) {
            showNotification('load data if');
            Office.context.mailbox.item.organizer.getAsync(function (asyncResult) {
                item.organizer = asyncResult.value;
            });
            Office.context.mailbox.item.start.getAsync(function (asyncResult) {
                item.start = asyncResult.value;
            });
            Office.context.mailbox.item.end.getAsync(function (asyncResult) {
                item.end = asyncResult.value;
            });
            Office.context.mailbox.item.location.getAsync(function (asyncResult) {
                item.location = asyncResult.value;
            });
            Office.context.mailbox.item.subject.getAsync(function (asyncResult) {
                item.subject = asyncResult.value;
            });
            Office.context.mailbox.item.requiredAttendees.getAsync(function (asyncResult) {
                item.requiredAttendees = asyncResult.value;
            });
            Office.context.mailbox.item.optionalAttendees.getAsync(function (asyncResult) {
                item.optionalAttendees = asyncResult.value;
            });
            item.body.getAsync("text", { asyncContext: "callback" }, function (result) {
                item.body = result.value;

            });
            //item.body = mailItem.body;
        }
        else {
            showNotification('load data else');
            item = mailItem;
            $('#dateTimeCreated').text(item.dateTimeCreated.toLocaleString());
            $('#dateTimeModified').text(item.dateTimeModified.toLocaleString());
            $('#itemClass').text(item.itemClass);
            $('#itemId').text(item.itemId);
            $('#itemType').text(item.itemType);
        }
        waitUntilDataRetrive();
    }

    function waitUntilDataRetrive() {
        showNotification('wait start');
        if (!item.start || !item.end || !item.subject) {
            setTimeout(waitUntilDataRetrive, 200);
        }
        else {
            showNotification('wait end');
            fillData();
        }
    }

    function fillData() {
        $('#message-props').show();

        //$('#start').val(item.start.format('yyyy-MM-dd'));
        //$('#end').text(item.end);
        $('#location').html(item.location);
        $('#normalizedSubject').text(item.subject);
        //$('#optionalAttendees').html(buildEmailAddressesString(item.optionalAttendees));
        //$('#requiredAttendees').html(buildEmailAddressesString(item.requiredAttendees));
        //$('#body').html(item.body);

        $('#submit').click(function (ev) {
            var button = $(this);
            button.prop('disabled', true);
            var message = 'pageId=53811457&f=meetingCollector&title01=' + item.subject +
                '&beginTm=' + item.start.format('dd.MM.yyyy HH:mm') +
                '&endTm=' + item.end.format('dd.MM.yyyy HH:mm') +
                '&obligMember=' + item.requiredAttendees.map(function (address) { return address.emailAddress; }) +
                '&optionalMember=' + item.optionalAttendees.map(function (address) { return address.emailAddress; }) +
                '&place=' + item.location +
                '&agenda=' + item.body +
                '&type=OutlookConfluence' +
                '&authorMeeting=' + item.organizer.emailAddress;
            $.ajax({
                url: 'https://confluence.beeline.kz/ajax/confiforms/rest/save.action',
                type: 'POST',
                headers: { "Authorization": "Basic " + btoa("tech_outlook_mom:~F4B?#?Z") },
                contentType: "application/x-www-form-urlencoded;",
                data: message,
                success: function (data) {
                    var jsonData;
                    try {
                        jsonData = JSON.parse(data);
                    }
                    catch (ex) {
                        // showNotification("Необходима авторизация в confluence, подтвердить переход?",
                        //     '<button onclick="auth()" class="ms-Button ms-Button--primary ms-sm6"><span class="ms-Button-label">Перейти</span></button>' +
                        //     '<button onclick="hideNotification()" class="ms-Button ms-Button--primary ms-sm5"><span class="ms-Button-label">Отмена</span></button>');

                        button.prop('disabled', false);
                        return;
                    }

                    var rId = jsonData.id;
                    $.ajax({
                        url: 'https://confluence.beeline.kz/ajax/confiforms/rest/filter.action',
                        type: 'GET',
                        headers: { "Authorization": "Basic " + btoa("tech_outlook_mom:~F4B?#?Z") },
                        contentType: "application/x-www-form-urlencoded;",
                        data: 'pageId=53811457&f=meetingCollector&q=id:' + rId,
                        success: function (jsonData) {
                            var pId = jsonData.list.entry[0].fields.meetingLink;
                            window.open('https://confluence.beeline.kz/pages/viewpage.action?pageId=' + pId, '_blank');
                        },
                        error: function (ctx, state, message) {
                            button.prop('disabled', false);
                            // showNotification('Ошибка создания МОМ встречи', state + ': ' + message);
                        }
                    });
                },
                error: function (ctx, state, message) {
                    button.prop('disabled', false);
                    // showNotification('Ошибка создания МОМ встречи', state + ': ' + message);
                }
            });
        });
    }

    // Вспомогательная функция для отображения уведомлений
    function showNotification(header, content) {
        $("#notificationHeader").text(header);
        $("#notificationBody").html(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }

    window.auth = function () {
        var url = new URL('./Redirect.html', window.location.href).href;
        var dialogOptions = { width: 20, height: 40, displayInIframe: false, promptBeforeOpen: false };
        Office.context.ui.displayDialogAsync(url, dialogOptions, function (result) {
            console.log(result);
            setTimeout(function () {
                messageBanner.hideBanner();
                result.value.close();
            }, 5000);
        });
    }

    window.hideNotification = function () {
        messageBanner.hideBanner();
    }

})();
