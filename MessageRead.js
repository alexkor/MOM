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
            item.body = mailItem.body;
        }
        else {
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
        if (!item.start || !item.end || !item.subject) {
            setTimeout(waitUntilDataRetrive, 200);
        }
        else {
            fillData();
        }
    }

    function fillData() {
        $('#message-props').show();

        //$('#attachments').html(buildAttachmentsString(item.attachments));
        var body = '';
        item.body.getAsync("text", { asyncContext: "callback" }, function (result) { body = result.value; $('#body').html(body) });
        $('#end').text(item.end);
        $('#location').html(item.location);
        $('#normalizedSubject').text(item.subject);
        //$('#notificationMessages').text(item.notificationMessages);

        $('#optionalAttendees').html(buildEmailAddressesString(item.optionalAttendees));
        $('#requiredAttendees').html(buildEmailAddressesString(item.requiredAttendees));

        //$('#organizer').text(buildEmailAddressesString(item.organizer));

        $('#start').val(item.start.format('yyyy-MM-dd'));
        //$('#subject').html(item.subject);

        $('#submit').click(function (ev) {
            var button = $(this);
            button.prop('disabled', true);
            $.ajax({
                url: 'https://confluence.beeline.kz/ajax/confiforms/rest/save.action',
                type: 'POST',
                xhrFields: { withCredentials: true },
                contentType: "application/x-www-form-urlencoded;",
                data: 'pageId=53811457&f=meetingCollector&title01=' + item.subject +
                    '&beginTm=' + item.start.format('dd.MM.yyyy HH:mm') +
                    '&endTm=' + item.end.format('dd.MM.yyyy HH:mm') +
                    '&obligMember=' + item.requiredAttendees.map(function (address) { return address.emailAddress; }) +
                    '&optionalMember=' + item.optionalAttendees.map(function (address) { return address.emailAddress; }) +
                    '&place=' + item.location +
                    '&agenda=' + body +
                    '&type=OutlookConfluence'
            }).fail(function (ctx, state, message) {
                button.prop('disabled', false);
                showNotification('Ошибка создания МОМ встречи', state + ': ' + message);
            }).done(function () {
                //showNotification('МОМ встречи создан', '<a target="_blank" href="https://confluence.beeline.kz/pages/viewpage.action?pageId=53812347">confluence</a>');
                window.open('https://confluence.beeline.kz/pages/viewpage.action?pageId=53812347', '_blank');
            });

            //Office.context.ui.displayDialogAsync('https://office.beeline.kz/sites/system/Lists/TMP/NewForm.aspx?Title=' + item.subject);

            //var sendData = {
            //    pageId: "53811457",
            //    f: "meetingCollector",
            //    title01: item.subject,
            //    beginTm: item.start.format('dd.MM.yyyy HH:mm'),
            //    endTm: item.end.format('dd.MM.yyyy HH:mm'),
            //    obligMember: item.requiredAttendees.map(function (address) { return address.emailAddress; }),
            //    optionalMember: item.optionalAttendees.map(function (address) { return address.emailAddress; }),
            //    place: item.location,
            //    agenda: body,
            //    type: "OutlookConfluence"
            //};
            //console.log(sendData);

            //var siteurl = "https://office.beeline.kz/sites/system";
            //var soapEnv =
            //    '<?xml version="1.0" encoding="utf-8"?> \
            //    <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" \
            //        xmlns:xsd="http://www.w3.org/2001/XMLSchema" \
            //        xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"> \
            //      <soap:Body> \
            //        <UpdateListItems xmlns="http://schemas.microsoft.com/sharepoint/soap/"> \
            //          <listName>TMP</listName> \
            //          <updates> \
            //            <Batch OnError="Continue"> \
            //        <Method ID="1" Cmd="New"> \
            //            <Field Name="Title">'+ item.subject + '</Field> \
            //             <Field Name="Body">'+ body + '</Field> \
            //                  </Method> \
            //    </Batch> </updates> \
            //        </UpdateListItems> \
            //      </soap:Body> \
            //    </soap:Envelope>';

            //$.ajax({
            //    url: siteurl + "/SitePages/Домашняя.aspx",
            //    method: 'GET',
            //    xhrFields: { withCredentials: true },
            //    success: function (data) {
            //        var digest = $('input#__REQUESTDIGEST', $(data)).val();
            //        $.ajax({
            //            url: siteurl + "/_vti_bin/Lists.asmx",
            //            type: "POST",
            //            dataType: "xml",
            //            xhrFields: { withCredentials: true },
            //            contentType: 'text/xml; charset="utf-8"',
            //            headers: {
            //                "X-RequestDigest": digest
            //            },
            //            data: soapEnv,
            //            complete: console.log,
            //            success: console.log,
            //            error: console.log
            //        });
            //    }
            //});
            //$.ajax({
            //    url: 'https://confluence.beeline.kz/ajax/confiforms/rest/save.action',
            //    //url: 'http://localhost:3000',
            //    type: 'POST',
            //    contentType: "application/x-www-form-urlencoded",
            //    data: sendData,
            //}).always(function (data) { $('#result').html(data); });

            //var data = 'https://confluence.beeline.kz/ajax/confiforms/rest/save.action?' +
            //    'pageId=53811457' +
            //    '&f=meetingCollector' +
            //    '&title01=' + item.subject +
            //    '&beginTm=' + item.start.toLocaleString('de-DE') +
            //    '&endTm=' + item.end.toLocaleString('de-DE') +
            //    '&obligMember=' + item.requiredAttendees.map(function (address) { return address.emailAddress; }) +
            //    '&optionalMember=' + item.optionalAttendees +
            //    '&place=' + item.location +
            //    '&agenda=' + encodeURI(body) +
            //    '&type=OutlookConfluence';
            //$.get(data).fail(
            //        function (xhr, status, error) {
            //            console.log(error);
            //            $('#result').html(error);
            //        });
        });
    }

    // Вспомогательная функция для отображения уведомлений
    function showNotification(header, content) {
        $("#notificationHeader").text(header);
        $("#notificationBody").html(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();