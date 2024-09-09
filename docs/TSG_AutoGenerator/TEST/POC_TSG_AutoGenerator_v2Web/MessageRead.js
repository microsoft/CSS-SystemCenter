'use strict';

(function () {

    Office.onReady(function () {
        // Office is ready
        $(document).ready(function () {
            // The document is ready
            loadItemProps(Office.context.mailbox.item);
            getEmailBody(Office.context.mailbox.item);
        });
    });

    function loadItemProps(item) {  
        $('#item-subject').text(item.subject);
        /*
        var subject = item.subject;
        const subjectRegex = /\b(SCSM|ORCH|SCOM|VMM|Service Manager|Orchestrator|Operations Manager|Virtual Machine Manager)\b/g;
        var subjectText = subject.match(subjectRegex);
        $('#item-product').text(subjectText);
        */
        $('#item-user').html(item.from.displayName + " &lt;" + item.from.emailAddress + "&gt;<br>");

    }

     function getEmailBody(item) {
         // Get the current body of the message
         item.body.getAsync(Office.CoercionType.Html, (bodyResult) => {
             if (bodyResult.status === Office.AsyncResultStatus.Succeeded) {
                 var body = bodyResult.value;
                 const bodyRegex = /(Issue[s]?|Symptom[s]?):[\s\S]*?(?=\b(Best Regards|Kind Regards|Thank you,)\b)/i;
                 //v2 const bodyRegex = /(Issue[s]?|Symptom[s]?):[\s\S]*?(?=\bBest Regards\b)/i;
                 //v1 const bodyRegex = /[\s\S]*?(?=\bBest Regards)\b/i;
                 const bodyMatch = body.match(bodyRegex);
                 const bodyText = bodyMatch ? bodyMatch[0].trim() : "No Body found";
                 $('#item-body').html(bodyText); 
             } else {
                 console.log(bodyResult.error.message);
             }
         });
     }

})();

