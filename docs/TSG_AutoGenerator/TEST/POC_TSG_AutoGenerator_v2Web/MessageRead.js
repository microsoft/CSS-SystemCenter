'use strict';

(function () {

    Office.onReady(function () {
        // Office is ready
        $(document).ready(function () {
            // The document is ready
            loadItemProps(Office.context.mailbox.item);
            getEmailBody(Office.context.mailbox.item);
            getImages();
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
                 const bodyRegex = /(Issue[s]?|Symptom[s]?):[\s\S]*?(?=\bBest Regards\b)/i;
                 //const bodyRegex = /[\s\S]*?(?=\bBest Regards)\b/i;
                 const bodyMatch = body.match(bodyRegex);
                 const bodyText = bodyMatch ? bodyMatch[0].trim() : "No Body found";



                 $('#item-body').html(bodyText); 
             } else {
                 console.log(bodyResult.error.message);
             }
         });
     }
     
    function getImages() {
        const tempDiv = document.getElementById('item-body').innerHTML;
        // Extract images and convert to Base64
        const images = tempDiv.getElementsByTagName('img');
        Array.from(images).forEach((img, index) => {
            const cid = img.getAttribute('src');
            const altText = img.getAttribute('alt');

            // Convert image src to Base64 (assuming it is a CID)
            fetch(cid)
                .then(response => response.blob())
                .then(blob => {
                    const reader = new FileReader();
                    reader.onloadend = function () {
                        const base64data = reader.result;

                        // Display Base64 string
                        const imgInfo = document.createElement('div');
                        imgInfo.innerHTML = `<p>Image ${index + 1}:</p>
                                                 <p><strong>CID:</strong> ${cid}</p>
                                                 <p><strong>Alt Text:</strong> ${altText}</p>
                                                 <textarea style="width:100%;height:100px;">${base64data}</textarea>`;
                        imagesContainer.appendChild(imgInfo);
                    }
                    reader.readAsDataURL(blob);
                });
        });

    }

})();

