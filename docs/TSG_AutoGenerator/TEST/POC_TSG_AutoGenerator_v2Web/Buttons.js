function copyToClipboard() {
    var bodyText = document.getElementById('item-body').innerHTML;
    bodyText = bodyText.replace(/<o:p>|<\/o:p>/g, ''); //special case where those elements are injected in order to enable Word to convert the HTML back to fully compatible Word document, with everything preserved.
    //bodyText = bodyText.replace(/(Issue|Issues|Symptom|Symptoms):/g, '<b>$1:</b>'); //special case to add bold for Issue
    bodyText = bodyText.replace(/(Issue|Issues|Symptom|Symptoms):/g, '<p class="x_wordsection1" style="margin:0in"><b>$1:</b>'); //special case to add bold for Issue
    bodyText = bodyText.replace(/color:#\d{6};/g, 'color:black;'); //special case for "color:#171717;"

    //fromat HTML for Word alike e.g. pretty print
    bodyText = formatHTML(bodyText);
    console.log(bodyText);

    // Extra phrase to be added
    const extraPhrase = ":::template /.templates/Common-Header.md\n:::\n\n::: template /.templates/Sandbox-Header.md\n:::\n\n";

    // Concatenate all the text
    const clipboardText = `${extraPhrase}${bodyText}`;

    openBrowser();


    // Attempt to use the Clipboard API
    if (navigator.clipboard && window.isSecureContext) {
        navigator.clipboard.writeText(clipboardText).then(function () {
            alert('Details copied to clipboard!');
        }).catch(function (err) {
            console.error('Failed to copy text: ', err);
            fallbackCopyTextToClipboard(clipboardText);
        });
    } else {
        // Fallback method for older browsers or when Clipboard API is blocked
        fallbackCopyTextToClipboard(clipboardText);
    }

}



function formatHTML(html) {
    const tab = '  '; // Indentation
    let formatted = '', indent = '';

    html.split(/>\s*</).forEach(function (element) {
        if (element.match(/^\/\w/)) {
            // Closing tag
            indent = indent.substring(tab.length);
        }
        formatted += indent + '<' + element + '>\n';
        if (element.match(/^<?\w[^>]*[^\/]$/)) {
            // Opening tag
            indent += tab;
        }
    });

    return formatted.substring(1, formatted.length - 2);
}


function fallbackCopyTextToClipboard(text) {
    // Create a temporary textarea element
    const textArea = document.createElement("textarea");
    textArea.value = text;

    // Avoid scrolling to bottom of page in MS Edge
    textArea.style.position = "fixed";
    textArea.style.top = 0;
    textArea.style.left = 0;

    document.body.appendChild(textArea);
    textArea.focus();
    textArea.select();

    try {
        const successful = document.execCommand('copy');
        const msg = successful ? 'Details copied to clipboard!' : 'Failed to copy text';
        alert(msg);
    } catch (err) {
        console.error('Fallback: Failed to copy text: ', err);
    }

    // Remove the temporary textarea
    document.body.removeChild(textArea);
}


function openBrowser() {

    //TODO for user
    var product = document.getElementById('item-product');
    var selectedValue = product.selectedOptions[0].text;

    if (selectedValue === "SCOM") {
        window.open("https://dev.azure.com/Supportability/AzureMonitor/_wiki/wikis/AzureMonitor.wiki/1446664/SCOM", "_blank");
    }
    if (selectedValue === "SCOM MI") {
        window.open("https://dev.azure.com/Supportability/AzureMonitor/_wiki/wikis/AzureMonitor.wiki/1565150/SCOMMI", "_blank");
    }
    if (selectedValue === "SCSM") {
        window.open("https://dev.azure.com/Supportability/AzureMonitor/_wiki/wikis/AzureMonitor.wiki/1446666/SCSM", "_blank");
    }
    if (selectedValue === "SCORCH") {
        window.open("https://dev.azure.com/Supportability/AzureMonitor/_wiki/wikis/AzureMonitor.wiki/1446668/SCORCH", "_blank");
    }
    if (selectedValue === "VMM") {
        window.open("https://dev.azure.com/Supportability/AzureMonitor/_wiki/wikis/AzureMonitor.wiki/1446670/SCVMM", "_blank");
    }
    else {
        window.open("https://dev.azure.com/Supportability/AzureMonitor/_wiki/wikis/AzureMonitor.wiki/1446663/Hybrid-Sandbox", "_blank");
    }   
}

window.onload = function () {
    document.querySelector("button").addEventListener("click", copyToClipboard);
};
