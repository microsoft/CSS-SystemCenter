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
    var extraPhrase = ":::template /.templates/Common-Header.md\n:::\n\n::: template /.templates/Sandbox-Header.md\n:::\n\n";

    // Concatenate all the text
    var clipboardText = `${extraPhrase}${bodyText}`;

    // Attempt to use the Clipboard API
    if (navigator.clipboard && window.isSecureContext) {
        navigator.clipboard.writeText(clipboardText).then(function () {
            console.log('Details copied to clipboard!');
        }).catch(function (err) {
            console.error('Failed to copy text: ', err);
            fallbackCopyTextToClipboard(clipboardText);
        });
    } else {
        // Fallback method for older browsers or when Clipboard API is blocked
        fallbackCopyTextToClipboard(clipboardText);
    }
   // openBrowser(); //conflict
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
        console.log(msg);
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


function copyAsPowerShell(){
    var bodyText = document.getElementById('item-body').innerHTML;
    bodyText = bodyText.replace(/<o:p>|<\/o:p>/g, ''); //special case where those elements are injected in order to enable Word to convert the HTML back to fully compatible Word document, with everything preserved.
    //bodyText = bodyText.replace(/(Issue|Issues|Symptom|Symptoms):/g, '<b>$1:</b>'); //special case to add bold for Issue
    bodyText = bodyText.replace(/(Issue|Issues|Symptom|Symptoms):/g, '<p class="x_wordsection1" style="margin:0in"><b>$1:</b>'); //special case to add bold for Issue
    bodyText = bodyText.replace(/color:#\d{6};/g, 'color:black;'); //special case for "color:#171717;"

    //fromat HTML for Word alike e.g. pretty print
    bodyText = formatHTML(bodyText);
    console.log(bodyText);

    // Extra phrase to be added
    var extraPhrase = ":::template /.templates/Common-Header.md\n:::\n\n::: template /.templates/Sandbox-Header.md\n:::\n\n";

    //find Product
    var product = document.getElementById('item-product');
    var selectedValue = product.selectedOptions[0].text;

    //find sender
    var sender = document.getElementById('item-user').innerHTML;

    // Concatenate all the text
    var clipboardText = `
$productName = "${selectedValue}"
$aliasName = "TSG AutoGenerator"
$subPageTitle = "test wiki submitted via ADO API"
$subPageBody =
@'
!!!Attention!!!
    Please review the document and remove any sensitive customer data.
    Page generated with PS by user ${sender}

${extraPhrase}
${bodyText}
'@

##########################################################################################################################
#region Azure CLI is required to get an ADO token
if (-not (get-command az)) {
    $msgStr = "Azure CLI has to be installed. Run the below in an ELEVATED PS window, open a new PS instance and then re-run this script."
    $msgStr += "    $ProgressPreference = 'SilentlyContinue'; Invoke-WebRequest -Uri https://aka.ms/installazurecliwindowsx64 -OutFile .\AzureCLI.msi; Start-Process msiexec.exe -Wait -ArgumentList '/I AzureCLI.msi /quiet'; Remove-Item .\AzureCLI.msi"
    Write - Error $msgStr
    return
}
#endregion
##########################################################################################################################
az login
$accessTokenJson = az account get-access-token | ConvertFrom-Json

$bearerTokenId = $accessTokenJson.accessToken
$adoHeaders = @{
    "Authorization" = "bearer $bearerTokenId"
    "accept"= "application/json;api-version=5.0-preview.1;excludeUrls=true;enumsAsNumbers=true;msDateFormat=true;noArrayWrap=true"
}
$wikiUrl = "https://dev.azure.com/supportability/AzureMonitor/_apis/wiki/wikis/AzureMonitor.wiki/pages?path=/Sandbox/Hybrid%20Sandbox/$productName/$aliasName/$subPageTitle&api-version=7.1-preview.1"

$payload = 
@'
{
    "content": "|||subPageBodyInPayLOAD|||"
}
'@
$payload = $payload.Replace("|||subPageBodyInPayLOAD|||", $subPageBody)
Invoke-RestMethod -Uri $wikiUrl -Method Put -ContentType "application/json" -Headers $adoHeaders -Body $payload
`;

    // Attempt to use the Clipboard API
    if (navigator.clipboard && window.isSecureContext) {
        navigator.clipboard.writeText(clipboardText).then(function () {
            console.log('Details copied to clipboard!');
        }).catch(function (err) {
            console.error('Failed to copy text: ', err);
            fallbackCopyTextToClipboard(clipboardText);
        });
    } else {
        // Fallback method for older browsers or when Clipboard API is blocked
        fallbackCopyTextToClipboard(clipboardText);
    }
}
