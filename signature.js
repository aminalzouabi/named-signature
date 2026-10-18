Office.onReady();

function onNewMessageComposeHandler(event) {
    insertSignature(event);
}

// One entry per sub-brand, keyed by the sender's email domain.
// "site" is the website address without "www." (added in code below).
const BRANDS = {
    "slash.ae": {
        name: "Slash",
        tagline: "Strategic Design Studio",
        site: "slash.ae"
    },
    "fabrica.ae": {
        name: "Fabrica",
        tagline: "Interdisciplinary Making Platform",
        site: "fabrica.ae"
    },
    "fountconcept.ae": {
        name: "Fount",
        tagline: "Curated Concept Store",
        site: "fountconcept.ae"
    },
    "central.ae": {
        name: "Central",
        tagline: "Logistics Platform",
        site: "central.ae"
    },
    "basearchitecture.ae": {
        name: "Base",
        tagline: "Collaborative Architecture Studio",
        site: "basearchitecture.ae"
    }
};

// Internal CSS is used because setSignatureAsync does not support inline CSS.
// Scope every rule to the signature so the message and quoted replies are unaffected.
// Helvetica must be installed on the viewing device; Arial is the fallback.
const SIGNATURE_STYLE = `<style type="text/css">
.rc-signature, .rc-signature span, .rc-signature a, .rc-signature strong, .rc-signature font {
    font-family: Helvetica, Arial, sans-serif !important;
    font-size: 10pt !important;
    line-height: 14pt !important;
    color: #000000;
}
.rc-signature a { text-decoration: underline; }
</style>`;

function escapeHtml(text) {
    return String(text || "").replace(/&/g, "&amp;").replace(/</g, "&lt;")
        .replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#39;");
}

function getBrand(email) {
    const domain = (String(email || "").trim().split("@")[1] || "").toLowerCase();
    return Object.prototype.hasOwnProperty.call(BRANDS, domain) ? BRANDS[domain] : null;
}

function makeLink(site) {
    const host = "www." + site;
    return '<a class="rc-signature-text" href="https://' + host + '">' +
        '<font face="Helvetica, Arial, sans-serif">' + host + '</font></a>';
}

function getSignatureHtml(name, email) {
    const brand = getBrand(email);
    const text = value => '<span class="rc-signature-text"><font face="Helvetica, Arial, sans-serif">' +
        escapeHtml(value) + '</font></span>';
    const bold = value => '<strong>' + text(value) + '</strong>';
    const lines = [
        text("Best,"), "", bold(name), "", text("–"),
        bold(brand ? brand.name : "Ripple Collective"),
        text(brand ? brand.tagline : "Multidisciplinary Design Ecosystem"),
        makeLink(brand ? brand.site : "ripplecollective.ae"), "",
        text(brand
            ? brand.name + " is part of Ripple Collective, a Multidisciplinary Design Ecosystem, based in Abu Dhabi."
            : "Ripple Collective is an ecosystem of projects, including Slash, Fabrica, Fount, Central, and Base, all based in Abu Dhabi.")
    ];
    if (brand) lines.push(makeLink("ripplecollective.ae"));
    return SIGNATURE_STYLE + '<br><br><div class="rc-signature">' + lines.join("<br>") + "</div>";
}

function getSignatureText(name, email) {
    const brand = getBrand(email);
    return ["", "", "Best,", "", name, "", "–",
        brand ? brand.name : "Ripple Collective",
        brand ? brand.tagline : "Multidisciplinary Design Ecosystem",
        "www." + (brand ? brand.site : "ripplecollective.ae"), "",
        brand
            ? brand.name + " is part of Ripple Collective, a Multidisciplinary Design Ecosystem, based in Abu Dhabi.\nwww.ripplecollective.ae"
            : "Ripple Collective is an ecosystem of projects, including Slash, Fabrica, Fount, Central, and Base, all based in Abu Dhabi."
    ].join("\n");
}

function insertSignature(event) {
    let completed = false;
    function finish() {
        if (!completed) {
            completed = true;
            event.completed();
        }
    }
    function fail(error) {
        console.error("Signature insertion failed:", error && error.message ? error.message : error);
        finish();
    }
    try {
        const profile = Office.context.mailbox.userProfile;
        const name = profile.displayName || "";
        const email = profile.emailAddress || "";
        const body = Office.context.mailbox.item.body;
        body.getTypeAsync(function (formatResult) {
            try {
                if (formatResult.status === Office.AsyncResultStatus.Failed) {
                    fail(formatResult.error);
                    return;
                }
                const isHtml = formatResult.value === Office.CoercionType.Html;
                const signature = isHtml ? getSignatureHtml(name, email) : getSignatureText(name, email);
                body.setSignatureAsync(signature, {
                    coercionType: isHtml ? Office.CoercionType.Html : Office.CoercionType.Text
                }, function (result) {
                    if (result.status === Office.AsyncResultStatus.Failed) {
                        fail(result.error);
                        return;
                    }
                    finish();
                });
            } catch (error) {
                fail(error);
            }
        });
    } catch (error) {
        fail(error);
    }
}

Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
