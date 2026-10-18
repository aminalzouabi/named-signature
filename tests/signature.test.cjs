const fs = require("node:fs");
const vm = require("node:vm");
const assert = require("node:assert/strict");
const source = fs.readFileSync(require("node:path").join(__dirname, "..", "signature.js"), "utf8");
function setup(options = {}) {
    let completed = 0, inserted, logged = 0;
    const handlers = {};
    const Office = {
        onReady() {}, actions: { associate(name, handler) { handlers[name] = handler; } },
        CoercionType: { Html: "html", Text: "text" },
        AsyncResultStatus: { Failed: "failed", Succeeded: "succeeded" },
        context: { mailbox: {
            userProfile: { displayName: "A <B> & C", emailAddress: options.email || "a@slash.ae" },
            item: { body: {
                getTypeAsync(callback) {
                    if (options.typeThrow) throw new Error("getType failure");
                    callback(options.typeFail ? {status: "failed", error: {message: "format error"}} :
                        {status: "succeeded", value: options.plain ? "text" : "html"});
                },
                setSignatureAsync(data, config, callback) {
                    if (options.setThrow) throw new Error("setSignature failure");
                    inserted = {data, config};
                    callback(options.setFail ? {status: "failed", error: {message: "insert error"}} :
                        {status: "succeeded"});
                }
            }}
        }}
    };
    const context = vm.createContext({Office, console: {error() { logged++; }}});
    vm.runInContext(source, context);
    handlers.onNewMessageComposeHandler({completed() { completed++; }});
    return {completed, inserted, logged};
}
for (const [domain, brand] of [
    ["slash.ae", "Slash"], ["fabrica.ae", "Fabrica"], ["fountconcept.ae", "Fount"],
    ["central.ae", "Central"], ["basearchitecture.ae", "Base"],
    ["ripplecollective.ae", "Ripple Collective"], ["other.example", "Ripple Collective"],
    ["SLASH.AE", "Slash"]
]) {
    const html = setup({email: "a@" + domain});
    assert.equal(html.completed, 1);
    assert(html.inserted.data.includes(">" + brand + "</font>"));
    assert(html.inserted.data.includes("A &lt;B&gt; &amp; C"));
    assert(html.inserted.data.includes("font-size: 10pt !important"));
    assert(!html.inserted.data.includes('style="'));
    assert.equal(html.inserted.config.coercionType, "html");
    const plain = setup({email: "a@" + domain, plain: true});
    assert.equal(plain.completed, 1);
    assert.equal(plain.inserted.config.coercionType, "text");
    assert(plain.inserted.data.includes(brand));
    assert(!plain.inserted.data.includes("<style"));
}
for (const flag of ["typeThrow", "typeFail", "setThrow", "setFail"]) {
    const result = setup({[flag]: true});
    assert.equal(result.completed, 1, flag);
    assert.equal(result.logged, 1, flag);
}
console.log("Passed: 8 brand cases in HTML and plain text, name escaping, scoped 10pt CSS, and 4 error paths.");
