// const fetch = require('node-fetch'); // Native fetch in Node 18+ 
// Actually, modern node has fetch.

const GAS_BASE = "https://script.google.com/macros/s/AKfycbx1u3qfMh7GxCZ6jMa2h3m2Q296w9ZgV3V8pKuWdXyop4r8TVocDS4eAP_lUKP16Jnq6A/exec";

async function testGas() {
    try {
        console.log("Testing GET...");
        const resGet = await fetch(GAS_BASE);
        const textGet = await resGet.text();
        console.log("GET Response:", textGet);

        console.log("Testing POST with action: getData...");
        const resPost = await fetch(GAS_BASE, {
            method: "POST",
            body: JSON.stringify({ action: "getData" }),
            headers: { "Content-Type": "application/json" } // GAS might ignore this but good to try
        });
        const textPost = await resPost.text();
        console.log("POST Response:", textPost);

    } catch (e) {
        console.error("Error:", e);
    }
}

testGas();
