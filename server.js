// npm i office-addin-dev-certs ; node server.js  -> https://localhost:3000
const https = require("https"), fs = require("fs"), path = require("path");
const certs = require("office-addin-dev-certs");
const types = { ".html": "text/html", ".js": "text/javascript", ".json": "application/json" };
certs.getHttpsServerOptions().then((opts) => {
  https.createServer(opts, (req, res) => {
    const f = path.join(__dirname, req.url.split("?")[0] === "/" ? "taskpane.html" : req.url.split("?")[0]);
    if (!f.startsWith(__dirname) || !fs.existsSync(f)) { res.writeHead(404); return res.end(); }
    res.writeHead(200, { "Content-Type": types[path.extname(f)] || "text/plain", "Access-Control-Allow-Origin": "*" });
    fs.createReadStream(f).pipe(res);
  }).listen(3000, () => console.log("https://localhost:3000"));
});
