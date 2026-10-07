const { spawn } = require("child_process"), fs = require("fs");
const [S, CAT] = process.argv.slice(2);
const jobs = fs.readFileSync(__dirname + "/jobs.tsv", "utf8").trim().split("\n").map(l => l.split("\t"));
let next = 0, running = 0;
function go() {
  while (running < 4 && next < jobs.length) {
    const [mode, i, f, pages] = jobs[next++]; running++;
    const env = Object.assign({}, process.env); if (mode === "old") env.OCR_NO_SKEW = "1";
    const t0 = Date.now();
    spawn(process.execPath, [__dirname + "/run1.js", CAT, f, pages, `${S}/reg/${mode}-${i}.json`], { env, stdio: "ignore" })
      .on("exit", c => { running--; console.log(`done ${mode} ${i} code=${c} ${((Date.now() - t0) / 1000) | 0}s`); if (next >= jobs.length && !running) console.log("ALL DONE"); go(); });
  }
}
go();
