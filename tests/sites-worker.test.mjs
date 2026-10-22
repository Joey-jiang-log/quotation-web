import assert from "node:assert/strict";
import { access } from "node:fs/promises";
import { readFileSync } from "node:fs";
import test from "node:test";
import worker from "../worker/index.js";
import { normalizeRedShellInternal } from "../server/red-shell.mjs";

test("serves existing static assets without a fallback", async () => {
  const calls = [];
  const response = await worker.fetch(new Request("https://example.test/assets/app.js"), {
    ASSETS: {
      fetch: async (request) => {
        calls.push(new URL(request.url).pathname);
        return new Response("asset", { status: 200 });
      },
    },
  });

  assert.equal(response.status, 200);
  assert.deepEqual(calls, ["/assets/app.js"]);
});

test("falls back to index.html for an unknown app route", async () => {
  const calls = [];
  const response = await worker.fetch(
    new Request("https://example.test/flow/step-two?source=share", {
      headers: { accept: "text/html" },
    }),
    {
      ASSETS: {
        fetch: async (request) => {
          const url = new URL(request.url);
          calls.push(url.pathname + url.search);
          return new Response(url.pathname === "/index.html" ? "app" : "missing", {
            status: url.pathname === "/index.html" ? 200 : 404,
          });
        },
      },
    },
  );

  assert.equal(response.status, 200);
  assert.deepEqual(calls, ["/flow/step-two?source=share", "/index.html"]);
});

test("does not turn missing API or write requests into the app shell", async () => {
  for (const request of [
    new Request("https://example.test/api/missing", { headers: { accept: "application/json" } }),
    new Request("https://example.test/flow", { method: "POST", headers: { accept: "text/html" } }),
  ]) {
    let calls = 0;
    const response = await worker.fetch(request, {
      ASSETS: {
        fetch: async () => {
          calls += 1;
          return new Response("missing", { status: 404 });
        },
      },
    });

    assert.equal(response.status, 404);
    assert.equal(calls, 1);
  }
});

test("emits the files required by Sites packaging", async () => {
  await access(new URL("../dist/client/index.html", import.meta.url));
  await access(new URL("../dist/server/index.js", import.meta.url));
  await access(new URL("../dist/.openai/hosting.json", import.meta.url));
});

test("red-shell catalog variants use standard internal code for quote output", () => {
  const red = { commonName: "丰田凯美瑞（红色）", vehicle: "丰田凯美瑞（红色）" };
  const catalog = JSON.parse(readFileSync(new URL("../data/source/产品目录60625.json", import.meta.url), "utf8")).records;
  const realVariant = catalog.find((item) => item.internal === "DQ91084812");
  assert.ok(realVariant, "the red variant exists in the source catalog");
  assert.equal(normalizeRedShellInternal("DQ91084812", realVariant), "DQ910848");
  assert.equal(normalizeRedShellInternal("DQ91084812", red), "DQ910848");
  assert.equal(normalizeRedShellInternal("DQ910848", red), "DQ910848");
  assert.equal(normalizeRedShellInternal("DQ91084812", { commonName: "丰田凯美瑞" }), "DQ91084812");
  assert.equal(normalizeRedShellInternal("DQ91084812", null), "DQ91084812");
  assert.equal(normalizeRedShellInternal("DQ9101AR", { commonName: "红色外壳" }), "DQ9101AR");
});
