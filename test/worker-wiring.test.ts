import { test, expect } from "bun:test";
import { FunctionRegistry, ReadOnlyCatalogInterface } from "@query-farm/vgi";
import { makeFunctions } from "../src/functions.js";
import { makeCatalog, AZURE_GRAPH_SECRET } from "../src/catalog.js";
import { FakeGraphFiles } from "./fake-files.js";

test("both content functions register and the azure catalog advertises them", () => {
  const g = new FakeGraphFiles();
  const clientFactory = () => ({ fetchJson: g.fetch, postJson: async () => ({}) });

  const functions = makeFunctions(clientFactory);
  expect(functions.length).toBe(2);

  const registry = new FunctionRegistry();
  for (const f of functions) registry.register(f);

  const cat = makeCatalog(functions);
  expect(cat.name).toBe("azure");
  expect(cat.secretTypes?.[0]).toBe(AZURE_GRAPH_SECRET);
  expect(cat.schemas[0]!.functions!.map((f) => (f as { meta: { name: string } }).meta.name).sort()).toEqual([
    "drive_items",
    "sharepoint_list_items",
  ]);

  // Constructs the read-only catalog interface without throwing.
  new ReadOnlyCatalogInterface(cat, registry);
});
