/** Configure unchanged Git-pinned worker functions with their existing client seam.
 * Authentication is outside this fixture: only loopback HTTP GETs are made.
 * The real functions still construct URLs, follow pages, decode changes, and
 * emit their actual Arrow schema and per-resource delta marker through VGI.
 */
import { FunctionRegistry, ReadOnlyCatalogInterface, Worker } from "@query-farm/vgi";
// launch_fixture.sh copies this adapter beside the immutable remote source.
import { makeFunctions } from "./src/functions.ts";
import { makeCatalog } from "./src/catalog.ts";

const rawBase = process.env.NOTEBOOK_AZURE_FILES_FIXTURE_URL;
if (!rawBase) throw new Error("Start the notebook's synthetic Graph fixture first");
const base = new URL(rawBase);
if (base.protocol !== "http:" || base.hostname !== "127.0.0.1") {
  throw new Error("This launcher accepts only an HTTP loopback fixture");
}
const clientFactory = () => ({
  postJson: async () => { throw new Error("This fixture supports read-only GETs"); },
  fetchJson: async (url: string): Promise<Record<string, unknown>> => {
    const original = new URL(url);
    if (original.origin !== "https://graph.microsoft.com") {
      throw new Error("Unexpected upstream origin in the pinned files function");
    }
    const local = new URL(original.pathname + original.search, base);
    const response = await fetch(local);
    if (!response.ok) throw new Error(`Synthetic fixture: ${response.status} ${await response.text()}`);
    return await response.json() as Record<string, unknown>;
  },
});
const functions = makeFunctions(clientFactory);
const registry = new FunctionRegistry();
for (const fn of functions) registry.register(fn);
const catalogInterface = new ReadOnlyCatalogInterface(makeCatalog(functions), registry);
new Worker({ functions, catalogInterface }).run();
