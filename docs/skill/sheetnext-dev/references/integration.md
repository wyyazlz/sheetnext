# SheetNext Integration Notes

Maintained in `docs/static-references/integration.md` and copied into the skill by `npm run docs`. For API source version and fingerprint, see [the API index](core-api.md).

## Distribution And Version

Identify the application's installed package version, lockfile and distribution before selecting an API. The generated source version is not a claim that the same version is published on npm. Pin JS, CSS and locale assets to the same released version; use the matching documentation from that release instead of assuming GitHub `master` matches an older package.

| Capability | Open Edition | Commercial distribution |
| --- | --- | --- |
| Workbook, cells, formulas, drawings, tables, comments, validation and printing | Available | Available |
| JSON object persistence: `SN.IO.getData()` / `setData()` | Available | Available |
| File import/export: `import`, `importFromUrl`, `export`, `exportAllImage` | Disabled | Requires the applicable license |
| Built-in AI runtime | Disabled | Requires the applicable license and your relay |
| PivotTable authoring and processing | Disabled | Requires the applicable license |

Open Edition has disabled implementations for restricted APIs. Adding a license key to that distribution does not supply the omitted implementation; use the commercial distribution when those capabilities are needed. Full source API documentation may include both editions.

Commercial builds may offer limited trial import/export behavior before activation; production entitlements and limits follow the actual delivery and license.

## Initialize

```js
import SheetNext from 'sheetnext';
import 'sheetnext/dist/sheetnext.css';
import zhCN from 'sheetnext/locales/zh-CN';

SheetNext.registerLocale('zh-CN', zhCN);
const container = document.querySelector('.sheetnext-container');
const SN = new SheetNext(container, { locale: 'zh-CN' });
```

Provide a mounted container with explicit width and height. The exported locale subpath is `sheetnext/locales/zh-CN`, without `.js`. Each component should use its own element ref and workbook variable. Initialize only in the browser, after mounting, when using SSR frameworks.

Write cell values and formulas through `editVal`. Its getter returns editing text; use `calcVal` for typed or calculated values and `showVal` for formatted display. A `.value` assignment does not use the cell editing API.

## Save And Restore

Persist the full result of `await SN.IO.getData()`. Do not reconstruct advanced workbook structures from partial cell data. See [JSON format](json-format.md#replacement-and-recovery) before replacing the workbook with `setData`.

File `import` and `importFromUrl` currently return `Promise<void>` in the commercial implementation. Some errors are caught internally and shown as editor notifications, so a fulfilled promise is not a reliable success result. Keep saving paused while the user reviews an imported file, then explicitly confirm persistence. An unattended importer needs a reliable result contract implemented in the editor first; do not invent one in the integration.

For automatic saving:

- Load the initial workbook before attaching persistence handlers. Pause persistence during replacement or recovery, and do not resume after a partial load failure until recovery is resolved.
- Debounce frequent edits, then send snapshots through one in-flight save queue. Record a dirty sequence before collecting each snapshot; edits made during the request must trigger a subsequent snapshot.
- Check HTTP status and the application response. Keep dirty state on any failure. Surface authentication, conflict and network errors to the user; avoid an unbounded retry loop.
- Use a server-issued revision and an atomic compare-and-update on every save. Client queuing does not prevent another tab or user from overwriting data. On a conflict, retain unsaved data and resolve it before resuming; do not adopt the latest revision and blindly retry.
- Use `afterOperation` for operations routed through the operation layer, and supplement it with `afterUndo`, `afterRedo`, sheet/workbook rename and structure events as needed. `afterOperation` is not a universal workbook-changed event, and its `summary` is not a persisted JSON patch contract.
- Call your application's `markDirty()` after changes that have no complete event coverage, such as direct drawing configuration or some annotation/protection changes. Explicitly mark a user file import dirty after the result is confirmed. Reading or selecting cells should not itself trigger a save.
- Flush pending changes before controlled navigation, then remove only your own handlers with `off(eventName, handler)`. A page unload cannot reliably complete asynchronous saving; expose unsaved state and offer a recovery draft when appropriate.

The Chinese [quick integration guide](https://www.sheetnext.com/doc) includes a serial save queue example and its server revision contract. It is an application integration pattern, not a built-in SheetNext persistence service.

## Lifecycle

The current documented workbook API does not expose `destroy()` or `dispose()`. Do not invent either method. Removing a component container and external subscriptions does not dispose the editor's internal document/window listeners and observers. Reuse the instance for a stable component lifetime; for repeated mount/unmount flows, use a separately managed iframe lifetime or implement an explicit disposal lifecycle in the editor before promising cleanup.

`SN.Event.offAll()` removes editor-owned protection and other internal listeners too. Detach individual application handlers instead.

## Permission And AI Boundaries

Workbook `readOnly`, sheet protection and cancelable UI hooks are editing behavior, not a server-side authorization boundary. Validate workbook ownership and write permission on the server. Imported JSON can replace read-only/protection state; reapply application permissions after restoration.

For the built-in AI runtime, read [the relay contract](ai-relay.md#security-and-execution-boundary). The current `Query` and `Edit` tools execute JavaScript in the page; `Query` being described as read-only does not enforce isolation. Prompt instructions and a relay allowlist alone cannot sandbox arbitrary model code.
