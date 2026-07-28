# @dukelib/sheets-wasm — Modus fork (built artifact)

This orphan branch holds the **built** WASM package consumed by peasebell as a git dependency.
It adds an `externalFnFn(name, args) -> value | null` host callback so the engine resolves
`[N]!FN(args)` CCH add-in calls (TBLink/CLIENTNAME/…) during `calculate()`, keeping formula text
pure. With no callback (or a declined call) it returns the cell's cached value (strict superset).

Source of the Modus-specific change: `SOURCE.patch`, generated from `main...modus-tb`. Rebuild with
`wasm-pack build --release --target web` in `bindings/wasm`. Consumed via tag `wasm-dist-<ver>`.
