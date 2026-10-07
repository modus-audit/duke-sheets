# Modus browser WASM distribution

Version `0.1.23-modus.2`, built from [source `c1a54fff42584741c9bab33f563d1d7bc048b135`](https://github.com/modus-audit/duke-sheets/commit/c1a54fff42584741c9bab33f563d1d7bc048b135) on upstream `wasm-v0.1.23`.

Preserves the Modus accounting formats 41–44 and Excel-compatible floating-point cancellation. Built without the engine's `parallel` feature, so workbooks with 5,000 or more formulas calculate serially instead of trapping on a thread pool the browser can't provide. The upstream release supplies the external add-in callback and cached-value handling. `SOURCE.patch` contains the complete changes from upstream.

Build from the linked source revision with:

```sh
wasm-pack build bindings/wasm --target web --out-dir pkg --release
node bindings/wasm/package-modus.mjs
```

The package manifest records the source revision. Native formula and number-format regression tests live in the source repository; browser integration tests live in `peasebell/e2e/tests/xlsx`.
