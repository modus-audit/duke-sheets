# Modus browser WASM distribution

Version `0.1.23-modus.1`, built from [source `e7832952e19b2b4016df292aa24aa2e80d3209b1`](https://github.com/modus-audit/duke-sheets/commit/e7832952e19b2b4016df292aa24aa2e80d3209b1) on upstream `wasm-v0.1.23`.

Preserves the Modus accounting formats 41–44 and Excel-compatible floating-point cancellation. The upstream release supplies the external add-in callback and cached-value handling. `SOURCE.patch` contains the complete changes from upstream.

Build from the linked source revision with:

```sh
wasm-pack build bindings/wasm --target web --out-dir pkg --release
node bindings/wasm/package-modus.mjs
```

The package manifest records the source revision. Native formula and number-format regression tests live in the source repository; browser integration tests live in `peasebell/e2e/tests/xlsx`.
