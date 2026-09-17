# Modus viewer engine

Based on upstream `wasm-v0.1.23` (`b30fff867622703f71edea6be16f20a2d4f96193`).

The Modus patches preserve built-in accounting formats 41–44 in both format registries and cancel floating-point reconciliation residue without rounding real differences or exact integer subtraction. External add-in callbacks and cached unresolved values now come from upstream.

Validate with `cargo test -p ssfmt -p duke-sheets-core` and `cargo test -p duke-sheets --test formula_evaluation`.

Build the browser package with `wasm-pack build bindings/wasm --target web --out-dir pkg --release`, then run `node bindings/wasm/package-modus.mjs` from a committed source revision. Copy the contents of `bindings/wasm/pkg` into a distribution branch and tag it `wasm-dist-0.1.23-modus.1`. The generated manifest records the exact source commit. Never move an existing distribution tag.

The application-level regression suite in `peasebell/e2e/tests/xlsx` exercises the packaged WASM through both the main-thread and worker viewer paths.
