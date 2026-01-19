/**
 * Renderer polyfills.
 *
 * Some dependencies (and/or webpack runtime in certain modes) may expect a Node-like `global`.
 * In Electron renderer with nodeIntegration disabled, `global` is not defined by default.
 */

// eslint-disable-next-line @typescript-eslint/no-explicit-any
;(globalThis as any).global = globalThis;
